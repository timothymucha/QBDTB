# dtb_to_iif_streamlit.py
import re
import pandas as pd
from io import StringIO
import streamlit as st
from fuzzywuzzy import fuzz

# ==========================
# Supplier / Vendor list
# ==========================
VENDORS = [
    "254 Brewing Company", "A.S.W Enterprises Limited", "A.W Water Boozer Services",
    "AAA Growers LTD", "Alyemda Enterprise Ltd", "Araali Limited", "Assmazab General Stores",
    "Baraka Israel Enterprises Limited", "Benchmark Distributors Limited", "Best Buy Distributors",
    "Beyond Fruits Limited", "Bio Food Products Limited", "Boos Ventures", "Bowip Agencies Ltd",
    "Branded Fine Foods Ltd", "Brookside Dairy Ltd", "Brown Bags", "Cafesserie Bread Store",
    "Casks And Barrels Ltd", "Chandaria Industries Ltd", "CHIRAG AFRICA LIMITED",
    "Coastal Bottlers Limited", "Crystal Frozen & Chilled Foods Ltd", "De Vries Africa Ventures",
    "Debenham & Fear Ltd", "Dekow Wholesale", "Deliveries", "Diamond Trust Bank", "Dilawers",
    "Dion Wine And Spirits East Africa Limited", "Disney Wines & Spirits", "Domaine Kenya Ltd",
    "Dormans Coffee Ltd", "Eco-Essentials Limited", "Ewca Marketing (Kenya) Limited",
    "Exotics Africa Ltd", "Express Shop Bofa", "Ezzi Traders Limited", "Farmers Choice Limited",
    "Fayaz Bakers Ltd", "Finsbury Trading Ltd", "Fratres Malindi", "FRAWAROSE LIMITED",
    "Galina Agencies", "Gilani's Distributors LTD", "Glacier Products Ltd",
    "Global Slacker Enterprises Ltd", "Handas Juice Ltd", "Hasbah Kenya Limited",
    "Healthy U Two Thousand Ltd", "HOME BEST HEALTH FOOD LIMITED", "House of Booch Ltd",
    "Ice Hub Limited", "Isinya Feeds Limited", "Jetlak Limited", "Kalon Foods Limited",
    "Karen Fork", "Kenchic Limited", "Kenya Commercial Bank", "Kenya Nut Company",
    "Kenya Power and Lighting Company", "Kenya Revenue Authority", "Khosal Wholesale Kilifi",
    "Kioko Enterprises", "Lakhani General Suppliers Lilmited", "Laki Laki Ltd", "Landlord",
    "LEXO ENERGY KENYA LIMITED", "Lindas Nut Butter", "Linkbizz E-Hub Commerce Ltd",
    "Loki Ventures Limited", "Malachite Limited", "Malindi Industries Limited", "Mill Bakers",
    "Mini Bakeries (NRB) ltd", "Mjengo Limited", "Mnarani Pens", "Mnarani Water Refil",
    "MohanS Oysterbay Drinks K Ltd", "Moonsun Picture International Limited",
    "Mudee Concepts Limited", "Mwanza Kambi Tsuma", "Mzuri Sweets Limited",
    "Naaman Muses & Co. Ltd", "Nairobi Java House Limited", "Naji Superstores",
    "National Social Security Fund", "Neema Stores Kilifi", "Njugu Supplier",
    "Nyali Air Conditioning & Refrigeration Se", "Pasagot Limited", "Plastic Cups",
    "Pride Industries Ltd", "Radbone-Clark Kenya limited", "Raisons Distributors Ltd",
    "Rehema Jumwa Muli", "RK'S Products", "S.H.I.F Payable", "Safaricom",
    "Savannah Brands Company Ltd", "SEA HARVEST (K) LTD", "Shiva Mombasa Limited",
    "SIDR Distributors Limited", "Slater", "Sliquor Limited", "Social Health Insurance Fund",
    "Soko (Market)", "Sol O Vino Limited", "South Lemon LTD", "Soy's Limited",
    "Sun Power Products Limited", "Supreme Filing Station", "Takataka",
    "Tandaa Networks Limited", "Taraji", "Tawakal Store Company Ltd",
    "The Happy Lamb Butchery", "The Standard Group Plc", "Thomas Mwachongo Mwangala",
    "Three Spears Limited", "TOP IT UP DISTRIBUTOR", "Towfiq Kenya Limited",
    "Traderoots Limited", "Under the Influence East Africa", "UvA Wines",
    "VEGAN WORLD LIMITED", "Moha Eggs", "Water Refil", "Wine and More Limited",
    "Wingu Box Ltd", "Zabach Enterprises Limited", "Zen Mahitaji Ltd", "Zenko Kenya Limited",
    "Zuri Central"
]

STOPWORDS = {
    "ltd", "limited", "enterprises", "company", "kenya", "plc", "group",
    "east", "africa", "distributors", "distributor", "trading", "suppliers", "supplier"
}

# ==========================
# Staff list & tokens (match any name-part)
# ==========================
STAFF = [
    "Grace Sanita Ngumbao",
    "Robert Githinji Mwangi",
    "Brian Mavuro Ngetsa",
    "Rehema Jumwa Muli",
    "Timothy Wafula Mucha",
    "Director John Ojal",
    "Director Esther Muthumbi"
]

STAFF_ALIAS_MAP = {}
for full in STAFF:
    for part in full.split():
        STAFF_ALIAS_MAP[part.lower()] = full

# ==========================
# Text helpers
# ==========================
def norm(s: str) -> str:
    return re.sub(r"[^a-z0-9]+", " ", str(s).strip().lower()).strip()

def tokens(s: str):
    return [t for t in norm(s).split() if t and t not in STOPWORDS]

def strip_stopwords(s: str):
    return " ".join(tokens(s))

def clean_memo(s: str) -> str:
    return str(s).replace('"', "").replace("\n", " ").strip()

# ==========================
# Build alias map for vendors (unique tokens -> vendor)
# ==========================
def build_alias_map(vendors):
    alias = {}
    for v in vendors:
        for t in set(tokens(v)):
            alias.setdefault(t, set()).add(v)

    out = {}
    for k, vs in alias.items():
        if len(vs) == 1:
            out[k] = list(vs)[0]

    # manual hints (common keywords)
    out.update({
        "brookside": "Brookside Dairy Ltd",
        "benchmark": "Benchmark Distributors Limited",
        "cigarettes": "Benchmark Distributors Limited",
        "best": "Best Buy Distributors",
        "buy": "Best Buy Distributors",
        "sun": "Sun Power Products Limited",
        "power": "Sun Power Products Limited",
        "malachite": "Malachite Limited",
        "glacier": "Glacier Products Ltd",
        "domain": "Domaine Kenya Ltd",
        "dormans": "Dormans Coffee Ltd",
        "254": "254 Brewing Company",
        "crystal": "Crystal Frozen & Chilled Foods Ltd",
        "raisons": "Raisons Distributors Ltd",
        "sidr": "SIDR Distributors Limited",
        "coke": "Coastal Bottlers Limited",
        "booch": "House of Booch Ltd",
        "bio": "Bio Food Products Limited",
        # Michael Simon Mwanyinge tokens mapped to Brookside
        "michael": "Brookside Dairy Ltd",
        "simon": "Brookside Dairy Ltd",
        "mwanyinge": "Brookside Dairy Ltd",
    })
    return out

ALIAS_MAP = build_alias_map(VENDORS)

# ==========================
# Supplier/staff matching (staff prioritized)
# ==========================
def match_supplier(detail: str, suppliers: list[str], threshold: int = 86) -> str | None:
    if not detail:
        return None

    detail_tokens = set(tokens(detail))

    # 1) staff check (any token -> staff)
    for t in detail_tokens:
        if t in STAFF_ALIAS_MAP:
            return STAFF_ALIAS_MAP[t]

    # 2) alias unique-token check
    for t in detail_tokens:
        if t in ALIAS_MAP:
            return ALIAS_MAP[t]

    # 3) fuzzy fallback against full vendor names
    stripped_detail = strip_stopwords(detail)
    best, best_score = None, -1
    for s in suppliers:
        score = fuzz.token_set_ratio(stripped_detail, strip_stopwords(s))
        if score > best_score:
            best_score, best = score, s
    if best_score >= threshold:
        return best

    return None

# ==========================
# Clean & extract payee + memo
# ==========================
def clean_transaction_details(details: str, threshold: int = 86):
    parts = str(details).split("|")
    if len(parts) >= 4:
        raw_payee = parts[1].strip()
        memo = parts[3].strip()
    elif len(parts) >= 2:
        raw_payee = parts[1].strip()
        memo = ""
    else:
        raw_payee = str(details).strip()[:60]
        memo = ""

    # Try to match on the full details first, else on raw_payee
    matched = match_supplier(details, VENDORS, threshold=threshold) or match_supplier(raw_payee, VENDORS, threshold=threshold)
    payee = matched if matched else (raw_payee or "General Supplier")
    return payee, clean_memo(memo or details or "DTB Transaction")

# ==========================
# Date helper
# ==========================
def qb_date(x):
    try:
        dt = pd.to_datetime(x, errors="coerce", dayfirst=False)
        if pd.isna(dt):
            return None
        return dt.strftime("%m/%d/%Y")
    except Exception:
        return None

# ==========================
# IIF Generator
# ==========================
def generate_iif(df: pd.DataFrame, threshold: int = 86) -> str:
    # Ensure necessary columns exist and convert to numeric
    for col in ['Credits', 'Debits', 'Charges', 'Commission Amount']:
        if col not in df.columns:
            df[col] = 0
        # remove commas and coerce
        df[col] = pd.to_numeric(df[col].astype(str).str.replace(',', ''), errors='coerce').fillna(0.0)

    # Ensure text columns exist
    for col in ['Transaction Type', 'Transaction Date', 'Reference', 'Transaction Details']:
        if col not in df.columns:
            df[col] = ""

    out = StringIO()
    out.write("!TRNS\tTRNSTYPE\tDATE\tACCNT\tNAME\tAMOUNT\tMEMO\tDOCNUM\tCLEAR\n")
    out.write("!SPL\tTRNSTYPE\tDATE\tACCNT\tNAME\tAMOUNT\tMEMO\tDOCNUM\tCLEAR\n")
    out.write("!ENDTRNS\n")

    bank_charge_types = {
        "PESA LINK TXN CHG", "EXCISE DUTY", "MOBILE BANKING TXN CHARGE",
        "I24/7 TXN CHARGE", "CHEQUE BOOK CHARGES"
    }
    ask_my_accountant_types = {"MOBILE BANKING TXN"}

    for _, row in df.iterrows():
        trn_type = str(row.get('Transaction Type', '')).strip().upper()

        # skip mpesa funds transfer lines
        if trn_type == "MPESA FUNDS TRANSFER":
            continue

        date_raw = row.get('Transaction Date', '')
        date_str = qb_date(date_raw)
        if not date_str:
            # skip transactions without a parsable date
            continue

        raw_ref = str(row.get('Reference', '') or "").strip()
        docnum = raw_ref[-9:] if len(raw_ref) > 9 else raw_ref or ""

        details = str(row.get('Transaction Details', '') or "")
        debit = float(row.get('Debits') or 0.0)
        credit = float(row.get('Credits') or 0.0)
        charges = float(row.get('Charges') or 0.0)
        commission = float(row.get('Commission Amount') or 0.0)

        payee, memo = clean_transaction_details(details, threshold=threshold)

        # Bank charge types (explicit)
        if trn_type in bank_charge_types:
            amount = debit if debit > 0 else (charges + commission)
            fee_memo = f"{trn_type} - {payee}"
            out.write(f"TRNS\tCHECK\t{date_str}\tDiamond Trust Bank\tBank Charges DTB\t{-amount:.2f}\t{clean_memo(fee_memo)}\t{docnum}\tN\n")
            out.write(f"SPL\tCHECK\t{date_str}\tBank Service Charges:Bank Charges - DTB\t\t{amount:.2f}\t{clean_memo(fee_memo)}\t{docnum}\tN\n")
            out.write("ENDTRNS\n")
            continue

        # Ask accountant type
        if trn_type in ask_my_accountant_types:
            amount = debit if debit > 0 else credit
            out.write(f"TRNS\tCHECK\t{date_str}\tDiamond Trust Bank\t{payee}\t{-amount:.2f}\t{memo}\t{docnum}\tN\n")
            out.write(f"SPL\tCHECK\t{date_str}\tAccounts Payable\t{payee}\t{amount:.2f}\t{memo}\t{docnum}\tN\n")
            out.write("ENDTRNS\n")
            continue

        # Debit customer/vendor payments (common patterns)
        if (trn_type == "MOBILE BANKING FT TXN" and debit > 0) or \
           (trn_type == "PESA LINK TRANSACTION" and debit > 0) or \
           (trn_type == "IN-HOUSE CHEQUE" and debit > 0) or \
           (trn_type == "INWARD CLEARING" and debit > 0):
            out.write(f"TRNS\tCHECK\t{date_str}\tDiamond Trust Bank\t{payee}\t{-debit:.2f}\t{memo}\t{docnum}\tN\n")
            out.write(f"SPL\tCHECK\t{date_str}\tAccounts Payable\t{payee}\t{debit:.2f}\t{memo}\t{docnum}\tN\n")
            out.write("ENDTRNS\n")
            continue

        # Credits (incoming transfers)
        if credit > 0:
            out.write(f"TRNS\tTRANSFER\t{date_str}\tDiamond Trust Bank\t{payee}\t{credit:.2f}\t{memo}\t{docnum}\tN\n")
            out.write(f"SPL\tTRANSFER\t{date_str}\tExpress Bofa\t{payee}\t{-credit:.2f}\t{memo}\t{docnum}\tN\n")
            out.write("ENDTRNS\n")
            continue

        # Bank fees if any leftover
        if charges > 0 or commission > 0:
            total_fees = charges + commission
            fee_memo = f"Bank charges & commissions - {payee}"
            out.write(f"TRNS\tCHECK\t{date_str}\tDiamond Trust Bank\tBank Charges DTB\t{-total_fees:.2f}\t{clean_memo(fee_memo)}\t{docnum}\tN\n")
            out.write(f"SPL\tCHECK\t{date_str}\tBank Service Charges:Bank Charges - DTB\t\t{total_fees:.2f}\t{clean_memo(fee_memo)}\t{docnum}\tN\n")
            out.write("ENDTRNS\n")
            continue

        # If none of the above rules matched, optionally treat as check to General Supplier
        # (keeps behaviour forgiving rather than dropping rows)
        # treat zero-debit/credit transactions as no-op (skip)
        if debit > 0:
            out.write(f"TRNS\tCHECK\t{date_str}\tDiamond Trust Bank\t{payee}\t{-debit:.2f}\t{memo}\t{docnum}\tN\n")
            out.write(f"SPL\tCHECK\t{date_str}\tAccounts Payable\t{payee}\t{debit:.2f}\t{memo}\t{docnum}\tN\n")
            out.write("ENDTRNS\n")

    return out.getvalue()

# ==========================
# Streamlit UI
# ==========================
st.set_page_config(page_title="DTB -> QuickBooks IIF", layout="wide")
st.title("🏦 DTB → QuickBooks IIF Converter")

with st.sidebar:
    st.header("Settings")
    fuzzy_threshold = st.slider(
        "Fuzzy match threshold",
        min_value=60, max_value=98, value=86, step=1,
        help="Higher = stricter supplier fuzzy matching. Lower = more permissive."
    )
    st.markdown("Staff tokens are matched by any name part (first/middle/last).")

uploaded_file = st.file_uploader("Upload DTB file (.xls, .xlsx, .csv)", type=["xls", "xlsx", "csv"])

if uploaded_file:
    try:
        if uploaded_file.name.lower().endswith(".xls"):
            df = pd.read_excel(uploaded_file, skiprows=17, engine="xlrd")
        elif uploaded_file.name.lower().endswith(".xlsx"):
            df = pd.read_excel(uploaded_file, skiprows=17, engine="openpyxl")
        else:
            # csv
            df = pd.read_csv(uploaded_file)
    except Exception as e:
        st.error(f"Failed to read file: {e}")
        st.stop()

    st.subheader("Raw preview")
    st.dataframe(df.head(20), use_container_width=True)

    # Quick preview of matched suppliers/staff and unmatched examples
    unmatched = []
    suggested = []
    for _, r in df.iterrows():
        details = str(r.get("Transaction Details", "") or r.get("Details", "") or "")
        if not details:
            continue
        match = match_supplier(details, VENDORS, threshold=fuzzy_threshold)
        if match:
            suggested.append((details, match))
        else:
            unmatched.append(details)

    if unmatched:
        st.warning("⚠️ Unmatched vendor/staff examples (first 15):")
        st.write("\n".join(list(dict.fromkeys(unmatched[:15]))))
    if suggested:
        with st.expander("📎 Suggested matches (sample)"):
            samp = pd.DataFrame(suggested[:50], columns=["Detail", "Matched Vendor/Staff"])
            st.dataframe(samp, use_container_width=True)

    if st.button("Generate QuickBooks IIF"):
        try:
            iif_text = generate_iif(df, threshold=fuzzy_threshold)
            st.success("IIF generated.")
            st.download_button(
                "📥 Download petty_DTB.iif",
                data=iif_text.encode("utf-8"),
                file_name="DTB_output.iif",
                mime="text/plain"
            )
        except Exception as e:
            st.error(f"Failed to generate IIF: {e}")
else:
    st.info("Upload your DTB statement (.xls/.xlsx/.csv). Typical sheets need headers including: Transaction Type, Transaction Date, Reference, Transaction Details, Debits, Credits, Charges, Commission Amount.")
