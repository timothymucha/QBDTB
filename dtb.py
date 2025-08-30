import streamlit as st
import pandas as pd
from io import StringIO
import re

# === Payee & Memo extractor ===
def clean_transaction_details(details):
    parts = str(details).split("|")
    if len(parts) >= 4:
        # Structured DTB format
        payee = parts[1].strip()           # always 2nd element
        memo = parts[3].strip()            # 4th element
    elif len(parts) >= 2:
        # If only 2 fields, use second one
        payee = parts[1].strip()
        memo = ""
    else:
        # No separators
        payee = str(details).strip()[:30]
        memo = ""
    return payee or "Unknown", memo or "DTB Transaction"


# === IIF Generator ===
def generate_iif(df):
    # Ensure necessary columns exist and convert to numeric
    for col in ['Credits', 'Debits', 'Charges', 'Commission Amount']:
        if col not in df.columns:
            df[col] = 0
        df[col] = pd.to_numeric(df[col].astype(str).str.replace(',', ''), errors='coerce').fillna(0.0)

    output = StringIO()
    output.write("!TRNS\tTRNSTYPE\tDATE\tACCNT\tNAME\tAMOUNT\tMEMO\tDOCNUM\tCLEAR\n")
    output.write("!SPL\tTRNSTYPE\tDATE\tACCNT\tNAME\tAMOUNT\tMEMO\tDOCNUM\tCLEAR\n")
    output.write("!ENDTRNS\n")

    bank_charge_types = {"PESA LINK TXN CHG", "EXCISE DUTY", "MOBILE BANKING TXN CHARGE", "I24/7 TXN CHARGE", "CHEQUE BOOK CHARGES"}
    ask_my_accountant_types = {"MOBILE BANKING TXN"}

    for _, row in df.iterrows():
        trn_type = str(row.get('Transaction Type')).strip().upper()

        # Skip MPESA FUNDS TRANSFER
        if trn_type == "MPESA FUNDS TRANSFER":
            continue

        try:
            date = pd.to_datetime(row['Transaction Date']).strftime('%m/%d/%Y')
        except Exception:
            continue

        raw_ref = str(row.get('Reference', '')).strip()
        docnum = raw_ref[-9:] if len(raw_ref) > 9 else raw_ref

        details = str(row.get('Transaction Details', ''))
        debit = float(row.get('Debits') or 0)
        credit = float(row.get('Credits') or 0)

        payee, memo = clean_transaction_details(details)
        charges = float(row.get('Charges') or 0)
        commission = float(row.get('Commission Amount') or 0)

        if trn_type in bank_charge_types:
            amount = debit if debit > 0 else (charges + commission)
            fee_memo = f"{trn_type} - {payee}"
            output.write(f"TRNS\tCHECK\t{date}\tDiamond Trust Bank\tBank Charges DTB\t{-amount:.2f}\t{fee_memo}\t{docnum}\tN\n")
            output.write(f"SPL\tCHECK\t{date}\tBank Service Charges:Bank Charges - DTB\t\t{amount:.2f}\t{fee_memo}\t{docnum}\tN\n")
            output.write("ENDTRNS\n")
            continue

        if trn_type in ask_my_accountant_types:
            amount = debit if debit > 0 else credit
            output.write(f"TRNS\tCHECK\t{date}\tDiamond Trust Bank\t{payee}\t{-amount:.2f}\t{memo}\t{docnum}\tN\n")
            output.write(f"SPL\tCHECK\t{date}\tAccounts Payable\t{payee}\t{amount:.2f}\t{memo}\t{docnum}\tN\n")
            output.write("ENDTRNS\n")
            continue

        if (trn_type == "MOBILE BANKING FT TXN" and debit > 0)  or (trn_type == "PESA LINK TRANSACTION" and debit > 0) or (trn_type == "IN-HOUSE CHEQUE" and debit > 0) or (trn_type == "INWARD CLEARING" and debit > 0):
            output.write(f"TRNS\tCHECK\t{date}\tDiamond Trust Bank\t{payee}\t{-debit:.2f}\t{memo}\t{docnum}\tN\n")
            output.write(f"SPL\tCHECK\t{date}\tAccounts Payable\t{payee}\t{debit:.2f}\t{memo}\t{docnum}\tN\n")
            output.write("ENDTRNS\n")
            continue

        if (credit > 0) or (trn_type == "i24/7 PESALINK" and credit > 0):
            output.write(f"TRNS\tTRANSFER\t{date}\tDiamond Trust Bank\t{payee}\t{credit:.2f}\t{memo}\t{docnum}\tN\n")
            output.write(f"SPL\tTRANSFER\t{date}\tExpress Bofa\t{payee}\t{-credit:.2f}\t{memo}\t{docnum}\tN\n")
            output.write("ENDTRNS\n")
            continue

        if charges > 0 or commission > 0:
            total_fees = charges + commission
            fee_memo = f"Bank charges & commissions - {payee}"
            output.write(f"TRNS\tCHECK\t{date}\tDiamond Trust Bank\tBank Charges DTB\t{-total_fees:.2f}\t{fee_memo}\t{docnum}\tN\n")
            output.write(f"SPL\tCHECK\t{date}\tBank Service Charges:Bank Charges - DTB\t\t{total_fees:.2f}\t{fee_memo}\t{docnum}\tN\n")
            output.write("ENDTRNS\n")

    return output.getvalue()

# === Streamlit UI ===
st.set_page_config(page_title="DTB to QuickBooks IIF Converter", layout="centered")
st.title("🏦 DTB to QuickBooks IIF Converter")

uploaded_file = st.file_uploader("📤 Upload DTB Excel File (.xls)", type=["xls"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file, skiprows=17, engine="xlrd")
        st.success("✅ File successfully loaded.")
        st.write("### 🔍 Preview of Data", df.head())

        iif_data = generate_iif(df)
        st.download_button("📥 Download IIF File", data=iif_data, file_name="DTB_output.iif", mime="text/plain")
    except Exception as e:
        st.error(f"❌ Failed to read file: {e}")
        st.info("Have a Good Day.")
