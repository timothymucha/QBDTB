# dtb_iif_converter.py

import streamlit as st
import pandas as pd
from io import StringIO
import re

# === Payee & Memo extractor ===
def clean_transaction_details(details):
    match = re.match(r".*\|(.*?)\|(.*?)\|(.*?)\s+-?\d+.*", str(details))
    if match:
        payee = match.group(1).strip()
        memo = match.group(3).strip()
    else:
        payee = str(details).split("|")[0].strip()[:30]
        memo = ""
    return payee or "Unknown", memo or "DTB Transaction"

# === IIF Generator ===
def generate_iif(df):
    output = StringIO()
    output.write("!TRNS\tTRNSTYPE\tDATE\tACCNT\tNAME\tAMOUNT\tMEMO\tDOCNUM\tCLEAR\n")
    output.write("!SPL\tTRNSTYPE\tDATE\tACCNT\tNAME\tAMOUNT\tMEMO\tDOCNUM\tCLEAR\n")
    output.write("!ENDTRNS\n")

    for _, row in df.iterrows():
        trn_type = str(row.get('Transaction Type')).strip()
        if trn_type == "MPESA FUNDS TRANSFER":
            continue  # Skip these

        try:
            date = pd.to_datetime(row['Transaction Date']).strftime('%m/%d/%Y')
        except:
            continue

        sheet['Credits'] = pd.to_numeric(sheet['Credits'].astype(str).str.replace(',', ''), errors='coerce').fillna(0.0)
        sheet['Debits'] = pd.to_numeric(sheet['Debits'].astype(str).str.replace(',', ''), errors='coerce').fillna(0.0)

        docnum = str(row.get('Reference', ''))
        details = str(row.get('Transaction Details', ''))
        debit = float(row.get('Debits') or 0)
        credit = float(row.get('Credits') or 0)

        payee, memo = clean_transaction_details(details)

        # Charges
        charges = float(row.get('Charges') or 0)
        commission = float(row.get('Commission Amount') or 0)

        # Pesalink payments to vendors
        if trn_type == "PESA LINK TRANSACTION" and debit > 0:
            output.write(f"TRNS\tCHECK\t{date}\tDiamond Trust Bank\t{payee}\t{-debit:.2f}\t{memo}\t{docnum}\tN\n")
            output.write(f"SPL\tCHECK\t{date}\tAccounts Payable\t{payee}\t{debit:.2f}\t{memo}\t{docnum}\tN\n")
            output.write("ENDTRNS\n")

        # Unexpected money in — go to suspense
        elif credit > 0:
            output.write(f"TRNS\tDEPOSIT\t{date}\tDiamond Trust Bank\t{payee}\t{credit:.2f}\t{memo}\t{docnum}\tN\n")
            output.write(f"SPL\tDEPOSIT\t{date}\tAsk My Accountant\t{payee}\t{-credit:.2f}\t{memo}\t{docnum}\tN\n")
            output.write("ENDTRNS\n")

        # Bank Charges & Commissions
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
        st.info("Make sure it's a valid `.xls` file and `xlrd==1.2.0` is installed.")
