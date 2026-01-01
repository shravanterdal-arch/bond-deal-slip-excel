import streamlit as st
import pdfplumber
import pandas as pd
import re
from io import BytesIO

# -------------------- PAGE SETUP --------------------
st.set_page_config(page_title="Bond Deal Slip → Excel", layout="centered")
st.title("📄 Bond Deal Slip → 📊 Excel")
st.caption("Supports mixed BSE (NDS-RST) and CBRICS deal slips")

# -------------------- HELPERS --------------------
def grab(pattern, text):
    m = re.search(pattern, text, re.DOTALL)
    return m.group(1).strip() if m else ""

def safe_float(val):
    try:
        return float(val.replace(",", ""))
    except:
        return ""

def safe_int(val):
    try:
        return int(val)
    except:
        return ""

# -------------------- BSE PARSER --------------------
def parse_bse(text):
    trade_value = safe_float(grab(r"TRADE VALUE\s+([\d.]+)", text))
    quantity = safe_int(grab(r"QUANTITY\s+(\d+)", text))

    fv = ""
    if trade_value and quantity:
        fv = round(trade_value / quantity, 2)

    return {
        "Deal Reference": grab(r"DEAL ID\s+(\S+)", text),
        "Buyer": grab(r"BUYER\s+(.+)", text),
        "Seller": grab(r"SELLER\s+(.+)", text),
        "Bond": grab(r"ISSUER NAME\s+(.+)", text),
        "ISIN": grab(r"ISIN\s+(\S+)", text),
        "Quantity": quantity,
        "FV per unit": fv,
        "Price": safe_float(grab(r"PRICE\s+([\d.]+)", text)),
        "SELLER CONSIDERATION": safe_float(grab(r"SELLER CONSIDERATION\s+([\d.]+)", text)),
        "BUYER CONSIDERATION": safe_float(grab(r"BUYER CONSIDERATION\s+([\d.]+)", text)),
        "YIELD(%)": safe_float(grab(r"YIELD\(%\)\s+([\d.]+)", text)),
    }

# -------------------- CBRICS PARSER (FIXED) --------------------
def parse_cbrics(text):
    lines = [l.strip() for l in text.split("\n") if l.strip()]

    def value_after(label):
        for i, l in enumerate(lines):
            if label.lower() in l.lower():
                if i + 1 < len(lines):
                    return lines[i + 1]
        return ""

    return {
        "Deal Reference": grab(r"CBRICS Transaction Id\s+(\d+)", text),

        "Buyer": grab(r"Participant\s+([A-Z0-9]+)", text),
        "Seller": grab(r"Counter Party\s+([A-Z0-9]+)", text),

        "Bond": value_after("Description"),
        "ISIN": grab(r"ISIN\s+(\S+)", text),

        "Quantity": safe_int(value_after("No. Of Bond")),
        "FV per unit": "",

        "Price": safe_float(grab(r"Price\s+([\d.]+)", text)),

        "SELLER CONSIDERATION": safe_float(
            grab(r"Actual Consideration\s+([\d,]+\.\d+)", text)
        ),

        "BUYER CONSIDERATION": safe_float(
            grab(r"Consideration Reported\s*\(incl\. Stamp Duty\)\s*([\d,]+\.\d+)", text)
        ),

        "YIELD(%)": safe_float(grab(r"Yield\s+([\d.]+)", text)),
    }

# -------------------- UI --------------------
uploaded_files = st.file_uploader(
    "Upload deal slip PDFs (BSE + CBRICS mixed)",
    type=["pdf"],
    accept_multiple_files=True
)

if uploaded_files:
    st.write(f"📂 {len(uploaded_files)} files ready")

    col1, col2 = st.columns(2)
    generate = col1.button("Generate Excel")
    clear = col2.button("Clear All")

    if clear:
        st.experimental_rerun()

    if generate:
        rows = []

        for pdf in uploaded_files:
            with pdfplumber.open(pdf) as p:
                text = "\n".join(
                    page.extract_text()
                    for page in p.pages
                    if page.extract_text()
                )

            if "CBRICS - CORPORATE BOND REPORTING" in text:
                rows.append(parse_cbrics(text))
            else:
                rows.append(parse_bse(text))

        df = pd.DataFrame(rows)

        output = BytesIO()
        df.to_excel(output, index=False, engine="openpyxl")
        output.seek(0)

        st.success("Excel generated successfully")

        st.download_button(
            "⬇️ Download Excel",
            data=output,
            file_name="Bond_Deal_Slips.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
