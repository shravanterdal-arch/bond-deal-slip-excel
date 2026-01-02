import streamlit as st
import pdfplumber
import pandas as pd
import re
from io import BytesIO

# ---------------- PAGE ----------------
st.set_page_config(page_title="Bond Deal Slip → Excel", layout="centered")
st.title("📄 Bond Deal Slip → 📊 Excel")

# ---------------- HELPERS ----------------
def grab(pattern, text):
    m = re.search(pattern, text, re.DOTALL)
    return m.group(1).strip() if m else ""

def to_float(x):
    try:
        return float(x.replace(",", ""))
    except:
        return ""

def to_int(x):
    try:
        return int(x)
    except:
        return ""

# ---------------- BSE ----------------
def parse_bse(text):
    trade_value = to_float(grab(r"TRADE VALUE\s+([\d.]+)", text))
    qty = to_int(grab(r"QUANTITY\s+(\d+)", text))

    fv = ""
    if trade_value and qty:
        fv = round(trade_value / qty, 2)

    return {
        "Deal Reference": grab(r"DEAL ID\s+(\S+)", text),
        "Buyer": grab(r"BUYER\s+(.+)", text),
        "Seller": grab(r"SELLER\s+(.+)", text),
        "Bond": grab(r"ISSUER NAME\s+(.+)", text),
        "ISIN": grab(r"ISIN\s+(\S+)", text),
        "Quantity": qty,
        "FV per unit": fv,
        "Price": to_float(grab(r"PRICE\s+([\d.]+)", text)),
        "SELLER CONSIDERATION": to_float(grab(r"SELLER CONSIDERATION\s+([\d.]+)", text)),
        "BUYER CONSIDERATION": to_float(grab(r"BUYER CONSIDERATION\s+([\d.]+)", text)),
        "YIELD(%)": to_float(grab(r"YIELD\(%\)\s+([\d.]+)", text)),
    }

# ---------------- CBRICS (TABLE-BASED) ----------------
def parse_cbrics(pdf):
    text = "\n".join(
        page.extract_text() or "" for page in pdf.pages
    )

    tables = []
    for page in pdf.pages:
        try:
            tables += page.extract_tables()
        except:
            pass

    # Defaults
    buyer = seller = bond = isin = ""
    qty = price = ytm = buyer_cons = seller_cons = ""
    deal_ref = grab(r"CBRICS Transaction Id\s+(\d+)", text)

    # Text-safe fields
    isin = grab(r"ISIN\s+(\S+)", text)
    bond = grab(r"Description\s+(.+)", text)
    buyer = grab(r"Participant\s+([A-Z0-9]+)", text)
    seller = grab(r"Counter Party\s+([A-Z0-9]+)", text)
    price = to_float(grab(r"Price\s+([\d.]+)", text))
    ytm = to_float(grab(r"Yield\s+([\d.]+)", text))

    # Table-driven fields (reliable)
    for table in tables:
        for row in table:
            row_text = " ".join([c or "" for c in row]).lower()

            if "no. of bond" in row_text:
                qty = to_int(row[0]) or to_int(row[-1])

            if "consideration reported" in row_text:
                buyer_cons = to_float(row[-1])

            if "actual consideration" in row_text:
                seller_cons = to_float(row[-1])

    return {
        "Deal Reference": deal_ref,
        "Buyer": buyer,
        "Seller": seller,
        "Bond": bond,
        "ISIN": isin,
        "Quantity": qty,
        "FV per unit": "",
        "Price": price,
        "SELLER CONSIDERATION": seller_cons,
        "BUYER CONSIDERATION": buyer_cons,
        "YIELD(%)": ytm,
    }

# ---------------- UI ----------------
uploaded_files = st.file_uploader(
    "Upload deal slip PDFs (BSE + CBRICS mixed)",
    type=["pdf"],
    accept_multiple_files=True
)

if uploaded_files:
    if st.button("Generate Excel"):
        rows = []

        for f in uploaded_files:
            with pdfplumber.open(f) as pdf:
                text = "\n".join(
                    page.extract_text() or "" for page in pdf.pages
                )

                if "CBRICS - CORPORATE BOND REPORTING" in text:
                    rows.append(parse_cbrics(pdf))
                else:
                    rows.append(parse_bse(text))

        df = pd.DataFrame(rows)

        out = BytesIO()
        df.to_excel(out, index=False, engine="openpyxl")
        out.seek(0)

        st.download_button(
            "⬇️ Download Excel",
            data=out,
            file_name="Bond_Deal_Slips.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
