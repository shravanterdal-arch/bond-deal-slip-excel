import streamlit as st
import pdfplumber
import pandas as pd
import re
from io import BytesIO

st.set_page_config(page_title="BSE Deal Slip → Excel", layout="centered")
st.title("📄 BSE Deal Slip → 📊 Excel")
st.caption("Upload BSE deal confirmation PDFs only")

# ---------------- HELPERS ----------------
def grab(pattern, text):
    m = re.search(pattern, text, re.IGNORECASE)
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

# ---------------- BSE PARSER ----------------
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
        "YIELD(%)": (
    grab(r"YIELD\(%\)\s+([\d.]+)", text) + "%"
    if grab(r"YIELD\(%\)\s+([\d.]+)", text)
    else ""
),

    }

# ---------------- UI ----------------
uploaded_files = st.file_uploader(
    "Upload BSE Deal Slip PDFs",
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
            rows.append(parse_bse(text))

        df = pd.DataFrame(rows)

        out = BytesIO()
        df.to_excel(out, index=False, engine="openpyxl")
        out.seek(0)

        st.download_button(
            "⬇️ Download Excel",
            data=out,
            file_name="BSE_Deal_Slips.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
