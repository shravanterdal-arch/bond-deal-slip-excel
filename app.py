import streamlit as st
import pdfplumber
import pandas as pd
import re
from io import BytesIO
from docx import Document

# ---------------- PAGE ----------------
st.set_page_config(page_title="Bond Deal Slip → Excel", layout="centered")
st.title("📄 Bond Deal Slip → 📊 Excel")
st.caption("BSE (PDF) + CBRICS (Word) supported")

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

# ---------------- BSE (PDF) ----------------
def parse_bse_pdf(text):
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

# ---------------- CBRICS (WORD) ----------------
def parse_cbrics_docx(file):
    doc = Document(file)
    rows = {}

    for table in doc.tables:
        for row in table.rows:
            cells = [c.text.strip() for c in row.cells]
            if len(cells) < 2:
                continue

            key = cells[0].lower()
            val = cells[1]

            if "transaction id" in key:
                rows["Deal Reference"] = val
            elif "participant" in key:
                rows["Buyer"] = val
            elif "counter party" in key:
                rows["Seller"] = val
            elif "description" in key:
                rows["Bond"] = val
            elif "isin" in key:
                rows["ISIN"] = val
            elif "no. of bond" in key:
                rows["Quantity"] = to_int(val)
            elif "price" in key:
                rows["Price"] = to_float(val)
            elif "actual consideration" in key:
                rows["SELLER CONSIDERATION"] = to_float(val)
            elif "consideration reported" in key:
                rows["BUYER CONSIDERATION"] = to_float(val)
            elif "yield" in key:
                rows["YIELD(%)"] = to_float(val)

    rows["FV per unit"] = ""

    return rows

# ---------------- UI ----------------
uploaded_files = st.file_uploader(
    "Upload BSE PDFs and CBRICS Word files",
    type=["pdf", "docx"],
    accept_multiple_files=True
)

if uploaded_files:
    if st.button("Generate Excel"):
        rows = []

        for f in uploaded_files:
            if f.name.lower().endswith(".pdf"):
                with pdfplumber.open(f) as pdf:
                    text = "\n".join(
                        page.extract_text() or "" for page in pdf.pages
                    )
                rows.append(parse_bse_pdf(text))

            elif f.name.lower().endswith(".docx"):
                rows.append(parse_cbrics_docx(f))

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
