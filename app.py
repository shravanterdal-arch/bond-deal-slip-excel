import streamlit as st
import pdfplumber
import pandas as pd
import re
from io import BytesIO
from pdf2image import convert_from_bytes
import pytesseract

st.set_page_config(page_title="Bond Deal Slip → Excel", layout="centered")
st.title("📄 Bond Deal Slip → 📊 Excel")

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

def ocr_pdf(file_bytes):
    images = convert_from_bytes(file_bytes)
    text = ""
    for img in images:
        text += pytesseract.image_to_string(img)
    return text

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

# ---------------- CBRICS (OCR) ----------------
def parse_cbrics(file_bytes):
    text = ocr_pdf(file_bytes)

    return {
        "Deal Reference": grab(r"CBRICS Transaction Id\s*[:\-]?\s*(\d+)", text),
        "Buyer": grab(r"Participant\s*[:\-]?\s*([A-Z ]+)", text),
        "Seller": grab(r"Counter Party\s*[:\-]?\s*([A-Z ]+)", text),
        "Bond": grab(r"Description\s*[:\-]?\s*(.+)", text),
        "ISIN": grab(r"ISIN\s*[:\-]?\s*(\S+)", text),
        "Quantity": to_int(grab(r"No\.?\s*Of\s*Bond\s*[:\-]?\s*(\d+)", text)),
        "FV per unit": "",
        "Price": to_float(grab(r"Price\s*[:\-]?\s*([\d.]+)", text)),
        "SELLER CONSIDERATION": to_float(
            grab(r"Actual Consideration\s*[:\-]?\s*([\d,\.]+)", text)
        ),
        "BUYER CONSIDERATION": to_float(
            grab(r"Consideration Reported.*?\s([\d,\.]+)", text)
        ),
        "YIELD(%)": to_float(grab(r"Yield\s*[:\-]?\s*([\d.]+)", text)),
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
            file_bytes = f.read()

            with pdfplumber.open(BytesIO(file_bytes)) as pdf:
                text = "\n".join(page.extract_text() or "" for page in pdf.pages)

            if "CBRICS" in text.upper():
                rows.append(parse_cbrics(file_bytes))
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
