"""
Invoice → Excel Extractor (FREE / OCR version)
-----------------------------------------------
No API key, no cost — runs 100% locally using Tesseract OCR.

Upload one or more invoice files (PDF or image: jpg/png). The app reads
each invoice with OCR and does its best to pull out these fields into rows
of an Excel sheet:

Vendor Name | Invoice Number | Invoice Date | Vendor GSTIN |
Total Basic Amount | Total GST Amount | Other / Discount Amount |
Total Invoice Amount | Total Amount Paid against Invoice

IMPORTANT — this free version is heuristic (regex-based), not AI-powered.
Scanned invoices vary a lot in layout, so ALWAYS check the extracted table
before downloading. Each row also keeps the raw OCR text and the invoice's
"amount in words" lines so you can quickly fix anything OCR got wrong or
missed (numbers are usually the least reliable part of OCR).

Requirements (one-time setup):
  1. Install Tesseract OCR itself (this is separate from the pip package):
       Windows: https://github.com/UB-Mannheim/tesseract/wiki  (download & run installer)
       Mac:     brew install tesseract
       Linux:   sudo apt install tesseract-ocr
  2. pip install -r requirements.txt
  3. If Windows and Tesseract isn't on PATH, set TESSERACT_PATH below or as
     an environment variable, e.g.:
       set TESSERACT_PATH=C:\\Program Files\\Tesseract-OCR\\tesseract.exe

Run:
    streamlit run app.py
"""

import io
import os
import re

import fitz  # PyMuPDF
import pandas as pd
import pytesseract
import streamlit as st
from PIL import Image

st.set_page_config(page_title="Invoice → Excel Extractor (Free)", layout="wide")

# Point pytesseract at the Tesseract binary if it's not on PATH (Windows).
tesseract_path = os.environ.get("TESSERACT_PATH")
if tesseract_path:
    pytesseract.pytesseract.tesseract_cmd = tesseract_path
elif os.name == "nt":
    default_win_path = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
    if os.path.exists(default_win_path):
        pytesseract.pytesseract.tesseract_cmd = default_win_path

COLUMNS = [
    "Vendor Name",
    "Invoice Number",
    "Invoice Date",
    "Vendor GSTIN",
    "Total Basic Amount",
    "Total GST Amount",
    "Other / Discount Amount",
    "Total Invoice Amount",
    "Total Amount Paid against Invoice",
]


# ---------------- file loading ----------------

def load_as_images(uploaded_file, dpi: int = 300):
    """Return a list of PIL Images for one uploaded file (PDF or image)."""
    raw = uploaded_file.read()
    name = uploaded_file.name.lower()
    if name.endswith(".pdf"):
        images = []
        doc = fitz.open(stream=raw, filetype="pdf")
        for page in doc:
            pix = page.get_pixmap(dpi=dpi)
            images.append(Image.open(io.BytesIO(pix.tobytes("png"))))
        return images
    return [Image.open(io.BytesIO(raw)).convert("RGB")]


def ocr_image(img: Image.Image) -> str:
    return pytesseract.image_to_string(img)


# ---------------- field extraction (heuristic regex) ----------------

def find(pattern, text, group=1, flags=re.IGNORECASE):
    m = re.search(pattern, text, flags)
    return m.group(group).strip() if m else None


def extract_fields(text: str) -> dict:
    vendor_name = find(
        r'TAX INVOICE\s*\n+\s*[^A-Za-z0-9]*([A-Z][A-Za-z0-9 &.,\-]+?)(?:\s{2,}|\n|\s+tti|\s+~)',
        text,
    )

    invoice_date = None
    for p in [
        r'\b(\d{1,2}[-\/][A-Za-z]{3,9}[-\/]\d{2,4})\b',
        r'\b(\d{1,2}[-\/]\d{1,2}[-\/]\d{2,4})\b',
    ]:
        m = re.search(p, text)
        if m:
            invoice_date = m.group(1)
            break

    invoice_number = find(r'Invoice\s*No\.?\s*[:\-]?\s*\n?\s*([A-Za-z0-9\/\-]{1,20})', text)
    if invoice_number and invoice_number.lower() in {"dated", "delivery", "mode", "note"}:
        invoice_number = None  # OCR grabbed the wrong neighbouring label

    gstins = re.findall(r'GSTIN\s*/?\s*UIN\s*[:\-]?\s*([A-Za-z0-9]{14,15})', text)
    vendor_gstin = gstins[0] if gstins else None

    tax_amounts = re.findall(
        r'(?:CGST|SGST|IGST)\s*@?\s*[\d.]+%[^\d\n]{0,10}([\d,]+\.\d{2})', text
    )
    tax_amounts = [float(a.replace(",", "")) for a in tax_amounts]
    total_gst = round(sum(tax_amounts), 2) if tax_amounts else None

    all_amounts = [
        float(a.replace(",", "")) for a in re.findall(r'\b(\d{1,3}(?:,\d{2,3})*\.\d{2})\b', text)
    ]
    total_invoice_amount = max(all_amounts) if all_amounts else None
    total_basic_amount = (
        round(total_invoice_amount - total_gst, 2)
        if (total_invoice_amount is not None and total_gst is not None)
        else None
    )

    words_lines = re.findall(
        r'((?:Amount Chargeable|Tax Amount)\s*\(in words\)[^\n]*\n[^\n]*)', text, re.IGNORECASE
    )

    return {
        "Vendor Name": vendor_name,
        "Invoice Number": invoice_number,
        "Invoice Date": invoice_date,
        "Vendor GSTIN": vendor_gstin,
        "Total Basic Amount": total_basic_amount,
        "Total GST Amount": total_gst,
        "Other / Discount Amount": 0,
        "Total Invoice Amount": total_invoice_amount,
        "Total Amount Paid against Invoice": total_invoice_amount,
        "_raw_text": text,
        "_amount_words": "\n".join(words_lines),
    }


def to_excel_bytes(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Invoices")
        ws = writer.sheets["Invoices"]
        for col_cells in ws.columns:
            length = max(len(str(c.value)) if c.value is not None else 0 for c in col_cells)
            ws.column_dimensions[col_cells[0].column_letter].width = min(max(length + 2, 12), 40)
    return buf.getvalue()


# ---------------- UI ----------------

st.title("📄 Invoice → Excel Extractor (Free / OCR)")
st.caption(
    "100% free, runs locally with Tesseract OCR — no API key, no cost. "
    "Accuracy is lower than an AI-based extractor, so always check each row."
)

if "rows" not in st.session_state:
    st.session_state.rows = []

uploaded_files = st.file_uploader(
    "Upload invoice files (PDF, JPG, PNG) — you can select several at once",
    type=["pdf", "jpg", "jpeg", "png"],
    accept_multiple_files=True,
)

run = st.button("Extract entries", type="primary", disabled=not uploaded_files)

if run:
    progress = st.progress(0.0, text="Starting...")
    new_rows = []
    errors = []
    total = len(uploaded_files)
    for i, f in enumerate(uploaded_files):
        progress.progress(i / total, text=f"Reading {f.name}...")
        try:
            images = load_as_images(f)
            for p_idx, img in enumerate(images):
                text = ocr_image(img)
                data = extract_fields(text)
                data["Source File"] = f.name if len(images) == 1 else f"{f.name} (p{p_idx + 1})"
                new_rows.append(data)
        except pytesseract.TesseractNotFoundError:
            st.error(
                "Tesseract OCR isn't installed / not found. See the setup steps at the "
                "top of app.py or in README.md — you need to install the Tesseract "
                "program itself, not just the Python package."
            )
            break
        except Exception as e:
            errors.append(f"{f.name}: {e}")
    progress.progress(1.0, text="Done")
    st.session_state.rows.extend(new_rows)
    if errors:
        st.warning("Some files could not be read:\n" + "\n".join(errors))
    if new_rows:
        st.success(
            f"Extracted {len(new_rows)} entr{'y' if len(new_rows) == 1 else 'ies'}. "
            "Please double-check the numbers below — free OCR often misses amounts."
        )

if st.session_state.rows:
    st.subheader("Extracted entries — check and correct before downloading")
    df = pd.DataFrame(st.session_state.rows)
    edit_cols = [c for c in COLUMNS if c in df.columns] + ["Source File"]
    edited_df = st.data_editor(df[edit_cols], num_rows="dynamic", use_container_width=True)

    for i, row in df.iterrows():
        with st.expander(f"🔍 Raw OCR reference — {row.get('Source File', f'row {i+1}')}"):
            if row.get("_amount_words"):
                st.markdown("**Amount lines (usually the most reliable for totals):**")
                st.text(row["_amount_words"])
            st.markdown("**Full OCR text:**")
            st.text(row.get("_raw_text", ""))

    excel_bytes = to_excel_bytes(edited_df)
    st.download_button(
        "⬇️ Download Excel",
        data=excel_bytes,
        file_name="invoice_entries.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    if st.button("Clear all entries"):
        st.session_state.rows = []
        st.rerun()
else:
    st.info("Upload invoices above and click 'Extract entries' to begin.")
