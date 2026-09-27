"""
Invoice → Excel Extractor (v4 — three modes)
-----------------------------------------------
Three modes, chosen in the sidebar:

  1. FREE (OCR)   — 100% local OCR (Tesseract). No API key, no cost,
                    lowest accuracy.
  2. FREE AI      — Google Gemini's free API tier. No cost (rate-limited,
                    no card needed), AI-level understanding — much better
                    than plain OCR at reading messy/handwritten bits.
  3. PAID (Claude)— Claude AI vision. Needs an Anthropic API key, costs
                    roughly ₹1-2 per invoice, the most consistently
                    accurate of the three.

New in v3:
  - Column headings are editable from the sidebar ("Customize column
    headings") — rename any of the 9 standard fields to match a
    particular client's / vendor's sheet format.
  - You can add extra blank columns (e.g. "Remarks", "Category") that
    you fill in manually per row — these get saved and exported too.

v2 improvements (still here):
  - Dual-pass OCR (plain + sharpened) merged for better free-mode accuracy.
  - GSTIN format check + Basic/GST/Discount vs Total math cross-check,
    surfaced in a "Flags" column.
  - "Save to master file" appends checked rows to invoice_master.xlsx in
    this folder so entries build up across sessions.

Setup:
  pip install -r requirements.txt
  FREE mode also needs the Tesseract OCR program itself installed
  separately (see README.md).

Run:
  streamlit run app.py
"""

import base64
import io
import json
import os
import re
import time
from typing import Optional

import fitz  # PyMuPDF
import pandas as pd
import streamlit as st
from PIL import Image, ImageOps, ImageFilter

st.set_page_config(page_title="Invoice → Excel Extractor", layout="wide")

MASTER_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "invoice_master.xlsx")

# Internal field key -> default display label. The key never changes;
# only the label (what shows in the table / Excel header) is editable.
FIELD_KEYS = [
    "vendor_name",
    "invoice_number",
    "invoice_date",
    "vendor_gstin",
    "basic_amount",
    "gst_amount",
    "discount_amount",
    "total_amount",
    "paid_amount",
]

DEFAULT_LABELS = {
    "vendor_name": "Vendor Name",
    "invoice_number": "Invoice Number",
    "invoice_date": "Invoice Date",
    "vendor_gstin": "Vendor GSTIN",
    "basic_amount": "Total Basic Amount",
    "gst_amount": "Total GST Amount",
    "discount_amount": "Other / Discount Amount",
    "total_amount": "Total Invoice Amount",
    "paid_amount": "Total Amount Paid against Invoice",
}

GSTIN_RE = re.compile(r'^\d{2}[A-Z]{5}\d{4}[A-Z][A-Z\d]Z[A-Z\d]$')


# ---------------- file loading ----------------

def load_as_images(uploaded_file, dpi: int = 300):
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


def preprocess_variants(img: Image.Image):
    """Return a couple of cleaned-up versions of the image. Different
    invoices respond differently to sharpening — sometimes it helps
    Tesseract catch a boxed total, sometimes it costs a faint date. We
    OCR both variants and merge the results rather than betting on one."""
    gray = ImageOps.grayscale(img)
    plain = ImageOps.autocontrast(gray, cutoff=1)
    sharp = plain.filter(ImageFilter.SHARPEN)
    return [plain, sharp]


# ---------------- shared post-processing (operates on internal keys) ----------------

def validate_gstin(gstin) -> Optional[str]:
    if not gstin:
        return None
    cleaned = re.sub(r'\s+', '', str(gstin)).upper()
    if GSTIN_RE.match(cleaned):
        return None
    return "GSTIN format looks off — please verify"


def math_check(row: dict) -> Optional[str]:
    try:
        basic = float(row.get("basic_amount") or 0)
        gst = float(row.get("gst_amount") or 0)
        disc = float(row.get("discount_amount") or 0)
        total = float(row.get("total_amount") or 0)
    except (TypeError, ValueError):
        return None
    if total == 0:
        return None
    expected = basic + gst - disc
    if abs(expected - total) > 1.0:  # allow ~₹1 rounding slack
        return f"Basic+GST-Discount = {expected:.2f}, doesn't match Total {total:.2f}"
    return None


def compute_flags(row: dict) -> str:
    flags = []
    g = validate_gstin(row.get("vendor_gstin"))
    if g:
        flags.append(g)
    m = math_check(row)
    if m:
        flags.append(m)
    return " | ".join(flags) if flags else ""


# ---------------- FREE mode: OCR + regex (returns internal keys) ----------------

def ocr_image(img: Image.Image) -> str:
    import pytesseract

    tesseract_path = os.environ.get("TESSERACT_PATH")
    if tesseract_path:
        pytesseract.pytesseract.tesseract_cmd = tesseract_path
    elif os.name == "nt":
        default_win_path = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
        if os.path.exists(default_win_path):
            pytesseract.pytesseract.tesseract_cmd = default_win_path

    texts = [pytesseract.image_to_string(v) for v in preprocess_variants(img)]
    return "\n".join(texts)


def find(pattern, text, group=1, flags=re.IGNORECASE):
    m = re.search(pattern, text, flags)
    return m.group(group).strip() if m else None


def extract_fields_free(text: str) -> dict:
    vendor_name = find(
        r'TAX INVOICE\s*\n+\s*[^A-Za-z0-9]*([A-Za-z][A-Za-z0-9 &.,\-]+?)'
        r'(?:\s{2,}|\n|\s+tti|\s+~|\s+Invoice\s*No\b)',
        text,
    )
    if vendor_name:
        parts = vendor_name.split(" ", 1)
        if len(parts) == 2 and len(parts[0]) <= 2 and parts[0].islower() and parts[1][:1].isupper():
            vendor_name = parts[1]
        vendor_name = vendor_name.strip(" .,-")

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
        invoice_number = None

    gstins = re.findall(r'GSTIN\s*/?\s*UIN\s*[:\-]?\s*([A-Za-z0-9]{14,15})', text)
    vendor_gstin = gstins[0] if gstins else None

    tax_matches = re.findall(
        r'(CGST|SGST|IGST)\s*@?\s*[\d.]+%[^\d\n]{0,10}([\d,]+\.\d{2})', text, re.IGNORECASE
    )
    seen = {(t.upper(), float(a.replace(",", ""))) for t, a in tax_matches}
    tax_amounts = [a for _, a in seen]
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
        "vendor_name": vendor_name,
        "invoice_number": invoice_number,
        "invoice_date": invoice_date,
        "vendor_gstin": vendor_gstin,
        "basic_amount": total_basic_amount,
        "gst_amount": total_gst,
        "discount_amount": 0,
        "total_amount": total_invoice_amount,
        "paid_amount": total_invoice_amount,
        "_raw_text": text,
        "_amount_words": "\n".join(words_lines),
    }


# ---------------- PAID mode: Claude vision (returns internal keys) ----------------

EXTRACTION_PROMPT = """You are extracting structured data from a scanned tax invoice image.
Return ONLY a raw JSON object (no markdown fences, no explanation) with exactly these keys:

- vendor_name            (the SELLER / supplier issuing the invoice, not the buyer)
- invoice_number
- invoice_date           (format DD-MM-YYYY)
- vendor_gstin           (the seller's GSTIN/UIN)
- total_basic_amount     (number only, taxable value before any tax, no commas or currency symbol)
- total_gst_amount       (number only, sum of CGST+SGST or IGST, no commas)
- other_discount_amount  (number only; 0 if none mentioned)
- total_invoice_amount   (number only, the final grand total on the invoice, no commas)
- total_amount_paid      (number only; if a handwritten note/stamp shows an amount actually
                           paid, use that; otherwise repeat total_invoice_amount)

If a field truly cannot be found, use null. Do not guess wildly — read carefully, this feeds
an accounting sheet."""


def _parse_json_response(text: str) -> dict:
    text = text.strip()
    if text.startswith("```"):
        text = text.strip("`")
        text = text.split("\n", 1)[1] if "\n" in text else text
        text = text.rsplit("```", 1)[0]
    return json.loads(text)


def _map_ai_fields(data: dict) -> dict:
    return {
        "vendor_name": data.get("vendor_name"),
        "invoice_number": data.get("invoice_number"),
        "invoice_date": data.get("invoice_date"),
        "vendor_gstin": data.get("vendor_gstin"),
        "basic_amount": data.get("total_basic_amount"),
        "gst_amount": data.get("total_gst_amount"),
        "discount_amount": data.get("other_discount_amount", 0),
        "total_amount": data.get("total_invoice_amount"),
        "paid_amount": data.get("total_amount_paid"),
        "_raw_text": "",
        "_amount_words": "",
    }


def extract_fields_paid(client, img: Image.Image) -> dict:
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    img_b64 = base64.b64encode(buf.getvalue()).decode()

    resp = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=1000,
        messages=[
            {
                "role": "user",
                "content": [
                    {
                        "type": "image",
                        "source": {"type": "base64", "media_type": "image/png", "data": img_b64},
                    },
                    {"type": "text", "text": EXTRACTION_PROMPT},
                ],
            }
        ],
    )
    text = "".join(b.text for b in resp.content if b.type == "text")
    return _map_ai_fields(_parse_json_response(text))


# ---------------- FREE AI mode: Google Gemini (free tier, needs a Google AI Studio key) ----------------

def extract_fields_gemini(gemini_client, img: Image.Image) -> dict:
    from google.genai import types

    buf = io.BytesIO()
    img.save(buf, format="PNG")
    img_part = types.Part.from_bytes(data=buf.getvalue(), mime_type="image/png")

    last_error = None
    for attempt in range(4):  # try up to 4 times: instant, then 3 backoff retries
        try:
            resp = gemini_client.models.generate_content(
                model="gemini-3.8-flash",
                contents=[EXTRACTION_PROMPT, img_part],
            )
            return _map_ai_fields(_parse_json_response(resp.text))
        except Exception as e:
            last_error = e
            msg = str(e)
            # 503/UNAVAILABLE is Gemini's free tier being overloaded — worth
            # a short retry. Anything else (bad key, quota fully exhausted)
            # won't fix itself, so fail fast instead of wasting time.
            if "503" not in msg and "UNAVAILABLE" not in msg:
                raise
            if attempt < 3:
                time.sleep(3 * (attempt + 1))  # 3s, 6s, 9s
    raise last_error


# ---------------- Excel helpers ----------------

def to_excel_bytes(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Invoices")
        ws = writer.sheets["Invoices"]
        for col_cells in ws.columns:
            length = max(len(str(c.value)) if c.value is not None else 0 for c in col_cells)
            ws.column_dimensions[col_cells[0].column_letter].width = min(max(length + 2, 12), 40)
    return buf.getvalue()


def append_to_master(df: pd.DataFrame) -> int:
    if os.path.exists(MASTER_FILE):
        existing = pd.read_excel(MASTER_FILE)
        combined = pd.concat([existing, df], ignore_index=True)
    else:
        combined = df
    combined.to_excel(MASTER_FILE, index=False)
    return len(combined)


# ---------------- session state setup ----------------

if "rows" not in st.session_state:
    st.session_state.rows = []  # each row: internal keys + Source File + _raw_text + _amount_words + extra col values
if "labels" not in st.session_state:
    st.session_state.labels = dict(DEFAULT_LABELS)
if "extra_columns" not in st.session_state:
    st.session_state.extra_columns = []


# ---------------- UI ----------------

st.title("📄 Invoice → Excel Extractor")

with st.sidebar:
    st.header("Settings")
    mode = st.radio(
        "Extraction mode",
        ["Free (OCR, local)", "Free AI (Gemini)", "Paid (AI, most accurate)"],
    )
    api_key_input = ""
    if mode.startswith("Paid"):
        api_key_input = st.text_input(
            "Anthropic API key", value=os.environ.get("ANTHROPIC_API_KEY", ""), type="password"
        )
        st.caption("~₹1-2 per invoice. Key is used only for this session, never saved.")
    elif mode.startswith("Free AI"):
        api_key_input = st.text_input(
            "Google AI Studio API key",
            value=os.environ.get("GOOGLE_API_KEY", ""),
            type="password",
        )
        st.caption(
            "Bilkul free — key banao aistudio.google.com/apikey se (card nahi chahiye). "
            "AI jaisi samajh milti hai but free tier me rate-limit hai, thoda slow ho sakta hai."
        )
    else:
        st.caption("Free and local. Lower accuracy — always check flagged rows below.")

    st.divider()
    with st.expander("✏️ Customize column headings"):
        st.caption("Rename any column to match how you want it in Excel.")
        for key in FIELD_KEYS:
            st.session_state.labels[key] = st.text_input(
                key.replace("_", " ").title(), value=st.session_state.labels[key], key=f"label_{key}"
            )
        if st.button("Reset headings to default"):
            st.session_state.labels = dict(DEFAULT_LABELS)
            st.rerun()

    with st.expander("➕ Extra columns (filled in manually)"):
        st.caption("Add blank columns for anything not auto-extracted, e.g. Remarks, Category.")
        new_col = st.text_input("New column name", key="new_extra_col")
        if st.button("Add column") and new_col.strip():
            if new_col.strip() not in st.session_state.extra_columns:
                st.session_state.extra_columns.append(new_col.strip())
                st.rerun()
        for c in list(st.session_state.extra_columns):
            col_a, col_b = st.columns([4, 1])
            col_a.write(c)
            if col_b.button("✕", key=f"remove_{c}"):
                st.session_state.extra_columns.remove(c)
                st.rerun()

st.caption("Upload invoice PDFs or photos. Each one becomes a row below.")

uploaded_files = st.file_uploader(
    "Upload invoice files (PDF, JPG, PNG) — you can select several at once",
    type=["pdf", "jpg", "jpeg", "png"],
    accept_multiple_files=True,
)

run = st.button("Extract entries", type="primary", disabled=not uploaded_files)

if run:
    client = None
    gemini_client = None
    if mode.startswith("Paid"):
        if not api_key_input:
            st.error("Enter your Anthropic API key in the sidebar first.")
            st.stop()
        from anthropic import Anthropic

        client = Anthropic(api_key=api_key_input)
    elif mode.startswith("Free AI"):
        if not api_key_input:
            st.error("Enter your Google AI Studio API key in the sidebar first.")
            st.stop()
        from google import genai

        gemini_client = genai.Client(api_key=api_key_input)

    progress = st.progress(0.0, text="Starting...")
    new_rows = []
    errors = []
    total = len(uploaded_files)
    for i, f in enumerate(uploaded_files):
        progress.progress(i / total, text=f"Reading {f.name}...")
        try:
            images = load_as_images(f)
            for p_idx, img in enumerate(images):
                if mode.startswith("Paid"):
                    data = extract_fields_paid(client, img)
                elif mode.startswith("Free AI"):
                    data = extract_fields_gemini(gemini_client, img)
                else:
                    text = ocr_image(img)
                    data = extract_fields_free(text)
                data["Source File"] = f.name if len(images) == 1 else f"{f.name} (p{p_idx + 1})"
                data["Flags"] = compute_flags(data)
                for c in st.session_state.extra_columns:
                    data.setdefault(c, "")
                new_rows.append(data)
        except Exception as e:
            errors.append(f"{f.name}: {e}")
    progress.progress(1.0, text="Done")
    st.session_state.rows.extend(new_rows)
    if errors:
        st.warning("Some files could not be read:\n" + "\n".join(errors))
    if new_rows:
        flagged = sum(1 for r in new_rows if r.get("Flags"))
        msg = f"Extracted {len(new_rows)} entr{'y' if len(new_rows) == 1 else 'ies'}."
        if flagged:
            msg += f" ⚠️ {flagged} row(s) flagged for review."
        st.success(msg)

if st.session_state.rows:
    st.subheader("Extracted entries — check flagged rows before saving")

    labels = st.session_state.labels
    extra_cols = st.session_state.extra_columns
    stored_rows = st.session_state.rows

    # Build the display table using current (possibly just-renamed) labels.
    display_records = []
    for row in stored_rows:
        rec = {labels[k]: row.get(k) for k in FIELD_KEYS}
        rec["Flags"] = row.get("Flags", "")
        for c in extra_cols:
            rec[c] = row.get(c, "")
        rec["Source File"] = row.get("Source File", "")
        display_records.append(rec)
    display_df = pd.DataFrame(display_records)

    edited_df = st.data_editor(display_df, num_rows="dynamic", use_container_width=True)

    # Map edits back onto internal keys (order-based; new manually-added
    # rows won't have OCR raw text behind them, which is expected).
    label_to_key = {v: k for k, v in labels.items()}
    updated_rows = []
    edited_records = edited_df.to_dict("records")
    for i, rec in enumerate(edited_records):
        new_row = {}
        for label, key in label_to_key.items():
            new_row[key] = rec.get(label)
        for c in extra_cols:
            new_row[c] = rec.get(c, "")
        new_row["Source File"] = rec.get("Source File", "")
        new_row["Flags"] = compute_flags(new_row)
        if i < len(stored_rows):
            new_row["_raw_text"] = stored_rows[i].get("_raw_text", "")
            new_row["_amount_words"] = stored_rows[i].get("_amount_words", "")
        else:
            new_row["_raw_text"] = ""
            new_row["_amount_words"] = ""
        updated_rows.append(new_row)
    st.session_state.rows = updated_rows

    if any(r.get("Flags") for r in updated_rows):
        st.warning("Some rows are still flagged — check the 'Flags' column before saving.")

    if mode.startswith("Free (OCR"):
        for i, row in enumerate(updated_rows):
            if row.get("_raw_text"):
                with st.expander(f"🔍 Raw OCR reference — {row.get('Source File', f'row {i+1}')}"):
                    if row.get("_amount_words"):
                        st.markdown("**Amount lines (usually most reliable for totals):**")
                        st.text(row["_amount_words"])
                    st.markdown("**Full OCR text:**")
                    st.text(row.get("_raw_text", ""))

    final_display = pd.DataFrame(
        [
            {**{labels[k]: r.get(k) for k in FIELD_KEYS},
             "Flags": r.get("Flags", ""),
             **{c: r.get(c, "") for c in extra_cols},
             "Source File": r.get("Source File", "")}
            for r in updated_rows
        ]
    )

    col1, col2 = st.columns(2)
    with col1:
        excel_bytes = to_excel_bytes(final_display.drop(columns=["Flags"], errors="ignore"))
        st.download_button(
            "⬇️ Download Excel (this batch)",
            data=excel_bytes,
            file_name="invoice_entries.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    with col2:
        if st.button("💾 Save to master file (invoice_master.xlsx)"):
            total_rows = append_to_master(final_display.drop(columns=["Flags"], errors="ignore"))
            st.success(f"Saved. Master file now has {total_rows} row(s) — {MASTER_FILE}")

    if st.button("Clear this batch"):
        st.session_state.rows = []
        st.rerun()
else:
    st.info("Upload invoices above and click 'Extract entries' to begin.")
