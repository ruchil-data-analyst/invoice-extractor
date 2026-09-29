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
import zipfile
import os
import re
import time
from datetime import datetime
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

# Line-item (second table) fields — same idea: keys fixed, labels editable.
ITEM_KEYS = ["item_description", "item_quantity", "item_basic", "item_total"]

DEFAULT_ITEM_LABELS = {
    "item_description": "Name of Item / Service / Description",
    "item_quantity": "Quantity",
    "item_basic": "Basic Amount",
    "item_total": "Total Amount",
}

GSTIN_RE = re.compile(r'^\d{2}[A-Z]{5}\d{4}[A-Z][A-Z\d]Z[A-Z\d]$')


# ---------------- file loading ----------------

def load_as_images(uploaded_file, dpi: int = 300, password: str = ""):
    raw = uploaded_file.read()
    name = uploaded_file.name.lower()
    if name.endswith(".pdf"):
        images = []
        doc = fitz.open(stream=raw, filetype="pdf")
        if doc.needs_pass and not (password and doc.authenticate(password)):
            raise ValueError(
                "this PDF is password-protected. Enter its password in the sidebar "
                "(\"PDF password\"), or save a copy without a password and upload that."
            )
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

    # Best-effort line items: rows like "1 Marshall Kilburn 2   85182200   1no".
    # Both OCR passes are in `text`, so dedupe. Amounts are only filled when
    # there is exactly one item (then they equal the invoice figures);
    # with several items OCR can't reliably tell which amount is whose.
    item_rows = re.findall(
        r'^\s*\d{1,2}\s+([A-Za-z][A-Za-z0-9 .\-/&]{2,60}?)\s+\d{4,8}\s+(\d+(?:\.\d+)?)\s*[A-Za-z]{0,5}\b',
        text,
        re.MULTILINE,
    )
    unique_items = []
    for desc, qty in item_rows:
        if (desc.strip(), qty) not in unique_items:
            unique_items.append((desc.strip(), qty))
    single = len(unique_items) == 1
    items = [
        {
            "item_description": d,
            "item_quantity": float(q) if "." in q else int(q),
            "item_basic": total_basic_amount if single else None,
            "item_total": total_invoice_amount if single else None,
        }
        for d, q in unique_items
    ]

    return {
        "_items": items,
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

- items                 (array with ONE object per line item / product / service row in the
                           invoice's item table. Each object has:
                             description  (item name as written),
                             quantity     (number only),
                             basic_amount (number only, that item's amount before tax),
                             total_amount (number only, that item's amount INCLUDING its share
                                           of GST; for a single-item invoice this equals the
                                           invoice grand total))
                         Do not invent items; do not include tax rows (CGST/SGST/IGST/rounding)
                         as items.

If a field truly cannot be found, use null. Do not guess wildly — read carefully, this feeds
an accounting sheet."""


def _parse_json_response(text: str) -> dict:
    text = re.sub(r"<think>.*?</think>", "", text or "", flags=re.DOTALL).strip()
    if text.startswith("```"):
        text = text.strip("`")
        text = text.split("\n", 1)[1] if "\n" in text else text
        text = text.rsplit("```", 1)[0]
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        # Some models wrap the JSON in extra words — grab the outermost {...}.
        start, end = text.find("{"), text.rfind("}")
        if start != -1 and end > start:
            return json.loads(text[start : end + 1])
        raise


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
        "_items": [
            {
                "item_description": it.get("description"),
                "item_quantity": it.get("quantity"),
                "item_basic": it.get("basic_amount"),
                "item_total": it.get("total_amount"),
            }
            for it in (data.get("items") or [])
            if isinstance(it, dict)
        ],
    }


def extract_fields_paid(client, img: Image.Image, model: str = "claude-sonnet-4-6") -> dict:
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    img_b64 = base64.b64encode(buf.getvalue()).decode()

    resp = client.messages.create(
        model=model,
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

def extract_fields_gemini(gemini_client, img: Image.Image, model: str = "gemini-3.8-flash") -> dict:
    from google.genai import types

    buf = io.BytesIO()
    img.save(buf, format="PNG")
    img_part = types.Part.from_bytes(data=buf.getvalue(), mime_type="image/png")

    last_error = None
    for attempt in range(4):  # try up to 4 times: instant, then 3 backoff retries
        try:
            resp = gemini_client.models.generate_content(
                model=model,
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


# ---------------- FREE AI: OpenAI-compatible providers (Groq, Mistral, OpenRouter, Zhipu, NVIDIA) ----------------

def img_to_jpeg_b64(img: Image.Image, max_side: int = 2000) -> str:
    """Downscale + JPEG so the request stays small (some providers cap image size ~4 MB)."""
    im = img.convert("RGB")
    scale = max_side / max(im.size)
    if scale < 1:
        im = im.resize((int(im.width * scale), int(im.height * scale)), Image.LANCZOS)
    buf = io.BytesIO()
    im.save(buf, format="JPEG", quality=88)
    return base64.b64encode(buf.getvalue()).decode()


def parse_retry_seconds(text: str) -> Optional[float]:
    """Read 'try again in 26.02s' / '6m30.5s' / '450ms' out of a rate-limit message."""
    m = re.search(r"try again in\s+(?:(\d+)m)?\s*([\d.]+)\s*(ms|s)\b", text or "", re.IGNORECASE)
    if not m:
        return None
    mins = float(m.group(1) or 0)
    secs = float(m.group(2)) / (1000 if m.group(3).lower() == "ms" else 1)
    return mins * 60 + secs


DAILY_LIMIT_RE = re.compile(r"per day|\bRPD\b|\bTPD\b|daily|per month|billing", re.IGNORECASE)


def rate_limit_wait(status: int, body: str) -> Optional[float]:
    """Seconds to wait if this 429 is a short per-minute limit worth retrying, else None
    (daily/monthly quota, or a wait too long to sit through)."""
    if status != 429 or DAILY_LIMIT_RE.search(body or ""):
        return None
    wait = parse_retry_seconds(body)
    if wait is None:
        wait = 20.0  # per-minute limits reset within a minute
    return wait + 1 if wait <= 90 else None


def extract_fields_openai_compat(base_url: str, api_key: str, model: str, img: Image.Image,
                                 url_as_string: bool = False) -> dict:
    import requests

    data_url = "data:image/jpeg;base64," + img_to_jpeg_b64(img)
    payload = {
        "model": model,
        "temperature": 0,
        "max_tokens": 2000,
        "messages": [
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": EXTRACTION_PROMPT},
                    {"type": "image_url", "image_url": data_url if url_as_string else {"url": data_url}},
                ],
            }
        ],
    }
    headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}
    last = None
    for attempt in range(5):
        resp = requests.post(f"{base_url.rstrip('/')}/chat/completions", json=payload,
                             headers=headers, timeout=120)
        if resp.status_code in (500, 502, 503, 504) and attempt < 4:
            last = f"{resp.status_code} {resp.text[:200]}"
            time.sleep(3 * (attempt + 1))
            continue
        wait = rate_limit_wait(resp.status_code, resp.text)
        if wait is not None and attempt < 4:
            last = f"{resp.status_code} {resp.text[:200]}"
            try:
                st.toast(f"Per-minute rate limit reached, waiting {wait:.0f}s and retrying...")
            except Exception:  # noqa: BLE001 - toast is cosmetic
                pass
            time.sleep(wait)
            continue
        if resp.status_code >= 400:
            raise RuntimeError(f"{resp.status_code} {resp.text[:300]}")
        content = resp.json()["choices"][0]["message"]["content"]
        if isinstance(content, list):  # some providers return content parts
            content = "".join(p.get("text", "") for p in content if isinstance(p, dict))
        return _map_ai_fields(_parse_json_response(content))
    raise RuntimeError(last or "provider unavailable")


# name -> settings. Model names change often, so every model is editable in the sidebar.
PROVIDERS = {
    "Gemini (free)": {
        "kind": "gemini", "env": "GOOGLE_API_KEY", "model": "gemini-3.8-flash",
        "link": "aistudio.google.com/apikey",
        "note": "Free tier has a small daily request quota (each PDF page = 1 request).",
    },
    "Groq (free)": {
        "kind": "openai", "env": "GROQ_API_KEY", "base_url": "https://api.groq.com/openai/v1",
        "model": "qwen/qwen3.8-27b", "link": "console.groq.com/keys",
        "note": "Fast, generous free limits, no card. Groq retires vision models often (Llama 4 Scout is gone), so edit the model name if you get a model-not-found error.",
    },
    "Mistral (free)": {
        "kind": "openai", "env": "MISTRAL_API_KEY", "base_url": "https://api.mistral.ai/v1",
        "model": "mistral-small-latest", "link": "console.mistral.ai/api-keys", "url_as_string": True,
        "note": "Free 'Experiment' plan, no card. Reads images.",
    },
    "OpenRouter (free)": {
        "kind": "openai", "env": "OPENROUTER_API_KEY", "base_url": "https://openrouter.ai/api/v1",
        "model": "openrouter/free", "link": "openrouter.ai/keys",
        "note": "'openrouter/free' picks a random free vision model, so quality varies. Low daily limit.",
    },
    "Zhipu GLM (free)": {
        "kind": "openai", "env": "ZHIPU_API_KEY", "base_url": "https://open.bigmodel.cn/api/paas/v4",
        "model": "glm-4.6v-flash", "link": "open.bigmodel.cn",
        "note": "Free GLM vision model; limits are not clearly documented.",
    },
    "NVIDIA NIM (free)": {
        "kind": "openai", "env": "NVIDIA_API_KEY", "base_url": "https://integrate.api.nvidia.com/v1",
        "model": "meta/llama-3.2-90b-vision-instruct", "link": "build.nvidia.com",
        "note": "Free developer credits/limits, no card. Model list changes often, so edit the model name if needed.",
    },
    "Claude (paid)": {
        "kind": "claude", "env": "ANTHROPIC_API_KEY", "model": "claude-sonnet-4-6",
        "link": "console.anthropic.com",
        "note": "About ₹1-2 per page, most accurate. A Claude Pro subscription does not include API access.",
    },
}
PROVIDER_ORDER = list(PROVIDERS)  # fallback order: free ones first, paid Claude last
MODE_OCR = "Free (OCR, local)"
MODE_AUTO = "Auto fallback (try every provider I add keys for)"


def make_provider_fn(name: str, api_key: str, model: str):
    p = PROVIDERS[name]
    cache = {}

    if p["kind"] == "gemini":
        def fn(img):
            if "c" not in cache:
                from google import genai
                cache["c"] = genai.Client(api_key=api_key)
            return extract_fields_gemini(cache["c"], img, model)
    elif p["kind"] == "claude":
        def fn(img):
            if "c" not in cache:
                from anthropic import Anthropic
                cache["c"] = Anthropic(api_key=api_key)
            return extract_fields_paid(cache["c"], img, model)
    else:
        def fn(img):
            return extract_fields_openai_compat(
                p["base_url"], api_key, model, img, p.get("url_as_string", False)
            )
    return fn


FATAL_ERROR_RE = re.compile(
    r"429|quota|RESOURCE_EXHAUSTED|rate.?limit|401|403|404|NOT_FOUND|unauthor|invalid.{0,20}key|api key",
    re.IGNORECASE,
)


def run_chain(chain: list, img: Image.Image, dead: dict):
    """Try each (label, fn) in order. A provider that hits quota / auth / model-not-found
    errors is skipped for the rest of this run. Returns (data, label_used)."""
    errors = []
    for label, fn in chain:
        if label in dead:
            errors.append(f"{label}: skipped ({dead[label][:120]})")
            continue
        try:
            return fn(img), label
        except Exception as e:  # noqa: BLE001 - we want every provider failure to fall through
            msg = str(e)
            errors.append(f"{label}: {msg[:200]}")
            if FATAL_ERROR_RE.search(msg):
                dead[label] = msg
    raise RuntimeError(" || ".join(errors))


# ---------------- Excel helpers ----------------

def _write_sheets(writer, sheets: dict):
    for name, df in sheets.items():
        df.to_excel(writer, index=False, sheet_name=name)
        ws = writer.sheets[name]
        for col_cells in ws.columns:
            length = max(len(str(c.value)) if c.value is not None else 0 for c in col_cells)
            ws.column_dimensions[col_cells[0].column_letter].width = min(max(length + 2, 12), 40)


def to_excel_bytes(sheets: dict) -> bytes:
    """sheets: {"Invoices": df, "Items": df} -> one workbook, one sheet each."""
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        _write_sheets(writer, sheets)
    return buf.getvalue()


def append_to_master(sheets: dict) -> dict:
    """Append each sheet's rows to the same-named sheet of the master file
    (older single-sheet master files are treated as the 'Invoices' sheet)."""
    existing = {}
    if os.path.exists(MASTER_FILE):
        book = pd.read_excel(MASTER_FILE, sheet_name=None)
        for name, old_df in book.items():
            key = name if name in sheets else ("Invoices" if name in ("Sheet1", "Sheet") else name)
            existing[key] = old_df
    combined = {}
    for name, df in sheets.items():
        old_df = existing.pop(name, None)
        combined[name] = pd.concat([old_df, df], ignore_index=True) if old_df is not None else df
    combined.update(existing)  # keep any other sheets untouched
    with pd.ExcelWriter(MASTER_FILE, engine="openpyxl") as writer:
        _write_sheets(writer, combined)
    return {name: len(df) for name, df in combined.items()}


def items_check(source: str, rows_basic, items) -> Optional[str]:
    """Flag when this invoice's line items don't add up to its basic amount."""
    mine = [it for it in items if it.get("Source File") == source]
    if not mine or rows_basic in (None, ""):
        return None
    try:
        item_sum = sum(float(it.get("item_basic") or 0) for it in mine)
        basic = float(rows_basic)
    except (TypeError, ValueError):
        return None
    if item_sum and abs(item_sum - basic) > 1.0:
        return f"Items basic total {item_sum:.2f} ≠ invoice basic {basic:.2f}"
    return None


def pdf_base_name(source: str) -> str:
    """'AST-APR-25-26-001.pdf (p2)' -> 'AST-APR-25-26-001' (pages of one PDF share a CSV)."""
    name = re.sub(r"\s*\(p\d+\)\s*$", "", str(source or "").strip())
    name = re.sub(r"\.(pdf|jpg|jpeg|png)$", "", name, flags=re.IGNORECASE)
    name = re.sub(r'[\\/:*?"<>|]', "_", name).strip()
    return name or "invoice"


def items_csv_per_pdf(items: list, item_labels: dict) -> dict:
    """One CSV per source PDF: {'<pdf name>.csv': bytes}. Only the item columns
    (same layout as the sample CSV); utf-8-sig so Excel opens it cleanly."""
    cols = [item_labels[k] for k in ITEM_KEYS]
    groups = {}
    for it in items:
        groups.setdefault(pdf_base_name(it.get("Source File")), []).append(
            {item_labels[k]: it.get(k) for k in ITEM_KEYS}
        )
    files = {}
    for base, rows in groups.items():
        df = pd.DataFrame(rows, columns=cols)
        files[f"{base}.csv"] = df.to_csv(index=False).encode("utf-8-sig")
    return files


def zip_files(files: dict) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in files.items():
            z.writestr(name, data)
    return buf.getvalue()


# ---------------- session state setup ----------------

if "rows" not in st.session_state:
    st.session_state.rows = []  # each row: internal keys + Source File + _raw_text + _amount_words + extra col values
if "labels" not in st.session_state:
    st.session_state.labels = dict(DEFAULT_LABELS)
if "extra_columns" not in st.session_state:
    st.session_state.extra_columns = []
if "api_keys" not in st.session_state:
    st.session_state.api_keys = {}
if "models" not in st.session_state:
    st.session_state.models = {}
if "line_items" not in st.session_state:
    st.session_state.line_items = []  # line items: item keys + Source File
if "dead_providers" not in st.session_state:
    st.session_state.dead_providers = {}  # label -> last error, persists until reset/app restart
if "key_slots" not in st.session_state:
    st.session_state.key_slots = {}  # provider name -> list of slot ids (one text box per id)
if "key_slot_counter" not in st.session_state:
    st.session_state.key_slot_counter = {}  # provider name -> next slot id to hand out
if "item_labels" not in st.session_state:
    st.session_state.item_labels = dict(DEFAULT_ITEM_LABELS)


# ---------------- UI ----------------

st.title("📄 Invoice → Excel Extractor")

with st.sidebar:
    st.header("Settings")
    mode = st.selectbox("Extraction mode", [MODE_OCR] + PROVIDER_ORDER + [MODE_AUTO])

    def ensure_key_slots(name: str):
        if not st.session_state.key_slots.get(name):
            seed = st.session_state.api_keys.get(name) or [os.environ.get(PROVIDERS[name]["env"], "")]
            ids = list(range(len(seed)))
            st.session_state.key_slots[name] = ids
            st.session_state.key_slot_counter[name] = len(ids)
            for i, val in zip(ids, seed):
                st.session_state.setdefault(f"widget_key_{name}_{i}", val)

    def provider_inputs(name: str):
        p = PROVIDERS[name]
        ensure_key_slots(name)
        slots = st.session_state.key_slots[name]
        keys = []
        for i, slot_id in enumerate(list(slots)):
            wkey = f"widget_key_{name}_{slot_id}"
            if len(slots) > 1:
                c1, c2 = st.columns([5, 1])
            else:
                c1 = st
            val = c1.text_input(
                f"{name} API key" + (f" #{i + 1}" if len(slots) > 1 else ""),
                type="password", key=wkey,
            )
            keys.append(val)
            if len(slots) > 1 and c2.button("✕", key=f"remove_{wkey}"):
                slots.remove(slot_id)
                st.session_state.pop(wkey, None)
                st.rerun()
        if st.button(f"+ Add another {name.split(' (')[0]} key", key=f"add_key_{name}"):
            new_id = st.session_state.key_slot_counter[name]
            st.session_state.key_slot_counter[name] += 1
            slots.append(new_id)
            st.rerun()
        st.session_state.api_keys[name] = [k.strip() for k in keys if k.strip()]
        st.caption("Add one box per key if you have several (e.g. from friends) — the app moves "
                   "to the next key automatically once one hits its daily limit.")
        st.session_state.models[name] = st.text_input(
            f"{name} model",
            value=st.session_state.models.get(name) or p["model"],
            key=f"widget_model_{name}",
            help="Model names change often. If you get a 'model not found' error, paste the current name from the provider's docs.",
        )
        st.caption(f"Get a key: {p['link']}. {p['note']}")

    use_ocr_last = True
    if mode == MODE_OCR:
        st.caption("Free and local, no key needed. Lowest accuracy: always check flagged rows.")
    elif mode == MODE_AUTO:
        st.caption(
            "Tries providers in this order and moves to the next when one runs out of quota or fails: "
            + " → ".join(n.split(" (")[0] for n in PROVIDER_ORDER)
            + ". Only providers with a key are used."
        )
        for n in PROVIDER_ORDER:
            with st.expander(n, expanded=bool(st.session_state.api_keys.get(n))):
                provider_inputs(n)
        use_ocr_last = st.checkbox(
            "Use Free OCR as the last resort if every provider fails", value=True
        )
        if st.session_state.dead_providers:
            st.caption(f"Currently out of quota: {', '.join(st.session_state.dead_providers)}")
            if st.button("Reset quota/rate-limit status"):
                st.session_state.dead_providers = {}
                st.rerun()
    else:
        provider_inputs(mode)

    pdf_password = st.text_input(
        "PDF password (only if your PDFs are protected)", type="password", key="pdf_password"
    )

    st.divider()
    with st.expander("✏️ Customize column headings"):
        st.caption("Rename any column to match how you want it in Excel.")
        for key in FIELD_KEYS:
            st.session_state.labels[key] = st.text_input(
                key.replace("_", " ").title(), value=st.session_state.labels[key], key=f"label_{key}"
            )
        st.caption("Item table headings:")
        for key in ITEM_KEYS:
            st.session_state.item_labels[key] = st.text_input(
                key.replace("_", " ").title(),
                value=st.session_state.item_labels[key],
                key=f"label_{key}",
            )
        if st.button("Reset headings to default"):
            st.session_state.labels = dict(DEFAULT_LABELS)
            st.session_state.item_labels = dict(DEFAULT_ITEM_LABELS)
            for k in FIELD_KEYS + ITEM_KEYS:
                st.session_state.pop(f"label_{k}", None)
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

keep_old = st.checkbox(
    "Keep previous entries (off = fresh entries on every extraction)",
    value=False,
    help="When off, each 'Extract entries' click clears the previous tables, so only "
    "the current invoices are shown and downloaded.",
)
run = st.button("Extract entries", type="primary", disabled=not uploaded_files)

def ocr_fn(img):
    return extract_fields_free(ocr_image(img))


def provider_key_chain(name: str) -> list:
    """One (label, fn) per key entered for this provider, e.g. 'Groq (free) #1', '#2'.
    A single key still gets numbered for a consistent, stable label."""
    model = st.session_state.models.get(name) or PROVIDERS[name]["model"]
    keys = st.session_state.api_keys.get(name) or []
    n = len(keys)
    return [
        (f"{name} #{i + 1}" if n > 1 else name, make_provider_fn(name, key, model))
        for i, key in enumerate(keys)
    ]


def build_chain():
    """Return the ordered list of (label, fn) for the selected mode, or None on a setup error."""
    if mode == MODE_OCR:
        return [("Free OCR", ocr_fn)]
    if mode == MODE_AUTO:
        names = [n for n in PROVIDER_ORDER if st.session_state.api_keys.get(n)]
    else:
        names = [mode]
        if not st.session_state.api_keys.get(mode):
            st.error(f"Enter at least one {mode} API key in the sidebar first.")
            return None
    chain = [entry for n in names for entry in provider_key_chain(n)]
    if mode == MODE_AUTO and use_ocr_last:
        chain.append(("Free OCR", ocr_fn))
    if not chain:
        st.error("Add at least one API key (or turn on the Free OCR last resort) in the sidebar.")
        return None
    return chain


if run:
    chain = build_chain()
    if chain is None:
        st.stop()
    dead = st.session_state.dead_providers  # persists across runs; reset via the sidebar button

    progress = st.progress(0.0, text="Starting...")
    new_rows = []
    new_items = []
    errors = []
    total = len(uploaded_files)
    for i, f in enumerate(uploaded_files):
        progress.progress(i / total, text=f"Reading {f.name}...")
        try:
            images = load_as_images(f, password=pdf_password)
            for p_idx, img in enumerate(images):
                data, used = run_chain(chain, img, dead)
                data["Extracted by"] = used
                data["Source File"] = f.name if len(images) == 1 else f"{f.name} (p{p_idx + 1})"
                for it in data.pop("_items", []) or []:
                    new_items.append({**it, "Source File": data["Source File"]})
                data["Flags"] = compute_flags(data)
                for c in st.session_state.extra_columns:
                    data.setdefault(c, "")
                new_rows.append(data)
        except Exception as e:
            errors.append(f"{f.name}: {e}")
    progress.progress(1.0, text="Done")
    if keep_old:
        st.session_state.rows.extend(new_rows)
        st.session_state.line_items.extend(new_items)
    else:
        st.session_state.rows = new_rows
        st.session_state.line_items = new_items
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
        rec["Extracted by"] = row.get("Extracted by", "")
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
        new_row["Extracted by"] = rec.get("Extracted by", "")
        new_row["Flags"] = compute_flags(new_row)
        if i < len(stored_rows):
            new_row["_raw_text"] = stored_rows[i].get("_raw_text", "")
            new_row["_amount_words"] = stored_rows[i].get("_amount_words", "")
        else:
            new_row["_raw_text"] = ""
            new_row["_amount_words"] = ""
        updated_rows.append(new_row)
    st.session_state.rows = updated_rows

    # ---- second table: line items (one row per product/service) ----
    item_labels = st.session_state.item_labels
    stored_items = st.session_state.line_items
    st.subheader("Item details — one row per item in each invoice")
    items_df = pd.DataFrame(
        [
            {**{item_labels[k]: it.get(k) for k in ITEM_KEYS}, "Source File": it.get("Source File", "")}
            for it in stored_items
        ],
        columns=[item_labels[k] for k in ITEM_KEYS] + ["Source File"],
    )
    edited_items = st.data_editor(
        items_df, num_rows="dynamic", use_container_width=True, key="items_editor"
    )
    ilabel_to_key = {v: k for k, v in item_labels.items()}
    updated_items = []
    for rec in edited_items.to_dict("records"):
        new_it = {ilabel_to_key[l]: rec.get(l) for l in ilabel_to_key}
        new_it["Source File"] = rec.get("Source File", "")
        if any(v not in (None, "") and v == v for k, v in new_it.items() if k != "Source File"):
            updated_items.append(new_it)
    st.session_state.line_items = updated_items

    csv_files = items_csv_per_pdf(updated_items, item_labels)
    if csv_files:
        stamp_csv = datetime.now().strftime("%Y-%m-%d_%H%M%S")
        st.markdown("**⬇️ Items CSV — one separate CSV per PDF**")
        if len(csv_files) > 1:
            st.download_button(
                f"⬇️ All {len(csv_files)} CSV files as one ZIP",
                data=zip_files(csv_files),
                file_name=f"item_csvs_{stamp_csv}.zip",
                mime="application/zip",
                key="dl_items_zip",
            )
        cols_dl = st.columns(min(len(csv_files), 3))
        for i, (fname, data) in enumerate(csv_files.items()):
            with cols_dl[i % len(cols_dl)]:
                st.download_button(
                    f"⬇️ {fname}",
                    data=data,
                    file_name=fname,
                    mime="text/csv",
                    key=f"dl_csv_{i}_{fname}",
                )

    # invoice-level flag: do the items add up to the invoice's basic amount?
    for r in updated_rows:
        extra = items_check(r.get("Source File"), r.get("basic_amount"), updated_items)
        if extra:
            r["Flags"] = (r["Flags"] + " | " if r["Flags"] else "") + extra
    st.session_state.rows = updated_rows

    if any(r.get("Flags") for r in updated_rows):
        st.warning("Some rows are still flagged — check the 'Flags' column before saving.")

    if True:
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
             "Extracted by": r.get("Extracted by", ""),
             **{c: r.get(c, "") for c in extra_cols},
             "Source File": r.get("Source File", "")}
            for r in updated_rows
        ]
    )

    final_items = pd.DataFrame(
        [
            {**{item_labels[k]: it.get(k) for k in ITEM_KEYS}, "Source File": it.get("Source File", "")}
            for it in updated_items
        ],
        columns=[item_labels[k] for k in ITEM_KEYS] + ["Source File"],
    )
    export_sheets = {
        "Invoices": final_display.drop(columns=["Flags", "Extracted by"], errors="ignore"),
        "Items": final_items,
    }

    stamp = datetime.now().strftime("%Y-%m-%d_%H%M%S")
    col1, col3 = st.columns(2)
    with col1:
        st.download_button(
            "⬇️ Download Invoices Excel",
            data=to_excel_bytes({"Invoices": export_sheets["Invoices"]}),
            file_name=f"invoice_entries_{stamp}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    with col3:
        if st.button("💾 Save invoices to master file (invoice_master.xlsx)"):
            counts = append_to_master({"Invoices": export_sheets["Invoices"]})
            st.success(
                "Saved. Master file now has "
                + ", ".join(f"{n} row(s) in '{name}'" for name, n in counts.items())
                + f" — {MASTER_FILE}"
            )

    if st.button("Clear this batch"):
        st.session_state.rows = []
        st.session_state.line_items = []
        st.rerun()
else:
    st.info("Upload invoices above and click 'Extract entries' to begin.")
