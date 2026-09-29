# Invoice → Excel Extractor

Upload invoice PDFs or photos and get two outputs from a single extraction:

1. **Invoice table**: vendor, invoice number, date, GSTIN, basic / GST / discount / total / paid amounts.
2. **Item table**: one row per item or service (Name of Item / Service / Description, Quantity, Basic Amount, Total Amount).

## Extraction modes (choose in the sidebar)

Pick one provider, or pick **Auto fallback**: enter keys for as many providers as you like and the app tries them in order, moving to the next when one runs out of quota or fails. Free OCR can run as the last resort. The **Extracted by** column shows which provider read each invoice.

**Multiple keys per provider**: each provider starts with one API key box; click "+ Add another key" to add more boxes (for example, keys from a few friends' accounts), or "✕" to remove one. The app uses the first key, and once it hits its daily limit moves to the next key automatically — no need to babysit it. A key that hits a daily/quota error stays skipped until you click "Reset quota/rate-limit status" in the sidebar (so a fresh batch doesn't waste a request re-discovering it's still exhausted); a short per-minute rate limit is retried automatically instead.

Worth knowing: most providers' free tiers are meant per account, so pooling friends' personal keys this way may be against that provider's terms of service, even though nothing stops the app from doing it technically.

| Provider | Cost | Get a key |
|---|---|---|
| Free OCR (local) | Free, no key | Needs Tesseract installed. Lowest accuracy. |
| Gemini | Free tier (small daily quota; each PDF page = 1 request) | aistudio.google.com/apikey |
| Groq | Free tier, no card | console.groq.com/keys |
| Mistral | Free plan, no card | console.mistral.ai/api-keys |
| OpenRouter | Free models, low daily limit | openrouter.ai/keys |
| Zhipu GLM | Free vision model | open.bigmodel.cn |
| NVIDIA NIM | Free developer access | build.nvidia.com |
| Claude | Paid, about ₹1-2 per page, most accurate | console.anthropic.com (a Claude Pro subscription does not include API access) |

Model names change often. Every provider has an editable **model** box in the sidebar; if you see a "model not found" error, paste the current model name from that provider's docs. Free tiers may use your data for training, so check each provider's policy for sensitive invoices, and always review the Flags column because small free models can misread GSTINs or amounts.

## Downloads

- **Items**: one separate CSV per PDF, named after the PDF (for example `AST-APR-25-26-001.csv`), with only the four item columns. When several PDFs are processed, a ZIP of all CSVs is also offered.
- **Invoices**: `Download Invoices Excel` gives a new timestamped file every time.
- **Master file**: `Save invoices to master file` appends invoice rows to `invoice_master.xlsx` in the app folder. Items are never added to the master file.
- Each "Extract entries" click starts fresh tables. Turn on "Keep previous entries" to add to the earlier ones instead.

## Other features

- **Flags column**: warns about a GSTIN with an invalid format, Basic + GST − Discount not matching the Total, and item amounts not adding up to the invoice's basic amount.
- **Editable tables**: double-click any cell to correct a value before downloading.
- **Customize column headings** (sidebar): rename any invoice or item column.
- **Extra columns** (sidebar): add blank columns such as Remarks or Category to fill in by hand.

## Setup (one time)

1. Install Python 3.9 or newer.
2. For Free (OCR) mode only, install Tesseract OCR (Windows installer from https://github.com/tesseract-ocr/tesseract/releases, default install location).
3. In PowerShell:

```
cd "C:\Users\harsh\OneDrive\Desktop\invoice_app"
pip install -r requirements.txt
```

## Run

```
streamlit run app.py
```

The app opens in your browser at http://localhost:8501. Keep the PowerShell window open while using it; press `Ctrl + C` in that window to stop the app.

## Deploying online (optional)

Upload `app.py`, `requirements.txt`, `packages.txt` and `README.md` to a public GitHub repository and deploy it on share.streamlit.io. `packages.txt` installs Tesseract on the server.
