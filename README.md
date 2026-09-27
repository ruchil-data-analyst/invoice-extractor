# Invoice → Excel Extractor (v4)

Ab TEEN modes hain (sidebar se switch karo):

1. **Free (OCR, local)** — bilkul free, Tesseract se local chalta hai.
   Sabse kam accurate.
2. **Free AI (Gemini)** — Google Gemini ka free tier, bilkul free (koi
   card nahi chahiye), AI jaisi samajh milti hai — handwritten notes,
   confusing layouts wagera Free OCR se kaafi behtar padh leta hai.
   Free tier rate-limited hai, isliye bulk invoices me thoda slow ho
   sakta hai ya kabhi-kabhi retry lag sakta hai.
3. **Paid (Claude)** — sabse accurate, ~₹1-2/invoice.

## Free AI (Gemini) key kaise banayein

1. https://aistudio.google.com/apikey pe jao
2. Google account se login karo (koi card/billing nahi chahiye)
3. "Create API key" dabao, key copy kar lo
4. App ke sidebar me "Free AI (Gemini)" mode select karke wahi key paste
   kar do

## v3 features (waise hi hain)

- **Column headings edit** kar sakte ho sidebar se
- **Extra columns** add kar sakte ho (Remarks, Category, etc.)
- **Flags column** — GSTIN format aur math mismatch automatically pakad
  leta hai
- **Master file save** — invoice_master.xlsx me entries jama hoti rahengi

## Setup

### 1) Tesseract OCR install karo (sirf "Free (OCR)" mode ke liye zaroori)

Windows: https://github.com/tesseract-ocr/tesseract/releases se
`tesseract-ocr-w64-setup-*.exe` download karke install karo.

### 2) Libraries install karo

```
pip install -r requirements.txt
```

## Chalao

```
streamlit run app.py
```

## Use kaise karein

1. Sidebar me mode choose karo.
2. Agar Free AI ya Paid mode hai, API key daalo.
3. Chaho to column headings/extra columns customize kar lo.
4. Invoice files upload karo, "Extract entries" dabao.
5. Flags column check karo, jo bhi manually fill karna hai wo bhar do.
6. "Download Excel" ya "Save to master file" se final file le lo.

## Kaunsa mode kab use karein

- Roz ke normal, clean-printed invoices → **Free (OCR)** try karo pehle
- Agar OCR bahut galtiyan kar raha ho, ya thodi messy/handwritten
  invoices hain → **Free AI (Gemini)** try karo, bilkul free hai
- Agar bulk/important invoices hain jaha galti afford nahi kar sakte,
  ya Gemini rate-limit se rukk raha hai → **Paid (Claude)**, bahut kam
  cost pe sabse reliable
