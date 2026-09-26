# Invoice → Excel Extractor (FREE version)

Ye version bilkul FREE hai — koi API key nahi, koi cost nahi. Sab kuch
tumhare computer pe hi chalta hai, Tesseract OCR ka use karke.

**Important:** Free OCR, AI jitna accurate nahi hota — especially amounts
(numbers) ke liye. Har entry ko download se pehle check/correct zaroor
karo. App har row ke saath raw OCR text aur "amount in words" wali line
bhi dikhata hai taaki galat number ko sahi karna easy ho.

## Setup (do cheezein install karni hongi)

### 1) Tesseract OCR program install karo (ye Python package se alag hai)

**Windows:**
1. Ye link kholo: https://github.com/UB-Mannheim/tesseract/wiki
2. Latest installer (.exe) download karo aur install karo (default location
   pe hi install hone do: `C:\Program Files\Tesseract-OCR\`)

**Mac:** terminal me `brew install tesseract`

**Linux:** terminal me `sudo apt install tesseract-ocr`

### 2) Python libraries install karo

Terminal me app wale folder me jaake:

```
pip install -r requirements.txt
```

## Chalao

```
streamlit run app.py
```

Browser khud khul jayega (localhost).

Agar Windows pe error aaye "Tesseract not found" jaisa, to app.py ke
upar wale hisse me dekho — waha ek line hai jo Tesseract ka path set karti
hai. Agar tumne default location pe install kiya hai to automatically
kaam kar jayega. Agar alag jagah install kiya hai, terminal me ye chalao
(apna path daal ke):

```
set TESSERACT_PATH=C:\Program Files\Tesseract-OCR\tesseract.exe
streamlit run app.py
```

## Use kaise karein

1. Ek ya multiple invoice files upload karo (PDF/JPG/PNG).
2. "Extract entries" dabao.
3. Table me values carefully check karo — GSTIN, amounts especially
   (OCR inme galti kar sakta hai). Neeche "Raw OCR reference" expand
   karke original text/amount-in-words dekh sakte ho agar koi field
   khaali ya galat lage.
4. Table me seedha cell edit kar sakte ho (double-click).
5. "Download Excel" se final .xlsx file mil jayegi.

## Free vs Paid — quick note

Ye free version regex/pattern-matching se kaam karta hai, isliye clean
printed invoices pe theek chalta hai but complex table layouts ya
handwritten notes (jaise "Paid" stamp) pe struggle karta hai. Agar
kabhi accuracy important ho jaye (jaise bulk processing ke liye), to
AI-powered (Claude API) version bhi maine banaya tha — bahut kam cost
pe (~1-2 paise/invoice) zyada accurate result deta hai. Bata dena agar
wo version chahiye future me.
