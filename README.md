# NEET / JEE Bilingual Paper Maker (v2.0 Stable)

**English → Hindi Side-by-Side Word Generator for DTP & Paper Setters**

Upload ready-made English NEET/JEE paper → Get clean bilingual DOCX  
(Left: English | Right: Hindi) in table format.

## What’s New in v2.0
- ✅ **Reliable translation** (googletrans + retry + exponential backoff)
- ✅ Smart question-wise splitting (detects `1. `, `12. `, `Q1.` etc.)
- ✅ Math, equations, vectors, special symbols **protected**
- ✅ Progress bar + clear status messages
- ✅ Better Hindi font handling (Mangal)
- ✅ Handles long papers better

## How to Run (Local)

```bash
git clone https://github.com/YOUR_USERNAME/neet-jee-bilingual-paper-tool.git
cd neet-jee-bilingual-paper-tool
pip install -r requirements.txt
streamlit run app.py
```

Browser opens at `http://localhost:8501`

## Deploy Free Website

### Streamlit Cloud (Recommended)
1. Push this repo to GitHub
2. Go to → https://share.streamlit.io
3. Connect GitHub → Select this repo → Deploy
4. Public URL mil jayega

## Usage Tips (DTP Operators)

1. **Best Input**
   - Text-based PDF (not scanned)
   - Or DOCX / plain text

2. **Math Protection**
   - Equations written as `$E = mc^2$` or `$$...$$` stay untouched
   - Vectors, √, π, α, β etc. are protected

3. **After Download**
   - Open in Word / LibreOffice
   - Check NCERT scientific terms (force → बल, acceleration → त्वरण etc. usually correct)
   - Insert diagrams / graphs manually (tool extracts only text)

4. **Long Papers**
   - 20-25 page paper usually takes 4–10 minutes
   - Progress bar dikhega

## Project Structure
```
neet-jee-bilingual-paper-tool/
├── app.py
├── requirements.txt
├── README.md
└── .gitignore
```

## Limitations
- Graphs / images / complex figures not extracted (text only)
- Very heavy rate-limit from Google can still happen (retry automatically tries 4 times)
- NCERT exact terminology should be reviewed once

## License
MIT – Free for personal & commercial use (coaching institutes welcome)

---
Made for Indian coaching teachers & DTP operators.
