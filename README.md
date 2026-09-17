# NEET / JEE Bilingual Paper Maker (v2.1)

**English → Hindi Side-by-Side Word Generator for DTP**

## Fixed in v2.1
- ✅ Removed `googletrans` (was breaking on Python 3.13/3.14 because of removed `cgi` module)
- ✅ Pure `requests`-based Google Translate (works on Streamlit Cloud)
- ✅ Retry + backoff
- ✅ Math / symbol protection
- ✅ Smart question splitting

## Run Locally
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deploy on Streamlit Cloud
1. Push this repo to GitHub
2. Go to https://share.streamlit.io
3. Deploy → Public URL ready

## Usage
1. Upload PDF / DOCX or paste text
2. Click **Translate & Generate**
3. Download bilingual Word file

**Note:** Graphs/diagrams need to be added manually after download.
