# NEET / JEE Bilingual Paper Maker

**English → Hindi (NCERT-style) Side-by-Side Word Generator for DTP & Paper Setters**

Upload ready-made English NEET/JEE papers → Auto translate to Hindi → Get clean bilingual DOCX  
(Left column: English | Right column: Hindi) in table format, ready for printing / further DTP.

## Features

- ✅ Upload **DOCX**, **PDF** (text-based) or paste text
- ✅ Auto translate English → Hindi using Google Translate
- ✅ **Math equations, LaTeX (`$...$`, `$$...$$`), chemical formulas preserved** (not translated)
- ✅ Clean **two-column table** layout in Word
- ✅ A4 page, narrow margins, professional formatting
- ✅ Paragraph-wise translation for better accuracy
- ✅ Free & open source

## How to Run Locally

### 1. Clone the repo
```bash
git clone https://github.com/YOUR_USERNAME/neet-jee-bilingual-paper-tool.git
cd neet-jee-bilingual-paper-tool
```

### 2. Install dependencies
```bash
pip install -r requirements.txt
```

### 3. Run the app
```bash
streamlit run app.py
```

Browser automatically opens at `http://localhost:8501`

## Deploy as Website (Free)

### Option A: Streamlit Community Cloud (Recommended)
1. Push this repo to your GitHub
2. Go to [share.streamlit.io](https://share.streamlit.io)
3. Connect your GitHub account
4. Select this repository → Deploy
5. Get a public URL like `https://your-app.streamlit.app`

### Option B: Hugging Face Spaces
1. Create a new Space (Streamlit SDK)
2. Upload the files
3. It will run automatically

## Tips for Best Results (Important for DTP)

1. **Math & Equations**  
   Write equations in LaTeX style before uploading:  
   `The energy is $E = mc^2$` or `$$ \int_0^\infty e^{-x} dx $$`  
   These will stay untouched.

2. **NCERT Scientific Terms**  
   Google Translate is good but not perfect for pure NCERT terminology.  
   After download, quickly search-replace important terms if needed (e.g. "force" → "बल", "acceleration" → "त्वरण" etc. are usually correct).

3. **Graphs / Diagrams / Images**  
   These are not auto-copied. Extract images separately and place them manually in the final Word file.

4. **Long Papers**  
   Translation takes time (depends on length). Progress bar is shown.

5. **PDF Quality**  
   Text-based PDFs work best. Scanned/image PDFs need OCR first (use Adobe / online OCR).

## Project Structure
```
neet-jee-bilingual-paper-tool/
├── app.py                 # Main Streamlit application
├── requirements.txt
└── README.md
```

## Limitations

- Translation quality depends on Google Translate (good for general + science, but review critical terms)
- Complex multi-column original layouts may need manual adjustment after generation
- Graphs/images not extracted automatically
- Very long documents may hit rate limits occasionally (retry works)

## Future Improvements (Contributions Welcome)

- [ ] Better NCERT scientific dictionary / post-processing
- [ ] Support for images extraction & placement
- [ ] Option for word-by-word vs natural translation
- [ ] Direct LaTeX input support
- [ ] Multiple language pairs
- [ ] Batch processing

## License

MIT License – free for personal & commercial use (coaching institutes welcome).

---

Made for Indian coaching teachers & DTP operators who make bilingual NEET/JEE papers daily.
""")