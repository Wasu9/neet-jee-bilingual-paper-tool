"""
NEET / JEE Bilingual Paper Maker v2.1
English → Hindi side-by-side Word generator for DTP
Compatible with Python 3.12+ / Streamlit Cloud (no googletrans / cgi issue)
"""

import streamlit as st
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import pdfplumber
import io
import re
import time
import traceback
import requests
from typing import List, Tuple, Dict

st.set_page_config(
    page_title="NEET/JEE Bilingual Paper Maker",
    page_icon="📝",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.title("📝 NEET / JEE Bilingual Paper Maker")
st.caption("English → Hindi side-by-side Word | Math protected | Streamlit Cloud compatible")

def google_translate(text: str, source: str = "en", target: str = "hi", max_retries: int = 4) -> str:
    if not text or not text.strip():
        return text

    protected, placeholders = protect_math(text)
    url = "https://translate.googleapis.com/translate_a/single"
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    }

    for attempt in range(max_retries):
        try:
            if len(protected) > 4200:
                chunks = split_long_text(protected, 4000)
                results = []
                for chunk in chunks:
                    params = {"client": "gtx", "sl": source, "tl": target, "dt": "t", "q": chunk}
                    r = requests.get(url, params=params, headers=headers, timeout=15)
                    r.raise_for_status()
                    data = r.json()
                    translated_chunk = "".join([item[0] for item in data[0] if item[0]])
                    results.append(translated_chunk)
                    time.sleep(0.3)
                translated = "".join(results)
            else:
                params = {"client": "gtx", "sl": source, "tl": target, "dt": "t", "q": protected}
                r = requests.get(url, params=params, headers=headers, timeout=15)
                r.raise_for_status()
                data = r.json()
                translated = "".join([item[0] for item in data[0] if item[0]])

            return restore_math(translated, placeholders)

        except Exception:
            wait = (1.5 ** attempt) + 0.4
            if attempt < max_retries - 1:
                time.sleep(wait)
            else:
                return text
    return text


def split_long_text(text: str, max_len: int = 4000) -> List[str]:
    if len(text) <= max_len:
        return [text]
    chunks = []
    current = ""
    for line in text.split("\n"):
        if len(current) + len(line) + 1 < max_len:
            current += line + "\n"
        else:
            if current:
                chunks.append(current.strip())
            current = line + "\n"
    if current.strip():
        chunks.append(current.strip())
    return chunks


def protect_math(text: str) -> Tuple[str, Dict[str, str]]:
    placeholders = {}
    counter = [0]
    def repl(match):
        key = f"__M{counter[0]}__"
        placeholders[key] = match.group(0)
        counter[0] += 1
        return key
    patterns = [
        r'\$\$[\s\S]+?\$\$', r'\$[^$\n]+\$', r'\\[a-zA-Z]+\{[^}]*\}', r'\\[a-zA-Z]+',
        r'[√≤≥≠≈≡∞∂∇∫∑∏πΔδθωαβγλμρσφψΩ±×÷→←↔⇒⇐⇔°′″†‡]',
        r'[a-zA-Z]̂', r'[a-zA-Z]⃗', r'[a-zA-Z]̅',
    ]
    for pat in patterns:
        text = re.sub(pat, repl, text)
    return text, placeholders


def restore_math(text: str, placeholders: Dict[str, str]) -> str:
    for key, value in placeholders.items():
        text = text.replace(key, value)
    return text


def extract_from_pdf(file) -> str:
    text_parts = []
    with pdfplumber.open(file) as pdf:
        for page in pdf.pages:
            t = page.extract_text()
            if t:
                text_parts.append(t)
    return "\n\n".join(text_parts)


def extract_from_docx(file) -> str:
    doc = Document(file)
    paras = [p.text for p in doc.paragraphs if p.text.strip()]
    for table in doc.tables:
        for row in table.rows:
            row_text = " | ".join(c.text.strip() for c in row.cells if c.text.strip())
            if row_text:
                paras.append(row_text)
    return "\n\n".join(paras)


def smart_split_exam_text(text: str) -> List[str]:
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    pattern = r'(?=\n\s*(?:\d{1,3}|Q\.?\s*\d{1,3})[\.\)]\s)'
    parts = re.split(pattern, text)
    chunks = []
    for p in parts:
        p = p.strip()
        if not p:
            continue
        if len(p) > 1800:
            chunks.extend(split_long_text(p, 1600))
        else:
            chunks.append(p)
    if len(chunks) <= 2 and len(text) > 2000:
        chunks = [c.strip() for c in text.split("\n\n") if c.strip()]
    return chunks


def set_run_font(run, font_name: str = "Mangal", size: int = 10, bold: bool = False):
    run.font.name = font_name
    run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
    run.font.size = Pt(size)
    run.bold = bold


def create_bilingual_docx(eng_chunks: List[str], hin_chunks: List[str],
                          title: str = "NEET / JEE Bilingual Question Paper") -> bytes:
    doc = Document()
    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(1.0)
    section.bottom_margin = Cm(1.0)
    section.left_margin = Cm(1.0)
    section.right_margin = Cm(1.0)

    tp = doc.add_paragraph()
    r = tp.add_run(title)
    r.bold = True
    r.font.size = Pt(14)
    tp.alignment = WD_ALIGN_PARAGRAPH.CENTER

    sp = doc.add_paragraph()
    r2 = sp.add_run("English  |  हिन्दी")
    r2.font.size = Pt(11)
    sp.alignment = WD_ALIGN_PARAGRAPH.CENTER

    table = doc.add_table(rows=1, cols=2)
    table.style = "Table Grid"
    table.autofit = False
    for cell in table.columns[0].cells:
        cell.width = Cm(9.2)
    for cell in table.columns[1].cells:
        cell.width = Cm(9.2)

    hdr = table.rows[0].cells
    hdr[0].text = "English"
    hdr[1].text = "हिन्दी"
    for cell in hdr:
        for p in cell.paragraphs:
            for run in p.runs:
                run.bold = True
                run.font.size = Pt(11)
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        shading = OxmlElement("w:shd")
        shading.set(qn("w:fill"), "D6EAF8")
        cell._tc.get_or_add_tcPr().append(shading)

    n = max(len(eng_chunks), len(hin_chunks))
    for i in range(n):
        eng = eng_chunks[i] if i < len(eng_chunks) else ""
        hin = hin_chunks[i] if i < len(hin_chunks) else ""
        row = table.add_row()
        p0 = row.cells[0].paragraphs[0]
        run0 = p0.add_run(eng)
        run0.font.name = "Times New Roman"
        run0.font.size = Pt(9)
        p0.paragraph_format.space_after = Pt(3)
        p1 = row.cells[1].paragraphs[0]
        run1 = p1.add_run(hin)
        set_run_font(run1, "Mangal", 9)
        p1.paragraph_format.space_after = Pt(3)

    doc.add_paragraph()
    note = doc.add_paragraph()
    nr = note.add_run("Note: Math equations & special characters are preserved. Please review NCERT scientific terminology. Generated for DTP use.")
    nr.font.size = Pt(8)
    nr.italic = True

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


with st.sidebar:
    st.header("⚙️ Settings")
    input_mode = st.radio("Input Type", ["Paste Text", "Upload DOCX", "Upload PDF"], index=2)
    st.markdown("---")
    st.markdown("**v2.1** · No googletrans · Python 3.14 ready")

english_text = ""

if input_mode == "Paste Text":
    english_text = st.text_area("English content paste karein:", height=320)
elif input_mode == "Upload DOCX":
    uploaded = st.file_uploader("English DOCX", type=["docx"])
    if uploaded:
        with st.spinner("Extracting..."):
            english_text = extract_from_docx(uploaded)
            st.success(f"{len(english_text)} characters")
elif input_mode == "Upload PDF":
    uploaded = st.file_uploader("English PDF", type=["pdf"])
    if uploaded:
        with st.spinner("Extracting..."):
            english_text = extract_from_pdf(uploaded)
            st.success(f"{len(english_text)} characters")

if english_text.strip():
    st.markdown("---")
    st.text_area("Preview", english_text, height=180, disabled=True)

    if st.button("🚀 Translate & Generate Bilingual Word", type="primary", use_container_width=True):
        start = time.time()
        status = st.empty()
        progress = st.progress(0)
        try:
            status.info("Splitting questions...")
            eng_chunks = smart_split_exam_text(english_text)
            progress.progress(0.1)
            status.info(f"Translating {len(eng_chunks)} parts...")
            hin_chunks = []
            for i, chunk in enumerate(eng_chunks):
                hin = google_translate(chunk)
                hin_chunks.append(hin)
                progress.progress(0.1 + 0.8 * (i + 1) / len(eng_chunks))
                status.info(f"Translating... {i+1}/{len(eng_chunks)}")
            status.info("Creating Word file...")
            docx_bytes = create_bilingual_docx(eng_chunks, hin_chunks)
            progress.progress(1.0)
            status.success(f"✅ Done in {time.time()-start:.1f}s")
            st.download_button("📥 Download Bilingual DOCX", data=docx_bytes,
                               file_name="bilingual_neet_jee_paper.docx",
                               mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                               use_container_width=True)
        except Exception:
            status.error("Error")
            st.code(traceback.format_exc())
