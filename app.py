"""
NEET / JEE Bilingual Paper Maker v3.0
Professional side-by-side format (like coaching institute papers)
English Left | Hindi Right  — per question
"""

import streamlit as st
from docx import Document
from docx.shared import Pt, Cm, RGBColor
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
st.caption("Professional format • English Left | Hindi Right • Math protected • v3.0")

# ============================================================
# TRANSLATOR
# ============================================================
def google_translate(text: str, source: str = "en", target: str = "hi", max_retries: int = 4) -> str:
    if not text or not text.strip():
        return text

    protected, placeholders = protect_math(text)
    url = "https://translate.googleapis.com/translate_a/single"
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
    }

    for attempt in range(max_retries):
        try:
            if len(protected) > 4000:
                chunks = split_long_text(protected, 3800)
                results = []
                for chunk in chunks:
                    params = {"client": "gtx", "sl": source, "tl": target, "dt": "t", "q": chunk}
                    r = requests.get(url, params=params, headers=headers, timeout=15)
                    r.raise_for_status()
                    data = r.json()
                    results.append("".join([item[0] for item in data[0] if item[0]]))
                    time.sleep(0.28)
                translated = "".join(results)
            else:
                params = {"client": "gtx", "sl": source, "tl": target, "dt": "t", "q": protected}
                r = requests.get(url, params=params, headers=headers, timeout=15)
                r.raise_for_status()
                data = r.json()
                translated = "".join([item[0] for item in data[0] if item[0]])

            return restore_math(translated, placeholders)
        except Exception:
            if attempt < max_retries - 1:
                time.sleep((1.6 ** attempt) + 0.3)
            else:
                return text
    return text


def split_long_text(text: str, max_len: int = 3800) -> List[str]:
    if len(text) <= max_len:
        return [text]
    chunks, current = [], ""
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
    def repl(m):
        key = f"__M{counter[0]}__"
        placeholders[key] = m.group(0)
        counter[0] += 1
        return key
    patterns = [
        r'\$\$[\s\S]+?\$\$', r'\$[^$\n]+\$',
        r'\\[a-zA-Z]+\{[^}]*\}', r'\\[a-zA-Z]+',
        r'[√≤≥≠≈≡∞∂∇∫∑∏πΔδθωαβγλμρσφψΩ±×÷→←↔⇒⇐⇔°′″†‡]',
        r'[a-zA-Z]̂', r'[a-zA-Z]⃗', r'[a-zA-Z]̅',
    ]
    for pat in patterns:
        text = re.sub(pat, repl, text)
    return text, placeholders


def restore_math(text: str, placeholders: Dict[str, str]) -> str:
    for k, v in placeholders.items():
        text = text.replace(k, v)
    return text


# ============================================================
# EXTRACTION + SMART SPLIT
# ============================================================
def extract_from_pdf(file) -> str:
    parts = []
    with pdfplumber.open(file) as pdf:
        for page in pdf.pages:
            t = page.extract_text()
            if t:
                parts.append(t)
    return "\n\n".join(parts)


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
        if p:
            if len(p) > 2000:
                chunks.extend(split_long_text(p, 1800))
            else:
                chunks.append(p)
    if len(chunks) <= 2 and len(text) > 1500:
        chunks = [c.strip() for c in text.split("\n\n") if c.strip()]
    return chunks


# ============================================================
# PROFESSIONAL DOCX (per-question two-column)
# ============================================================
def set_cell_border(cell, **kwargs):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for edge in ('top', 'left', 'bottom', 'right'):
        edge_data = kwargs.get(edge)
        if edge_data:
            tag = OxmlElement(f'w:{edge}')
            tag.set(qn('w:val'), edge_data.get('val', 'single'))
            tag.set(qn('w:sz'), str(edge_data.get('sz', 4)))
            tag.set(qn('w:color'), edge_data.get('color', '000000'))
            tcBorders.append(tag)
    tcPr.append(tcBorders)


def set_run_font(run, name="Mangal", size=9, bold=False):
    run.font.name = name
    run._element.rPr.rFonts.set(qn('w:eastAsia'), name)
    run.font.size = Pt(size)
    run.bold = bold


def create_professional_bilingual(eng_chunks: List[str], hin_chunks: List[str],
                                  title: str = "NEET / JEE Bilingual Question Paper") -> bytes:
    doc = Document()
    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(0.9)
    section.bottom_margin = Cm(0.9)
    section.left_margin = Cm(1.0)
    section.right_margin = Cm(1.0)

    # Title
    tp = doc.add_paragraph()
    r = tp.add_run(title)
    r.bold = True
    r.font.size = Pt(13)
    r.font.name = "Arial"
    tp.alignment = WD_ALIGN_PARAGRAPH.CENTER

    sp = doc.add_paragraph()
    r2 = sp.add_run("English  |  हिन्दी")
    r2.font.size = Pt(10)
    sp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sp.paragraph_format.space_after = Pt(6)

    # Each question as separate two-column table
    for eng, hin in zip(eng_chunks, hin_chunks):
        table = doc.add_table(rows=1, cols=2)
        table.autofit = False
        table.allow_autofit = False

        for cell in table.columns[0].cells:
            cell.width = Cm(9.3)
        for cell in table.columns[1].cells:
            cell.width = Cm(9.3)

        for cell in table.rows[0].cells:
            set_cell_border(cell,
                            top={"val": "single", "sz": 4, "color": "AAAAAA"},
                            bottom={"val": "single", "sz": 4, "color": "AAAAAA"},
                            left={"val": "single", "sz": 4, "color": "AAAAAA"},
                            right={"val": "single", "sz": 4, "color": "AAAAAA"})

        # English
        cell0 = table.rows[0].cells[0]
        p0 = cell0.paragraphs[0]
        run0 = p0.add_run(eng)
        run0.font.name = "Times New Roman"
        run0.font.size = Pt(9)
        p0.paragraph_format.space_after = Pt(2)
        p0.paragraph_format.space_before = Pt(2)

        # Hindi
        cell1 = table.rows[0].cells[1]
        p1 = cell1.paragraphs[0]
        run1 = p1.add_run(hin)
        set_run_font(run1, "Mangal", 9)
        p1.paragraph_format.space_after = Pt(2)
        p1.paragraph_format.space_before = Pt(2)

        # Small gap
        gap = doc.add_paragraph()
        gap.paragraph_format.space_before = Pt(1)
        gap.paragraph_format.space_after = Pt(1)

    # Footer
    note = doc.add_paragraph()
    nr = note.add_run(
        "Note: Math & special characters preserved. Review NCERT terms if needed. Generated for DTP use."
    )
    nr.font.size = Pt(7)
    nr.italic = True
    nr.font.color.rgb = RGBColor(100, 100, 100)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


# ============================================================
# UI
# ============================================================
with st.sidebar:
    st.header("⚙️ Settings")
    input_mode = st.radio("Input Type", ["Paste Text", "Upload DOCX", "Upload PDF"], index=2)
    st.markdown("---")
    st.markdown("""
**v3.0 Features**
- Per-question side-by-side
- Professional exam look
- Math protected
- NEET / JEE optimized
    """)
    st.markdown("---")
    st.caption("Format inspired by standard coaching bilingual papers")

english_text = ""

if input_mode == "Paste Text":
    english_text = st.text_area("English paper content:", height=300)
elif input_mode == "Upload DOCX":
    uploaded = st.file_uploader("English DOCX", type=["docx"])
    if uploaded:
        with st.spinner("Extracting..."):
            english_text = extract_from_docx(uploaded)
            st.success(f"{len(english_text)} characters")
elif input_mode == "Upload PDF":
    uploaded = st.file_uploader("English PDF (text-based)", type=["pdf"])
    if uploaded:
        with st.spinner("Extracting from PDF..."):
            english_text = extract_from_pdf(uploaded)
            st.success(f"{len(english_text)} characters extracted")

if english_text.strip():
    st.markdown("---")
    with st.expander("English Preview (first 1200 chars)"):
        st.text(english_text[:1200] + ("..." if len(english_text) > 1200 else ""))

    if st.button("🚀 Translate & Generate Professional Bilingual Paper", type="primary", use_container_width=True):
        start = time.time()
        status = st.empty()
        progress = st.progress(0)

        try:
            status.info("Step 1/3 → Detecting questions...")
            eng_chunks = smart_split_exam_text(english_text)
            progress.progress(0.08)

            status.info(f"Step 2/3 → Translating {len(eng_chunks)} parts to Hindi...")
            hin_chunks = []
            total = len(eng_chunks)
            for i, chunk in enumerate(eng_chunks):
                hin = google_translate(chunk)
                hin_chunks.append(hin)
                progress.progress(0.08 + 0.82 * (i + 1) / total)
                status.info(f"Translating... {i+1}/{total}")

            status.info("Step 3/3 → Creating professional Word file...")
            progress.progress(0.95)

            docx_bytes = create_professional_bilingual(eng_chunks, hin_chunks)

            progress.progress(1.0)
            elapsed = time.time() - start
            status.success(f"✅ Done in {elapsed:.1f} seconds | {len(eng_chunks)} questions/parts")

            st.download_button(
                "📥 Download Professional Bilingual DOCX",
                data=docx_bytes,
                file_name="bilingual_neet_jee_professional.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )

            with st.expander("Quick Preview (first 2 parts)"):
                for e, h in zip(eng_chunks[:2], hin_chunks[:2]):
                    col1, col2 = st.columns(2)
                    with col1:
                        st.markdown("**English**")
                        st.text(e[:220] + "...")
                    with col2:
                        st.markdown("**हिन्दी**")
                        st.text(h[:220] + "...")
                    st.markdown("---")

        except Exception:
            status.error("Error occurred")
            st.code(traceback.format_exc())
            st.warning("Agar rate-limit aaye to 30-40 second baad phir try karein.")

st.markdown("---")
st.caption("v3.0 • Professional per-question side-by-side format for NEET & JEE DTP")
