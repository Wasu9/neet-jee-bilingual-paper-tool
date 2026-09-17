"""
NEET / JEE Bilingual Paper Maker v5.0 PROFESSIONAL
Clean UI • Better math • Clean formatting • Header separate • Questions clean
"""

import streamlit as st
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import io
import re
import time
import traceback
import requests
from typing import List, Tuple, Dict

# -------------------- PAGE CONFIG --------------------
st.set_page_config(
    page_title="NEET/JEE Bilingual Maker",
    page_icon="📝",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
    .main-title { font-size: 1.8rem; font-weight: 700; color: #1a365d; margin-bottom: 0.2rem; }
    .sub-title { color: #4a5568; font-size: 0.95rem; margin-bottom: 1.5rem; }
    .stButton>button { border-radius: 8px; font-weight: 600; }
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-title">📝 NEET / JEE Bilingual Paper Maker</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">Professional side-by-side format • Clean math • Header on page 1 • Questions from page 2</div>', unsafe_allow_html=True)

# -------------------- TRANSLATOR --------------------
def google_translate(text: str, max_retries: int = 4) -> str:
    if not text or not text.strip():
        return text
    protected, placeholders = protect_math(text)
    url = "https://translate.googleapis.com/translate_a/single"
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}

    for attempt in range(max_retries):
        try:
            if len(protected) > 3500:
                chunks = split_long_text(protected, 3400)
                results = []
                for chunk in chunks:
                    params = {"client": "gtx", "sl": "en", "tl": "hi", "dt": "t", "q": chunk}
                    r = requests.get(url, params=params, headers=headers, timeout=12)
                    r.raise_for_status()
                    data = r.json()
                    results.append("".join([item[0] for item in data[0] if item[0]]))
                    time.sleep(0.25)
                translated = "".join(results)
            else:
                params = {"client": "gtx", "sl": "en", "tl": "hi", "dt": "t", "q": protected}
                r = requests.get(url, params=params, headers=headers, timeout=12)
                r.raise_for_status()
                data = r.json()
                translated = "".join([item[0] for item in data[0] if item[0]])
            return restore_math(translated, placeholders)
        except Exception:
            if attempt < max_retries - 1:
                time.sleep(1.5 ** attempt + 0.3)
            else:
                return text
    return text


def split_long_text(text: str, max_len: int = 3400) -> List[str]:
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


# -------------------- MATH CLEANER --------------------
def improve_math(text: str) -> str:
    sup = str.maketrans("0123456789", "⁰¹²³⁴⁵⁶⁷⁸⁹")
    text = re.sub(r'([a-zA-Z])(\d)\b', lambda m: m.group(1) + m.group(2).translate(sup), text)
    text = re.sub(r'([a-zA-Z])\^(\d+)', lambda m: m.group(1) + m.group(2).translate(sup), text)

    replacements = [
        (r'\bsqrt\b', '√'),
        (r'\bpi\b', 'π'),
        (r'\bPi\b', 'π'),
        (r'\btheta\b', 'θ'),
        (r'\bTheta\b', 'θ'),
        (r'\balpha\b', 'α'),
        (r'\bbeta\b', 'β'),
        (r'\bgamma\b', 'γ'),
        (r'\bdelta\b', 'δ'),
        (r'\bmu\b', 'μ'),
        (r'\bsigma\b', 'σ'),
        (r'\bomega\b', 'ω'),
        (r'\blambda\b', 'λ'),
        (r'\brho\b', 'ρ'),
        (r'\bphi\b', 'φ'),
        (r'\bapprox\b', '≈'),
        (r'\binfinity\b', '∞'),
        (r'\bdegree\b', '°'),
    ]
    for pat, repl in replacements:
        text = re.sub(pat, repl, text, flags=re.IGNORECASE)

    text = re.sub(r'\b([ijk])\b(?=\s|$|[,.\)])', r'\1̂', text)
    return text


def protect_math(text: str) -> Tuple[str, Dict[str, str]]:
    text = improve_math(text)
    placeholders = {}
    counter = [0]
    def repl(m):
        key = f"__M{counter[0]}__"
        placeholders[key] = m.group(0)
        counter[0] += 1
        return key
    patterns = [
        r'\$\$[\s\S]+?\$\$', r'\$[^$\n]+\$',
        r'[√≤≥≠≈≡∞∂∇∫∑∏πΔδθωαβγλμρσφψΩ±×÷→←↔°′″]',
        r'[a-zA-Z]̂', r'[a-zA-Z]⃗',
        r'[⁰¹²³⁴⁵⁶⁷⁸⁹₀₁₂₃₄₅₆₇₈₉]',
    ]
    for pat in patterns:
        text = re.sub(pat, repl, text)
    return text, placeholders


def restore_math(text: str, placeholders: Dict[str, str]) -> str:
    for k, v in placeholders.items():
        text = text.replace(k, v)
    return text


# -------------------- EXTRACTION --------------------
def extract_from_pdf(file) -> str:
    import fitz
    pdf_bytes = file.read()
    file.seek(0)
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    texts = []
    for page in doc:
        t = page.get_text("text")
        if t.strip():
            texts.append(t)
    doc.close()
    return "\n\n".join(texts)


def extract_from_docx(file) -> str:
    doc = Document(file)
    paras = [p.text for p in doc.paragraphs if p.text.strip()]
    for table in doc.tables:
        for row in table.rows:
            row_text = " | ".join(c.text.strip() for c in row.cells if c.text.strip())
            if row_text:
                paras.append(row_text)
    return "\n\n".join(paras)


# -------------------- SMART SPLIT --------------------
def split_header_and_questions(text: str) -> Tuple[str, List[str]]:
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = re.sub(r'\n{3,}', '\n\n', text)

    match = re.search(r'(?m)^\s*(?:1[\.\)]\s+|Q\.?\s*1[\.\)]\s+)', text)
    if match:
        header = text[:match.start()].strip()
        qtext = text[match.start():].strip()
    else:
        header = ""
        qtext = text

    header = re.sub(r'Page\s*\d+\s*of\s*\d+', '', header, flags=re.I)
    header = re.sub(r'VERSION[- ]?\d+', '', header, flags=re.I)
    header = re.sub(r'\n{2,}', '\n', header).strip()

    parts = re.split(r'(?m)(?=^\s*(?:\d{1,3}|Q\.?\s*\d{1,3})[\.\)]\s+)', qtext)
    questions = []
    for p in parts:
        p = p.strip()
        if len(p) < 20:
            continue
        p = re.sub(r'\n\s*I\s*PUC.*Page\s*\d+.*$', '', p, flags=re.I)
        p = re.sub(r'\n\s*Page\s*\d+\s*$', '', p, flags=re.I)
        p = p.strip()
        if p:
            questions.append(p)

    return header, questions


# -------------------- DOCX BUILDER --------------------
def set_border(cell, color="CCCCCC", sz="4"):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    borders = OxmlElement('w:tcBorders')
    for edge in ('top', 'left', 'bottom', 'right'):
        tag = OxmlElement(f'w:{edge}')
        tag.set(qn('w:val'), 'single')
        tag.set(qn('w:sz'), sz)
        tag.set(qn('w:color'), color)
        borders.append(tag)
    tcPr.append(borders)


def set_hindi_font(run, size=9):
    run.font.name = "Mangal"
    run._element.rPr.rFonts.set(qn('w:eastAsia'), "Mangal")
    run.font.size = Pt(size)


def add_page_break(doc):
    p = doc.add_paragraph()
    run = p.add_run()
    run.add_break(WD_BREAK.PAGE)


def build_docx(header_eng, header_hin, questions_eng, questions_hin) -> bytes:
    doc = Document()
    sec = doc.sections[0]
    sec.page_width = Cm(21.0)
    sec.page_height = Cm(29.7)
    sec.top_margin = Cm(1.2)
    sec.bottom_margin = Cm(1.2)
    sec.left_margin = Cm(1.2)
    sec.right_margin = Cm(1.2)

    # Title
    t = doc.add_paragraph()
    r = t.add_run("NEET / JEE Bilingual Question Paper")
    r.bold = True
    r.font.size = Pt(14)
    r.font.name = "Arial"
    t.alignment = WD_ALIGN_PARAGRAPH.CENTER

    s = doc.add_paragraph()
    r2 = s.add_run("English  |  हिन्दी")
    r2.font.size = Pt(10)
    r2.font.color.rgb = RGBColor(80, 80, 80)
    s.alignment = WD_ALIGN_PARAGRAPH.CENTER
    s.paragraph_format.space_after = Pt(10)

    # Header Page 1
    if header_eng.strip():
        h = doc.add_paragraph()
        hr = h.add_run("HEADER & INSTRUCTIONS")
        hr.bold = True
        hr.font.size = Pt(11)
        h.alignment = WD_ALIGN_PARAGRAPH.CENTER

        table = doc.add_table(rows=1, cols=2)
        table.autofit = False
        table.columns[0].width = Cm(9.0)
        table.columns[1].width = Cm(9.0)

        for cell in table.rows[0].cells:
            set_border(cell, "666666", "8")

        c0 = table.rows[0].cells[0]
        p0 = c0.paragraphs[0]
        run0 = p0.add_run(header_eng)
        run0.font.name = "Times New Roman"
        run0.font.size = Pt(9)

        c1 = table.rows[0].cells[1]
        p1 = c1.paragraphs[0]
        run1 = p1.add_run(header_hin)
        set_hindi_font(run1, 9)

    add_page_break(doc)

    # Questions
    qt = doc.add_paragraph()
    qtr = qt.add_run("QUESTIONS  /  प्रश्न")
    qtr.bold = True
    qtr.font.size = Pt(12)
    qt.alignment = WD_ALIGN_PARAGRAPH.CENTER
    qt.paragraph_format.space_after = Pt(10)

    for eng, hin in zip(questions_eng, questions_hin):
        table = doc.add_table(rows=1, cols=2)
        table.autofit = False
        table.columns[0].width = Cm(9.0)
        table.columns[1].width = Cm(9.0)

        for cell in table.rows[0].cells:
            set_border(cell, "BBBBBB", "4")

        cell0 = table.rows[0].cells[0]
        p0 = cell0.paragraphs[0]
        run0 = p0.add_run(eng)
        run0.font.name = "Times New Roman"
        run0.font.size = Pt(9)
        p0.paragraph_format.space_before = Pt(3)
        p0.paragraph_format.space_after = Pt(3)

        cell1 = table.rows[0].cells[1]
        p1 = cell1.paragraphs[0]
        run1 = p1.add_run(hin)
        set_hindi_font(run1, 9)
        p1.paragraph_format.space_before = Pt(3)
        p1.paragraph_format.space_after = Pt(3)

        sp = doc.add_paragraph()
        sp.paragraph_format.space_before = Pt(4)
        sp.paragraph_format.space_after = Pt(4)

    note = doc.add_paragraph()
    nr = note.add_run("Note: Math cleaned with unicode. Please review NCERT terms & final formatting for print.")
    nr.font.size = Pt(7)
    nr.italic = True
    nr.font.color.rgb = RGBColor(120, 120, 120)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.getvalue()


# -------------------- UI --------------------
with st.sidebar:
    st.markdown("### ⚙️ Settings")
    input_mode = st.radio("Input", ["Upload PDF", "Upload DOCX", "Paste Text"], index=0)
    st.markdown("---")
    st.markdown("""
**v5.0 Professional**
- Clean UI
- Header → Page 1
- Questions → Page 2+
- Better math unicode
- Clean borders & spacing
    """)
    st.caption("Leave images — add manually in Word")

english_text = ""

if input_mode == "Paste Text":
    english_text = st.text_area("Paste English content", height=280)
elif input_mode == "Upload DOCX":
    up = st.file_uploader("DOCX file", type=["docx"])
    if up:
        with st.spinner("Extracting..."):
            english_text = extract_from_docx(up)
            st.success(f"{len(english_text)} characters")
elif input_mode == "Upload PDF":
    up = st.file_uploader("PDF file (text-based)", type=["pdf"])
    if up:
        with st.spinner("Extracting text from PDF..."):
            english_text = extract_from_pdf(up)
            st.success(f"{len(english_text)} characters extracted")

if english_text.strip():
    with st.expander("Preview extracted text"):
        st.text(english_text[:1400] + ("..." if len(english_text) > 1400 else ""))

    if st.button("🚀 Generate Professional Bilingual Paper", type="primary", use_container_width=True):
        start = time.time()
        status = st.empty()
        bar = st.progress(0)

        try:
            status.info("Separating header & questions...")
            header_eng, questions_eng = split_header_and_questions(english_text)
            bar.progress(0.12)

            status.info("Translating header...")
            header_hin = google_translate(header_eng) if header_eng else ""
            bar.progress(0.22)

            status.info(f"Translating {len(questions_eng)} questions...")
            questions_hin = []
            for i, q in enumerate(questions_eng):
                questions_hin.append(google_translate(q))
                bar.progress(0.22 + 0.65 * (i + 1) / max(len(questions_eng), 1))
                status.info(f"Translating {i+1}/{len(questions_eng)}")

            status.info("Building clean Word document...")
            bar.progress(0.92)

            docx_bytes = build_docx(header_eng, header_hin, questions_eng, questions_hin)
            bar.progress(1.0)

            elapsed = time.time() - start
            status.success(f"✅ Ready in {elapsed:.1f}s  •  {len(questions_eng)} questions")

            st.download_button(
                "📥 Download Professional DOCX",
                data=docx_bytes,
                file_name="bilingual_neet_jee_PROFESSIONAL.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )
        except Exception:
            status.error("Error")
            st.code(traceback.format_exc())

st.markdown("---")
st.caption("v5.0 Professional • Clean formatting • Better math • Images add manually")
