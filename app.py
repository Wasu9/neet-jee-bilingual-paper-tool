"""
NEET / JEE Bilingual Paper Maker
English → Hindi (side-by-side Word) for DTP / Paper Setting
Stable version with retry, better splitting & math protection
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
from typing import List, Tuple, Dict

# ---------------- Page Config ----------------
st.set_page_config(
    page_title="NEET/JEE Bilingual Paper Maker",
    page_icon="📝",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.title("📝 NEET / JEE Bilingual Paper Maker")
st.caption("English → Hindi side-by-side Word generator for DTP | Math & special characters protected")

# ---------------- Translator (with retry) ----------------
@st.cache_resource
def get_translator():
    from googletrans import Translator
    return Translator()

def translate_with_retry(text: str, max_retries: int = 4) -> str:
    """Translate English → Hindi with exponential backoff"""
    if not text or not text.strip():
        return text

    translator = get_translator()
    protected, placeholders = protect_math(text)

    for attempt in range(max_retries):
        try:
            # googletrans has limit around 5k chars
            if len(protected) > 4500:
                # Split by sentences / lines
                chunks = split_long_text(protected, 4000)
                results = []
                for chunk in chunks:
                    res = translator.translate(chunk, src='en', dest='hi')
                    results.append(res.text if res and res.text else chunk)
                    time.sleep(0.35)
                translated = " ".join(results)
            else:
                res = translator.translate(protected, src='en', dest='hi')
                translated = res.text if res and res.text else protected

            return restore_math(translated, placeholders)

        except Exception as e:
            wait = (2 ** attempt) + 0.5
            if attempt < max_retries - 1:
                time.sleep(wait)
            else:
                # Last attempt failed → return original
                return text
    return text


def split_long_text(text: str, max_len: int = 4000) -> List[str]:
    """Split long text into chunks without breaking mid-sentence if possible"""
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


# ---------------- Math / Symbol Protection ----------------
def protect_math(text: str) -> Tuple[str, Dict[str, str]]:
    """Protect LaTeX, equations, special math symbols from translation"""
    placeholders = {}
    counter = [0]

    def repl(match):
        key = f"__MATH{counter[0]}__"
        placeholders[key] = match.group(0)
        counter[0] += 1
        return key

    # Order matters – more specific first
    patterns = [
        r'\$\$[\s\S]+?\$\$',                    # $$ ... $$
        r'\$[^$\n]+\$',                         # $ ... $
        r'\\[a-zA-Z]+\{[^}]*\}',                # \frac{ }{ }, \sqrt{} etc
        r'\\[a-zA-Z]+',                         # \alpha, \beta, \pi
        r'[√≤≥≠≈≡∞∂∇∫∑∏πΔδθωαβγλμρσφψΩ±×÷→←↔⇒⇐⇔°′″†‡]',
        r'\b[A-Za-z]̂\b',                       # î ĵ k̂
        r'[a-zA-Z]⃗|[a-zA-Z]̅',                 # vector / bar
    ]

    for pat in patterns:
        text = re.sub(pat, repl, text)

    return text, placeholders


def restore_math(text: str, placeholders: Dict[str, str]) -> str:
    for key, value in placeholders.items():
        text = text.replace(key, value)
    return text


# ---------------- Text Extraction ----------------
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


# ---------------- Smart Chunking for Exam Papers ----------------
def smart_split_exam_text(text: str) -> List[str]:
    """
    Split exam paper into logical units:
    - Header / instructions
    - Each question (starts with number.)
    """
    # Normalize
    text = text.replace("\r\n", "\n").replace("\r", "\n")

    # Split on question numbers like "1. ", "12. ", "Q1.", "Q.1" etc.
    pattern = r'(?=\n\s*(?:\d{1,3}|Q\.?\s*\d{1,3})[\.\)]\s)'
    parts = re.split(pattern, text)

    chunks = []
    for p in parts:
        p = p.strip()
        if not p:
            continue
        # Further split very long chunks
        if len(p) > 1800:
            sub = split_long_text(p, 1600)
            chunks.extend(sub)
        else:
            chunks.append(p)

    # If almost no splits happened, fallback to paragraph split
    if len(chunks) <= 2 and len(text) > 2000:
        chunks = [c.strip() for c in text.split("\n\n") if c.strip()]

    return chunks


# ---------------- DOCX Creation ----------------
def set_run_font(run, font_name: str = "Mangal", size: int = 10, bold: bool = False):
    run.font.name = font_name
    run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
    run.font.size = Pt(size)
    run.bold = bold


def create_bilingual_docx(
    eng_chunks: List[str],
    hin_chunks: List[str],
    title: str = "NEET / JEE Bilingual Question Paper"
) -> bytes:

    doc = Document()
    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(1.0)
    section.bottom_margin = Cm(1.0)
    section.left_margin = Cm(1.0)
    section.right_margin = Cm(1.0)

    # Title
    tp = doc.add_paragraph()
    r = tp.add_run(title)
    r.bold = True
    r.font.size = Pt(14)
    tp.alignment = WD_ALIGN_PARAGRAPH.CENTER

    sp = doc.add_paragraph()
    r2 = sp.add_run("English  |  हिन्दी")
    r2.font.size = Pt(11)
    sp.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Main table
    table = doc.add_table(rows=1, cols=2)
    table.style = "Table Grid"
    table.autofit = False

    for cell in table.columns[0].cells:
        cell.width = Cm(9.2)
    for cell in table.columns[1].cells:
        cell.width = Cm(9.2)

    # Header
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

    # Content
    n = max(len(eng_chunks), len(hin_chunks))
    for i in range(n):
        eng = eng_chunks[i] if i < len(eng_chunks) else ""
        hin = hin_chunks[i] if i < len(hin_chunks) else ""

        row = table.add_row()

        # English cell
        p0 = row.cells[0].paragraphs[0]
        run0 = p0.add_run(eng)
        run0.font.name = "Times New Roman"
        run0.font.size = Pt(9)
        p0.paragraph_format.space_after = Pt(3)

        # Hindi cell
        p1 = row.cells[1].paragraphs[0]
        run1 = p1.add_run(hin)
        set_run_font(run1, "Mangal", 9)
        p1.paragraph_format.space_after = Pt(3)

    # Footer note
    doc.add_paragraph()
    note = doc.add_paragraph()
    nr = note.add_run(
        "Note: Math equations & special characters are preserved. "
        "Please review NCERT scientific terminology for final accuracy. "
        "Generated for DTP / Paper setting use."
    )
    nr.font.size = Pt(8)
    nr.italic = True

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


# ---------------- UI ----------------
with st.sidebar:
    st.header("⚙️ Settings")
    input_mode = st.radio(
        "Input Type",
        ["Paste Text", "Upload DOCX", "Upload PDF"],
        index=2
    )
    st.markdown("---")
    st.markdown("**Tips for best results**")
    st.markdown("""
- Math ko `$E=mc^2$` ya `$$...$$` me likho
- PDF text-based hona chahiye (scanned + OCR pehle kar lo)
- Graphs / diagrams alag se daalne padenge
- Bahut lamba paper ho to 5-10 min lag sakte hain
    """)
    st.markdown("---")
    st.markdown("**Version:** 2.0 (Stable)")

english_text = ""

if input_mode == "Paste Text":
    english_text = st.text_area(
        "English paper content paste karein:",
        height=320,
        placeholder="1. For a body moving along x-axis...\n\n(1) speed (2) distance..."
    )
elif input_mode == "Upload DOCX":
    uploaded = st.file_uploader("English DOCX upload karein", type=["docx"])
    if uploaded:
        with st.spinner("DOCX se text nikal raha hoon..."):
            english_text = extract_from_docx(uploaded)
            st.success(f"Extracted {len(english_text)} characters")
            with st.expander("Preview (first 1500 chars)"):
                st.text(english_text[:1500] + ("..." if len(english_text) > 1500 else ""))
elif input_mode == "Upload PDF":
    uploaded = st.file_uploader("English PDF upload karein (text-based)", type=["pdf"])
    if uploaded:
        with st.spinner("PDF se text nikal raha hoon..."):
            english_text = extract_from_pdf(uploaded)
            st.success(f"Extracted {len(english_text)} characters from PDF")
            with st.expander("Preview (first 1500 chars)"):
                st.text(english_text[:1500] + ("..." if len(english_text) > 1500 else ""))

if english_text.strip():
    st.markdown("---")
    col1, col2 = st.columns([1, 1])
    with col1:
        st.subheader("English Preview")
        st.text_area("en_preview", english_text, height=220, disabled=True, label_visibility="collapsed")

    if st.button("🚀 Translate & Generate Bilingual Word", type="primary", use_container_width=True):
        start_time = time.time()
        status = st.empty()
        progress = st.progress(0)

        try:
            status.info("Step 1/3 → Smart splitting questions...")
            eng_chunks = smart_split_exam_text(english_text)
            progress.progress(0.1)

            status.info(f"Step 2/3 → Translating {len(eng_chunks)} parts to Hindi... (please wait)")
            hin_chunks = []
            total = len(eng_chunks)

            for i, chunk in enumerate(eng_chunks):
                hin = translate_with_retry(chunk)
                hin_chunks.append(hin)
                progress.progress(0.1 + 0.8 * (i + 1) / total)
                status.info(f"Translating... {i+1}/{total} parts done")

            status.info("Step 3/3 → Creating Word document...")
            progress.progress(0.95)

            docx_bytes = create_bilingual_docx(
                eng_chunks,
                hin_chunks,
                title="NEET / JEE Bilingual Question Paper"
            )

            progress.progress(1.0)
            elapsed = time.time() - start_time
            status.success(f"✅ Ready! Time taken: {elapsed:.1f} seconds | {len(eng_chunks)} parts translated")

            st.download_button(
                label="📥 Download Bilingual DOCX",
                data=docx_bytes,
                file_name="bilingual_neet_jee_paper.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )

            # Quick preview of first few
            with st.expander("Quick Side-by-Side Preview (first 4 parts)"):
                for e, h in zip(eng_chunks[:4], hin_chunks[:4]):
                    st.markdown(f"**EN:** {e[:180]}{'...' if len(e)>180 else ''}")
                    st.markdown(f"**HI:** {h[:180]}{'...' if len(h)>180 else ''}")
                    st.markdown("---")

        except Exception as e:
            status.error("Error aaya. Details below:")
            st.code(traceback.format_exc())
            st.warning("Agar rate-limit aaye to thoda wait karke phir se try karein.")

st.markdown("---")
st.caption("Made for Indian coaching DTP operators | Math protected | NCERT terms review recommended")
