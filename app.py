import streamlit as st
from docx import Document
from docx.shared import Inches, Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from deep_translator import GoogleTranslator
import pdfplumber
import io
import re
import time
from typing import List, Tuple

st.set_page_config(
    page_title="NEET/JEE Bilingual Paper Maker",
    page_icon="📝",
    layout="wide"
)

st.title("📝 NEET / JEE Bilingual Paper Maker")
st.markdown("""
**English → Hindi (NCERT style) Side-by-Side Word Generator**

Upload English paper (DOCX / PDF / Text) → Auto translate → Download bilingual DOCX  
(Left: English | Right: Hindi) in clean table format.
""")

# ---------------- Helper Functions ----------------

def set_run_font(run, font_name="Mangal", size=11, bold=False):
    """Set font for Hindi support"""
    run.font.name = font_name
    run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
    run.font.size = Pt(size)
    run.bold = bold


def add_horizontal_line(paragraph):
    """Add a bottom border to a paragraph"""
    p = paragraph._p
    pPr = p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '6')
    bottom.set(qn('w:space'), '1')
    bottom.set(qn('w:color'), '000000')
    pBdr.append(bottom)
    pPr.append(pBdr)


def protect_math(text: str) -> Tuple[str, dict]:
    """
    Protect LaTeX math, equations, and special patterns from translation.
    Returns protected text and a dict of placeholders.
    """
    placeholders = {}
    counter = 0

    # Protect $...$ and $$...$$
    def repl_math(match):
        nonlocal counter
        key = f"__MATH_{counter}__"
        placeholders[key] = match.group(0)
        counter += 1
        return key

    text = re.sub(r'\$\$[^$]+\$\$', repl_math, text)
    text = re.sub(r'\$[^$]+\$', repl_math, text)

    # Protect common chemical formulas / equations patterns roughly
    text = re.sub(r'\\[a-zA-Z]+\{[^}]*\}', repl_math, text)

    return text, placeholders


def restore_math(text: str, placeholders: dict) -> str:
    for key, value in placeholders.items():
        text = text.replace(key, value)
    return text


def translate_text(text: str, target="hi") -> str:
    """Translate while protecting math"""
    if not text or not text.strip():
        return text

    protected, placeholders = protect_math(text)

    try:
        # Split long text to avoid limits
        max_len = 4500
        if len(protected) > max_len:
            parts = []
            current = ""
            for para in protected.split("\n"):
                if len(current) + len(para) < max_len:
                    current += para + "\n"
                else:
                    if current:
                        parts.append(current)
                    current = para + "\n"
            if current:
                parts.append(current)

            translated_parts = []
            for part in parts:
                try:
                    t = GoogleTranslator(source='en', target=target).translate(part)
                    translated_parts.append(t if t else part)
                except Exception:
                    translated_parts.append(part)
            translated = "\n".join(translated_parts)
        else:
            translated = GoogleTranslator(source='en', target=target).translate(protected)
            if not translated:
                translated = protected
    except Exception as e:
        st.warning(f"Translation error (kept original): {e}")
        translated = protected

    return restore_math(translated, placeholders)


def extract_text_from_pdf(file) -> str:
    text = ""
    with pdfplumber.open(file) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text()
            if page_text:
                text += page_text + "\n\n"
    return text.strip()


def extract_text_from_docx(file) -> str:
    doc = Document(file)
    paragraphs = []
    for para in doc.paragraphs:
        if para.text.strip():
            paragraphs.append(para.text)
    # Also tables
    for table in doc.tables:
        for row in table.rows:
            row_text = " | ".join(cell.text.strip() for cell in row.cells if cell.text.strip())
            if row_text:
                paragraphs.append(row_text)
    return "\n\n".join(paragraphs)


def create_bilingual_docx(english_paragraphs: List[str], hindi_paragraphs: List[str], title: str = "Bilingual Paper") -> bytes:
    doc = Document()

    # Set narrow margins for more content
    section = doc.sections[0]
    section.top_margin = Cm(1.2)
    section.bottom_margin = Cm(1.2)
    section.left_margin = Cm(1.2)
    section.right_margin = Cm(1.2)
    section.page_width = Cm(21.0)   # A4
    section.page_height = Cm(29.7)

    # Title
    title_para = doc.add_paragraph()
    title_run = title_para.add_run(title)
    title_run.bold = True
    title_run.font.size = Pt(16)
    title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER

    subtitle = doc.add_paragraph()
    sub_run = subtitle.add_run("English  |  हिन्दी (NCERT Style)")
    sub_run.font.size = Pt(11)
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph()

    # Create main table
    table = doc.add_table(rows=1, cols=2)
    table.style = 'Table Grid'
    table.autofit = False
    table.allow_autofit = False

    # Set column widths (roughly equal)
    for cell in table.columns[0].cells:
        cell.width = Cm(9.0)
    for cell in table.columns[1].cells:
        cell.width = Cm(9.0)

    # Header row
    header_cells = table.rows[0].cells
    header_cells[0].text = "English"
    header_cells[1].text = "हिन्दी"

    for cell in header_cells:
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                run.bold = True
                run.font.size = Pt(12)
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        # Light gray background
        shading = OxmlElement('w:shd')
        shading.set(qn('w:fill'), 'D9E2F3')
        cell._tc.get_or_add_tcPr().append(shading)

    # Content rows
    max_len = max(len(english_paragraphs), len(hindi_paragraphs))
    for i in range(max_len):
        row = table.add_row()
        eng = english_paragraphs[i] if i < len(english_paragraphs) else ""
        hin = hindi_paragraphs[i] if i < len(hindi_paragraphs) else ""

        # English cell
        cell0 = row.cells[0]
        p0 = cell0.paragraphs[0]
        run0 = p0.add_run(eng)
        run0.font.name = "Times New Roman"
        run0.font.size = Pt(10)
        p0.paragraph_format.space_after = Pt(4)

        # Hindi cell
        cell1 = row.cells[1]
        p1 = cell1.paragraphs[0]
        run1 = p1.add_run(hin)
        set_run_font(run1, "Mangal", 10)
        p1.paragraph_format.space_after = Pt(4)

    # Footer note
    doc.add_paragraph()
    note = doc.add_paragraph()
    note_run = note.add_run(
        "Note: Math equations, formulas and special characters are preserved as-is. "
        "Please review Hindi translation for NCERT-specific scientific terminology accuracy. "
        "Generated for DTP / Paper setting use."
    )
    note_run.font.size = Pt(8)
    note_run.italic = True

    # Save to bytes
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


# ---------------- UI ----------------

with st.sidebar:
    st.header("⚙️ Settings")
    input_type = st.radio("Input Type", ["Text / Paste", "Upload DOCX", "Upload PDF"])
    preserve_paragraphs = st.checkbox("Keep original paragraph breaks", value=True)
    st.markdown("---")
    st.markdown("**Tips for best results:**")
    st.markdown("""
    - Math ko `$...$` ya `$$...$$` me rakho (LaTeX)
    - Chemical formulas plain text me bhi theek rehte hain
    - Graphs/images alag se handle karne padenge
    - Final Hindi me NCERT terms manually check karo
    """)

english_text = ""

if input_type == "Text / Paste":
    english_text = st.text_area(
        "Paste English paper content here:",
        height=300,
        placeholder="Q1. A particle moves with velocity...\n\n(a) option1\n(b) option2\n..."
    )
elif input_type == "Upload DOCX":
    uploaded = st.file_uploader("Upload English DOCX", type=["docx"])
    if uploaded:
        with st.spinner("Extracting text from DOCX..."):
            english_text = extract_text_from_docx(uploaded)
            st.success("Text extracted successfully!")
            with st.expander("Preview extracted text"):
                st.text(english_text[:2000] + ("..." if len(english_text) > 2000 else ""))
elif input_type == "Upload PDF":
    uploaded = st.file_uploader("Upload English PDF", type=["pdf"])
    if uploaded:
        with st.spinner("Extracting text from PDF (text-based PDFs work best)..."):
            english_text = extract_text_from_pdf(uploaded)
            st.success("Text extracted!")
            with st.expander("Preview extracted text"):
                st.text(english_text[:2000] + ("..." if len(english_text) > 2000 else ""))

if english_text.strip():
    st.markdown("---")
    col1, col2 = st.columns(2)

    with col1:
        st.subheader("English Preview")
        st.text_area("English", english_text, height=250, disabled=True, label_visibility="collapsed")

    if st.button("🚀 Translate to Hindi & Generate Bilingual Paper", type="primary", use_container_width=True):
        with st.spinner("Translating... (this may take a minute for long papers)"):
            # Split into paragraphs
            if preserve_paragraphs:
                eng_paras = [p.strip() for p in english_text.split("\n\n") if p.strip()]
                if len(eng_paras) == 1:
                    # fallback to single newlines
                    eng_paras = [p.strip() for p in english_text.split("\n") if p.strip()]
            else:
                eng_paras = [english_text.strip()]

            hindi_paras = []
            progress = st.progress(0)
            status = st.empty()

            for i, para in enumerate(eng_paras):
                status.text(f"Translating paragraph {i+1}/{len(eng_paras)}...")
                hin = translate_text(para)
                hindi_paras.append(hin)
                progress.progress((i + 1) / len(eng_paras))
                time.sleep(0.3)  # avoid Google rate limits

            status.text("Creating Word document...")

            docx_bytes = create_bilingual_docx(
                eng_paras,
                hindi_paras,
                title="NEET / JEE Bilingual Question Paper"
            )

            progress.empty()
            status.empty()

            st.success("✅ Bilingual paper ready!")

            st.download_button(
                label="📥 Download Bilingual DOCX",
                data=docx_bytes,
                file_name="bilingual_neet_jee_paper.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )

            # Side by side preview
            st.markdown("### Side-by-Side Preview (first few paragraphs)")
            preview_table = ""
            for e, h in zip(eng_paras[:6], hindi_paras[:6]):
                preview_table += f"| {e[:120]}{'...' if len(e)>120 else ''} | {h[:120]}{'...' if len(h)>120 else ''} |\n"
            st.markdown("| English | हिन्दी |\n|---|---|\n" + preview_table)

st.markdown("---")
st.caption("Made for DTP / Coaching paper setters | Math & special chars preserved | Review NCERT terms manually")
