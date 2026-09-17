"""
NEET / JEE Bilingual Paper Maker v4.0 PERFECT
- Header + Instructions on Page 1
- Questions start from Page 2
- Better math (superscript, fractions, vectors)
- Image/Graph extraction from PDF
- Professional coaching style layout
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

st.set_page_config(
    page_title="NEET/JEE Bilingual Paper Maker",
    page_icon="📝",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.title("📝 NEET / JEE Bilingual Paper Maker")
st.caption("v4.0 PERFECT • Header separate page • Proper math • Images extract • Professional format")

# ============================================================
# TRANSLATOR
# ============================================================
def google_translate(text: str, source: str = "en", target: str = "hi", max_retries: int = 4) -> str:
    if not text or not text.strip():
        return text
    protected, placeholders = protect_math(text)
    url = "https://translate.googleapis.com/translate_a/single"
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}

    for attempt in range(max_retries):
        try:
            if len(protected) > 3800:
                chunks = split_long_text(protected, 3600)
                results = []
                for chunk in chunks:
                    params = {"client": "gtx", "sl": source, "tl": target, "dt": "t", "q": chunk}
                    r = requests.get(url, params=params, headers=headers, timeout=15)
                    r.raise_for_status()
                    data = r.json()
                    results.append("".join([item[0] for item in data[0] if item[0]]))
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
            if attempt < max_retries - 1:
                time.sleep((1.7 ** attempt) + 0.4)
            else:
                return text
    return text


def split_long_text(text: str, max_len: int = 3600) -> List[str]:
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


# ============================================================
# MATH PROTECTION + BETTER UNICODE
# ============================================================
def protect_math(text: str) -> Tuple[str, Dict[str, str]]:
    placeholders = {}
    counter = [0]
    def repl(m):
        key = f"__M{counter[0]}__"
        placeholders[key] = m.group(0)
        counter[0] += 1
        return key

    text = improve_math_unicode(text)

    patterns = [
        r'\$\$[\s\S]+?\$\$', r'\$[^$\n]+\$',
        r'\\[a-zA-Z]+\{[^}]*\}', r'\\[a-zA-Z]+',
        r'[√≤≥≠≈≡∞∂∇∫∑∏πΔδθωαβγλμρσφψΩ±×÷→←↔⇒⇐⇔°′″†‡]',
        r'[a-zA-Z]̂', r'[a-zA-Z]⃗', r'[a-zA-Z]̅',
        r'[₀₁₂₃₄₅₆₇₈₉⁽⁾]', r'[⁰¹²³⁴⁵⁶⁷⁸⁹]',
    ]
    for pat in patterns:
        text = re.sub(pat, repl, text)
    return text, placeholders


def improve_math_unicode(text: str) -> str:
    """Convert common text math to better looking unicode"""
    sup_map = str.maketrans("0123456789+-=()", "⁰¹²³⁴⁵⁶⁷⁸⁹⁺⁻⁼⁽⁾")
    text = re.sub(r'([a-zA-Z])(\d)\b', lambda m: m.group(1) + m.group(2).translate(sup_map), text)
    text = re.sub(r'([a-zA-Z])\^(\d)', lambda m: m.group(1) + m.group(2).translate(sup_map), text)
    text = text.replace("sqrt", "√").replace("√ ", "√")
    text = re.sub(r'\b([ijk])\b(?=\s|$|,|\.|\)|])', r'\1̂', text)
    text = re.sub(r'\bpi\b', 'π', text)
    text = re.sub(r'\bPi\b', 'π', text)
    replacements = {
        'theta': 'θ', 'Theta': 'θ', 'alpha': 'α', 'beta': 'β', 'gamma': 'γ',
        'delta': 'δ', 'mu': 'μ', 'sigma': 'σ', 'omega': 'ω', 'lambda': 'λ',
        'rho': 'ρ', 'phi': 'φ', 'psi': 'ψ'
    }
    for k, v in replacements.items():
        text = re.sub(rf'\b{k}\b', v, text)
    return text


def restore_math(text: str, placeholders: Dict[str, str]) -> str:
    for k, v in placeholders.items():
        text = text.replace(k, v)
    return text


# ============================================================
# PDF EXTRACTION (text + images)
# ============================================================
def extract_text_and_images(pdf_file) -> Tuple[str, List[dict]]:
    import fitz  # pymupdf

    pdf_bytes = pdf_file.read()
    pdf_file.seek(0)

    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    full_text = []
    images = []

    for page_num, page in enumerate(doc):
        text = page.get_text("text")
        if text.strip():
            full_text.append(text)

        for img_index, img in enumerate(page.get_images(full=True)):
            try:
                xref = img[0]
                base_image = doc.extract_image(xref)
                image_bytes = base_image["image"]
                if len(image_bytes) > 5000:
                    images.append({
                        "page": page_num + 1,
                        "bytes": image_bytes,
                        "ext": base_image["ext"],
                        "index": img_index
                    })
            except Exception:
                continue

    doc.close()
    return "\n\n".join(full_text), images


def extract_from_docx(file) -> str:
    doc = Document(file)
    paras = [p.text for p in doc.paragraphs if p.text.strip()]
    for table in doc.tables:
        for row in table.rows:
            row_text = " | ".join(c.text.strip() for c in row.cells if c.text.strip())
            if row_text:
                paras.append(row_text)
    return "\n\n".join(paras)


# ============================================================
# SMART SPLIT: Header vs Questions
# ============================================================
def split_header_and_questions(text: str) -> Tuple[str, List[str]]:
    text = text.replace("\r\n", "\n").replace("\r", "\n")

    match = re.search(r'\n\s*(?:1[\.\)]\s+|Q\.?\s*1[\.\)]\s+)', text)
    if match:
        header = text[:match.start()].strip()
        questions_text = text[match.start():].strip()
    else:
        header = ""
        questions_text = text

    pattern = r'(?=\n\s*(?:\d{1,3}|Q\.?\s*\d{1,3})[\.\)]\s)'
    parts = re.split(pattern, questions_text)
    questions = []
    for p in parts:
        p = p.strip()
        if p and len(p) > 15:
            if len(p) > 2200:
                questions.extend(split_long_text(p, 2000))
            else:
                questions.append(p)

    return header, questions


# ============================================================
# DOCX CREATION - PROFESSIONAL
# ============================================================
def set_cell_border(cell, color="888888", sz="4"):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for edge in ('top', 'left', 'bottom', 'right'):
        tag = OxmlElement(f'w:{edge}')
        tag.set(qn('w:val'), 'single')
        tag.set(qn('w:sz'), sz)
        tag.set(qn('w:color'), color)
        tcBorders.append(tag)
    tcPr.append(tcBorders)


def set_run_font(run, name="Mangal", size=9, bold=False):
    run.font.name = name
    run._element.rPr.rFonts.set(qn('w:eastAsia'), name)
    run.font.size = Pt(size)
    run.bold = bold


def add_page_break(doc):
    p = doc.add_paragraph()
    run = p.add_run()
    run.add_break(WD_BREAK.PAGE)


def create_perfect_bilingual(
    header_eng: str,
    header_hin: str,
    questions_eng: List[str],
    questions_hin: List[str],
    images: List[dict] = None,
    title: str = "NEET / JEE Bilingual Question Paper"
) -> bytes:

    doc = Document()
    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(1.0)
    section.bottom_margin = Cm(1.0)
    section.left_margin = Cm(1.1)
    section.right_margin = Cm(1.1)

    # ========== PAGE 1 : HEADER + INSTRUCTIONS ==========
    tp = doc.add_paragraph()
    r = tp.add_run(title)
    r.bold = True
    r.font.size = Pt(14)
    r.font.name = "Arial"
    tp.alignment = WD_ALIGN_PARAGRAPH.CENTER

    sp = doc.add_paragraph()
    r2 = sp.add_run("English  |  हिन्दी")
    r2.font.size = Pt(11)
    sp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sp.paragraph_format.space_after = Pt(8)

    if header_eng.strip():
        htable = doc.add_table(rows=1, cols=2)
        htable.autofit = False
        for cell in htable.columns[0].cells:
            cell.width = Cm(9.2)
        for cell in htable.columns[1].cells:
            cell.width = Cm(9.2)

        for cell in htable.rows[0].cells:
            set_cell_border(cell, color="555555", sz="8")

        c0 = htable.rows[0].cells[0]
        p0 = c0.paragraphs[0]
        run0 = p0.add_run(header_eng)
        run0.font.name = "Times New Roman"
        run0.font.size = Pt(9)

        c1 = htable.rows[0].cells[1]
        p1 = c1.paragraphs[0]
        run1 = p1.add_run(header_hin)
        set_run_font(run1, "Mangal", 9)

    add_page_break(doc)

    # ========== PAGE 2 onwards : QUESTIONS ==========
    q_title = doc.add_paragraph()
    qr = q_title.add_run("QUESTIONS / प्रश्न")
    qr.bold = True
    qr.font.size = Pt(12)
    q_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    q_title.paragraph_format.space_after = Pt(8)

    for eng, hin in zip(questions_eng, questions_hin):
        table = doc.add_table(rows=1, cols=2)
        table.autofit = False
        for cell in table.columns[0].cells:
            cell.width = Cm(9.2)
        for cell in table.columns[1].cells:
            cell.width = Cm(9.2)

        for cell in table.rows[0].cells:
            set_cell_border(cell, color="AAAAAA", sz="4")

        cell0 = table.rows[0].cells[0]
        p0 = cell0.paragraphs[0]
        run0 = p0.add_run(eng)
        run0.font.name = "Times New Roman"
        run0.font.size = Pt(9)
        p0.paragraph_format.space_before = Pt(2)
        p0.paragraph_format.space_after = Pt(2)

        cell1 = table.rows[0].cells[1]
        p1 = cell1.paragraphs[0]
        run1 = p1.add_run(hin)
        set_run_font(run1, "Mangal", 9)
        p1.paragraph_format.space_before = Pt(2)
        p1.paragraph_format.space_after = Pt(2)

        gap = doc.add_paragraph()
        gap.paragraph_format.space_before = Pt(3)
        gap.paragraph_format.space_after = Pt(3)

    # Images at the end
    if images:
        add_page_break(doc)
        img_title = doc.add_paragraph()
        ir = img_title.add_run("EXTRACTED DIAGRAMS / GRAPHS (from original PDF)")
        ir.bold = True
        ir.font.size = Pt(11)
        img_title.alignment = WD_ALIGN_PARAGRAPH.CENTER

        for img_data in images[:12]:
            try:
                img_stream = io.BytesIO(img_data["bytes"])
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run = p.add_run()
                run.add_picture(img_stream, width=Cm(12))
                cap = doc.add_paragraph()
                cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
                cr = cap.add_run(f"Figure from page {img_data['page']}")
                cr.font.size = Pt(8)
                cr.italic = True
            except Exception:
                continue

    note = doc.add_paragraph()
    nr = note.add_run(
        "Note: Math improved with unicode. Review NCERT terms. Diagrams extracted separately. Generated for DTP."
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
    input_mode = st.radio("Input Type", ["Upload PDF (Recommended)", "Upload DOCX", "Paste Text"], index=0)
    extract_images = st.checkbox("Extract Graphs / Diagrams from PDF", value=True)
    st.markdown("---")
    st.markdown("""
**v4.0 PERFECT**
- Header + Instructions → Page 1
- Questions → New page
- Better math (² √ π θ î ĵ)
- Image extraction
- Clean professional look
    """)

english_text = ""
extracted_images = []

if input_mode == "Paste Text":
    english_text = st.text_area("English paper content:", height=300)
elif input_mode == "Upload DOCX":
    uploaded = st.file_uploader("English DOCX", type=["docx"])
    if uploaded:
        with st.spinner("Extracting text..."):
            english_text = extract_from_docx(uploaded)
            st.success(f"{len(english_text)} characters")
elif input_mode == "Upload PDF (Recommended)":
    uploaded = st.file_uploader("English PDF", type=["pdf"])
    if uploaded:
        with st.spinner("Extracting text + images from PDF..."):
            english_text, extracted_images = extract_text_and_images(uploaded)
            st.success(f"Text: {len(english_text)} chars | Images found: {len(extracted_images)}")
            if extracted_images:
                st.info(f"{len(extracted_images)} diagrams/graphs detected")

if english_text.strip():
    st.markdown("---")
    with st.expander("English Preview"):
        st.text(english_text[:1500] + ("..." if len(english_text) > 1500 else ""))

    if st.button("🚀 Generate Perfect Bilingual Paper", type="primary", use_container_width=True):
        start = time.time()
        status = st.empty()
        progress = st.progress(0)

        try:
            status.info("Step 1/4 → Separating Header & Questions...")
            header_eng, questions_eng = split_header_and_questions(english_text)
            progress.progress(0.1)

            status.info("Step 2/4 → Translating Header...")
            header_hin = google_translate(header_eng) if header_eng else ""
            progress.progress(0.2)

            status.info(f"Step 3/4 → Translating {len(questions_eng)} questions...")
            questions_hin = []
            for i, q in enumerate(questions_eng):
                hin = google_translate(q)
                questions_hin.append(hin)
                progress.progress(0.2 + 0.65 * (i + 1) / max(len(questions_eng), 1))
                status.info(f"Translating question {i+1}/{len(questions_eng)}")

            status.info("Step 4/4 → Creating professional Word file...")
            progress.progress(0.92)

            imgs = extracted_images if extract_images else []
            docx_bytes = create_perfect_bilingual(
                header_eng, header_hin,
                questions_eng, questions_hin,
                images=imgs
            )

            progress.progress(1.0)
            elapsed = time.time() - start
            status.success(f"✅ Done in {elapsed:.1f}s | Header + {len(questions_eng)} questions | {len(imgs)} images")

            st.download_button(
                "📥 Download Perfect Bilingual DOCX",
                data=docx_bytes,
                file_name="bilingual_neet_jee_PERFECT.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )

        except Exception:
            status.error("Error occurred")
            st.code(traceback.format_exc())

st.markdown("---")
st.caption("v4.0 PERFECT • Header on page 1 • Questions from page 2 • Math improved • Images extracted")
