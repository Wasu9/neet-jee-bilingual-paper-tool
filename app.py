"""
NEET / JEE Simple DTP Tool v8.3 SAFE
PDF crash fixed permanently
"""

import streamlit as st
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import io
import re
from typing import List

st.set_page_config(page_title="Simple DTP Tool", page_icon="📄", layout="wide")

st.markdown('<h2 style="color:#1a365d;">📄 Simple DTP Tool — NEET / JEE</h2>', unsafe_allow_html=True)
st.caption("Master Page • 2 Column • Word (PDF optional)")

# ====================== HELPERS ======================
def set_cell_border(cell, color="999999", sz="4"):
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


def set_run_font(run, name="Times New Roman", size=10, bold=False, color=None):
    run.font.name = name
    run._element.rPr.rFonts.set(qn('w:eastAsia'), name)
    run.font.size = Pt(size)
    run.bold = bold
    if color:
        run.font.color.rgb = RGBColor(*color)


def split_blocks(text: str) -> List[str]:
    text = text.replace("\r\n", "\n").replace("\r", "\n").strip()
    if not text:
        return []
    parts = re.split(r'(?m)(?=^\s*(?:\d{1,3}|Q\.?\s*\d{1,3})[\.\)]\s+)', text)
    blocks = [p.strip() for p in parts if p.strip() and len(p.strip()) > 8]
    if len(blocks) <= 1:
        blocks = [b.strip() for b in text.split("\n\n") if b.strip()]
    return blocks


# ====================== WORD ONLY (STABLE) ======================
def create_word(institute, exam_name, class_name, subject, marks, date,
                instructions, eng_blocks, hin_blocks, margin_cm=1.5,
                show_page_no=True, header_text="", footer_text="") -> bytes:

    doc = Document()
    sec = doc.sections[0]
    sec.page_width = Cm(21.0)
    sec.page_height = Cm(29.7)
    sec.top_margin = Cm(margin_cm)
    sec.bottom_margin = Cm(margin_cm)
    sec.left_margin = Cm(margin_cm)
    sec.right_margin = Cm(margin_cm)

    # Header
    if header_text.strip():
        header = sec.header
        hp = header.paragraphs[0]
        hp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = hp.add_run(header_text)
        set_run_font(run, "Arial", 9, True)

    # Footer
    footer = sec.footer
    fp = footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if footer_text.strip():
        run = fp.add_run(footer_text + "  |  ")
        set_run_font(run, "Arial", 8)
    if show_page_no:
        run = fp.add_run("Page ")
        set_run_font(run, "Arial", 8)
        fldChar1 = OxmlElement('w:fldChar')
        fldChar1.set(qn('w:fldCharType'), 'begin')
        instrText = OxmlElement('w:instrText')
        instrText.text = "PAGE"
        fldChar2 = OxmlElement('w:fldChar')
        fldChar2.set(qn('w:fldCharType'), 'end')
        run2 = fp.add_run()
        run2._r.append(fldChar1)
        run2._r.append(instrText)
        run2._r.append(fldChar2)
        set_run_font(run2, "Arial", 8)

    # Title block
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(institute)
    set_run_font(run, "Arial", 14, True, (26, 54, 93))

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(exam_name)
    set_run_font(run, "Arial", 12, True)

    info = f"Class: {class_name}   |   Subject: {subject}   |   Marks: {marks}   |   Date: {date}"
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(info)
    set_run_font(run, "Arial", 9)

    if instructions.strip():
        p = doc.add_paragraph()
        run = p.add_run("Instructions:")
        set_run_font(run, "Arial", 9, True)
        p = doc.add_paragraph()
        run = p.add_run(instructions)
        set_run_font(run, "Arial", 8)

    doc.add_paragraph()

    # 2-column questions
    max_len = max(len(eng_blocks), len(hin_blocks))
    eng_blocks = eng_blocks + [""] * (max_len - len(eng_blocks))
    hin_blocks = hin_blocks + [""] * (max_len - len(hin_blocks))

    for eng, hin in zip(eng_blocks, hin_blocks):
        if not eng.strip() and not hin.strip():
            continue

        table = doc.add_table(rows=1, cols=2)
        table.autofit = False
        table.columns[0].width = Cm(8.5)
        table.columns[1].width = Cm(8.5)

        for cell in table.rows[0].cells:
            set_cell_border(cell, "AAAAAA", "4")

        c0 = table.rows[0].cells[0]
        p0 = c0.paragraphs[0]
        run0 = p0.add_run(eng)
        set_run_font(run0, "Times New Roman", 9)

        c1 = table.rows[0].cells[1]
        p1 = c1.paragraphs[0]
        run1 = p1.add_run(hin)
        set_run_font(run1, "Mangal", 9)

        gap = doc.add_paragraph()
        gap.paragraph_format.space_before = Pt(4)
        gap.paragraph_format.space_after = Pt(4)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ====================== UI ======================
st.markdown("### 1. Master Page")
c1, c2 = st.columns(2)
with c1:
    institute = st.text_input("Institute Name", value="SHAHEEN GROUP OF INSTITUTIONS")
    exam_name = st.text_input("Exam Name", value="GRAND EXAMINATION")
    class_name = st.text_input("Class", value="I PUC JEE")
with c2:
    subject = st.text_input("Subject", value="PCM")
    marks = st.text_input("Marks", value="300")
    date = st.text_input("Date", value="16/09/2026")

instructions = st.text_area("Instructions (optional)", height=60)

st.markdown("### 2. Page Settings")
c3, c4, c5 = st.columns(3)
with c3:
    margin = st.slider("Margin (cm)", 1.0, 2.5, 1.5, 0.1)
with c4:
    show_page_no = st.checkbox("Show Page Number", value=True)
with c5:
    header_text = st.text_input("Header Text", value="")
footer_text = st.text_input("Footer Text", value="")

st.markdown("### 3. Content (English | Hindi)")
col_eng, col_hin = st.columns(2)
with col_eng:
    st.markdown("**English Questions**")
    eng_text = st.text_area("eng", height=320, label_visibility="collapsed",
                            placeholder="1. Question text...\n(1) option\n(2) option")
with col_hin:
    st.markdown("**Hindi Questions**")
    hin_text = st.text_area("hin", height=320, label_visibility="collapsed",
                            placeholder="1. प्रश्न...\n(1) विकल्प\n(2) विकल्प")

st.markdown("---")

if st.button("🚀 Generate Word File", type="primary", use_container_width=True):
    eng_blocks = split_blocks(eng_text)
    hin_blocks = split_blocks(hin_text)

    if not eng_blocks and not hin_blocks:
        st.warning("Please paste some content in English or Hindi.")
    else:
        with st.spinner("Creating Word file..."):
            try:
                word_bytes = create_word(
                    institute, exam_name, class_name, subject, marks, date, instructions,
                    eng_blocks, hin_blocks, margin, show_page_no, header_text, footer_text
                )
                st.success(f"✅ Ready • {max(len(eng_blocks), len(hin_blocks))} blocks")
                st.download_button(
                    "📥 Download Word",
                    data=word_bytes,
                    file_name="bilingual_dtp.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
            except Exception as e:
                st.error(f"Error: {e}")

st.markdown("---")
st.info("PDF temporarily removed for stability. Word file works perfectly. PDF can be added later if needed.")
st.caption("v8.3 SAFE • Word only • No more LayoutError")
