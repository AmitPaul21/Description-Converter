import streamlit as st
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from io import BytesIO
import re
import pandas as pd

# ----------------- Load Replacement Dictionary from Excel -----------------
excel_file_path = r"replacement_dict.xlsx"

try:
    df = pd.read_excel(excel_file_path)
    df = df.dropna(subset=['Find'])
    replacement_dict = dict(zip(df['Find'], df['Replace With']))
    print("✅ Replacement dictionary loaded successfully!")
except Exception as e:
    print(f"❌ Error loading replacement dictionary: {e}")
    replacement_dict = {}

# ---------------- Safe Whole Word Replacement ----------------
def replace_words_safe(text, replacement_dict):
    for key, val in replacement_dict.items():
        pattern = r'\b{}\b'.format(re.escape(key))
        text = re.sub(pattern, val, text)
    return text

# ---------------- Clean Special Formatting ----------------
def clean_text(text):
    while '  ' in text:
        text = text.replace('  ', ' ')
    text = re.sub(r',\s*', ', ', text)
    text = re.sub(r'\(', ' (', text)
    while '  ' in text:
        text = text.replace('  ', ' ')
    text = re.sub(r'\b([A-Z]+)\s*\(', lambda m: m.group(1).capitalize() + ' (', text)
    return text

# ---------------- Remove Extra Empty Paragraphs ----------------
def remove_extra_empty_paragraphs(doc):
    paras = doc.paragraphs
    i = 0
    while i < len(paras) - 1:
        if paras[i].text.strip() == '' and paras[i+1].text.strip() == '':
            p = paras[i]._element
            p.getparent().remove(p)
            paras = doc.paragraphs
            i -= 1
        i += 1
    return doc

# ---------------- Styled Line Break Helper ----------------
def add_styled_break(para, break_type=WD_BREAK.LINE):
    r = para.add_run("")  # empty run
    r.font.name = "Times New Roman"
    r.font.size = Pt(10)

    rPr = r._element.get_or_add_rPr()
    rFonts = OxmlElement("w:rFonts")
    rFonts.set(qn("w:ascii"), "Times New Roman")
    rFonts.set(qn("w:hAnsi"), "Times New Roman")
    rFonts.set(qn("w:cs"), "Times New Roman")
    rFonts.set(qn("w:eastAsia"), "Times New Roman")
    rPr.append(rFonts)

    r.add_break(break_type)
    return r

# ---------------- Helper: write text with manual breaks ----------------
def set_para_text_with_manual_breaks(para, text):
    text = text.replace('\r\n', '\n').replace('\r', '\n')
    pattern = re.compile(r'([05])\s([SACDL])|\n')
    pos = 0

    for m in pattern.finditer(text):
        pre = text[pos:m.start()]
        if pre:
            r = para.add_run(pre)
            r.font.name = "Times New Roman"
            r.font.size = Pt(10)

        if m.group(0) == '\n':
            add_styled_break(para)
        else:
            num = m.group(1)
            let = m.group(2)

            r1 = para.add_run(num)
            r1.font.name = "Times New Roman"
            r1.font.size = Pt(10)
            add_styled_break(para)

            r2 = para.add_run(let)
            r2.font.name = "Times New Roman"
            r2.font.size = Pt(10)

        pos = m.end()

    tail = text[pos:]
    if tail:
        r = para.add_run(tail)
        r.font.name = "Times New Roman"
        r.font.size = Pt(10)

    if not para.runs:
        r = para.add_run('')
        r.font.name = "Times New Roman"
        r.font.size = Pt(10)

# ---------------- Document Formatting Function ----------------
def format_document(doc):
    try:
        normal = doc.styles['Normal']
        normal.font.name = 'Times New Roman'
        normal.font.size = Pt(10)
        normal.paragraph_format.line_spacing = 1.5
        normal.paragraph_format.space_before = Pt(0)
        normal.paragraph_format.space_after = Pt(0)
    except Exception:
        normal = None

    # ---------------- Body Paragraphs ----------------
    for para in doc.paragraphs:
        raw = para.text
        text = replace_words_safe(raw, replacement_dict)
        text = clean_text(text).strip()

        if text:
            text = re.sub(r'[^A-Za-z0-9]+$', '', text)
            text += "."

        for run in list(para.runs):
            run.text = ''
        try:
            para._p.clear_content()
        except Exception:
            pass

        set_para_text_with_manual_breaks(para, text)

        if normal is not None:
            para.style = normal
        para.paragraph_format.line_spacing = 1.5
        para.paragraph_format.space_before = Pt(0)
        para.paragraph_format.space_after = Pt(0)

    doc = remove_extra_empty_paragraphs(doc)

    # ---------------- Tables ----------------
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    raw = para.text
                    text = replace_words_safe(raw, replacement_dict)
                    text = clean_text(text).strip()

                    for run in list(para.runs):
                        run.text = ''
                    try:
                        para._p.clear_content()
                    except Exception:
                        pass

                    set_para_text_with_manual_breaks(para, text)

                    if normal is not None:
                        para.style = normal
                    para.paragraph_format.line_spacing = 1.5
                    para.paragraph_format.space_before = Pt(0)
                    para.paragraph_format.space_after = Pt(0)

    # ---------------- Headers and Footers ----------------
    for section in doc.sections:
        for hf in (section.header, section.footer):
            for para in hf.paragraphs:
                raw = para.text
                text = replace_words_safe(raw, replacement_dict)
                text = clean_text(text).strip()

                for run in list(para.runs):
                    run.text = ''
                try:
                    para._p.clear_content()
                except Exception:
                    pass

                set_para_text_with_manual_breaks(para, text)

                if normal is not None:
                    para.style = normal
                para.paragraph_format.line_spacing = 1.5
                para.paragraph_format.space_before = Pt(0)
                para.paragraph_format.space_after = Pt(0)

    # ---------------- Page Setup Margins ----------------
    for section in doc.sections:
        section.top_margin = Cm(2.29)
        section.bottom_margin = Cm(1.27)
        section.left_margin = Cm(2.54)
        section.right_margin = Cm(2.54)

    return doc

# ---------------- Custom Header Formatter ----------------
def format_header(doc, well_name):
    for section in doc.sections:
        header = section.header

        # Clear existing header content
        for para in header.paragraphs:
            p = para._element
            p.getparent().remove(p)

        # Well name (centered, all caps)
        para1 = header.add_paragraph()
        para1.alignment = 1  # Center
        para1.paragraph_format.line_spacing = 1.0
        run1 = para1.add_run(well_name.upper())
        run1.font.name = "Times New Roman"
        run1.font.size = Pt(10)

        # Empty line
        header.add_paragraph().paragraph_format.line_spacing = 1.0

        # SAMPLE DESCRIPTIONS (centered, underlined)
        para2 = header.add_paragraph()
        para2.alignment = 1  # Center
        para2.paragraph_format.line_spacing = 1.0
        run2 = para2.add_run("SAMPLE DESCRIPTIONS")
        run2.font.name = "Times New Roman"
        run2.font.size = Pt(10)
        run2.underline = True

        # Empty line
        header.add_paragraph().paragraph_format.line_spacing = 1.0

        # Depth (m), left aligned
        para3 = header.add_paragraph()
        para3.alignment = 0  # Left
        para3.paragraph_format.line_spacing = 1.0
        run3 = para3.add_run("Depth (m)")
        run3.font.name = "Times New Roman"
        run3.font.size = Pt(10)

        # Empty line
        header.add_paragraph().paragraph_format.line_spacing = 1.0

    # ---------------- Remove all footers ----------------
    for section in doc.sections:
        footer = section.footer
        for para in footer.paragraphs:
            p = para._element
            p.getparent().remove(p)

    return doc


# ---------------- Streamlit UI ----------------
st.title("Word File Formatter 📝")

uploaded_file = st.file_uploader("Upload a Word file (.docx)", type=["docx"])
well_name = st.text_input("Enter Well Name (will appear in header)", "")

if uploaded_file is not None and st.button("Format File"):
    try:
        doc = Document(uploaded_file)
        updated_doc = format_document(doc)

        if well_name.strip():
            updated_doc = format_header(updated_doc, well_name)

        buffer = BytesIO()
        updated_doc.save(buffer)
        buffer.seek(0)

        st.download_button(
            label="Download Updated File",
            data=buffer,
            file_name="formatted_output.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    except Exception as e:
        st.error(f"Error processing the document: {e}")

