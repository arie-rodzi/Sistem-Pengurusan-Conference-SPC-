import io
import os
import zipfile
from datetime import datetime

import streamlit as st
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docxcompose.composer import Composer

st.set_page_config(page_title="SPC — Folder Order + TOC + Page Numbers", layout="wide")
st.title("📚 SPC — Susun Ikut Folder (ZIP) + TOC + Muka Surat")
st.caption("Gabung .docx ikut susunan folder dalam ZIP. Tambah TOC di awal, tambah muka surat. Format asal dikekalkan (tiada heading/page break tambahan).")

def zip_docx_entries_in_order(zip_bytes: bytes):
    """Pulangkan senarai (name_in_zip, bytes) mengikut susunan entri dalam ZIP (ZipInfo order)."""
    result = []
    with zipfile.ZipFile(io.BytesIO(zip_bytes)) as zf:
        for info in zf.infolist():
            if info.is_dir():
                continue
            name = info.filename
            if name.lower().endswith(".docx"):
                with zf.open(info, 'r') as fp:
                    result.append((name, fp.read()))
    return result  # kekalkan susunan asal dalam ZIP

def add_toc_field(doc: Document, title="Table of Contents"):
    # Tajuk TOC (tengah)
    title_para = doc.add_paragraph()
    run = title_para.add_run(title)
    run.bold = True
    run.font.size = Pt(16)
    title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Medan TOC: { TOC \o "1-3" \h \z \u }
    p = doc.add_paragraph()
    fldSimple = OxmlElement('w:fldSimple')
    fldSimple.set(qn('w:instr'), 'TOC \\o "1-3" \\h \\z \\u')
    p._p.append(fldSimple)

def add_field_run(paragraph, field):
    """Sisip kod medan Word (PAGE / NUMPAGES) dalam perenggan sedia ada."""
    r = OxmlElement("w:r")
    fldChar = OxmlElement("w:fldChar")
    fldChar.set(qn("w:fldCharType"), "begin")
    r.append(fldChar)
    paragraph._p.append(r)

    r = OxmlElement("w:r")
    instrText = OxmlElement("w:instrText")
    instrText.set(qn("xml:space"), "preserve")
    instrText.text = f" {field} "
    r.append(instrText)
    paragraph._p.append(r)

    r = OxmlElement("w:r")
    fldChar = OxmlElement("w:fldChar")
    fldChar.set(qn("w:fldCharType"), "end")
    r.append(fldChar)
    paragraph._p.append(r)

def add_page_numbers(doc: Document):
    """Tambah 'Page X of Y' di footer, penjajaran tengah."""
    section = doc.sections[0]
    footer = section.footer
    paragraph = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph.add_run("Page ")
    add_field_run(paragraph, "PAGE")
    paragraph.add_run(" of ")
    add_field_run(paragraph, "NUMPAGES")

def combine_in_zip_order(zip_bytes: bytes) -> bytes:
    """
    Gabungkan dokumen mengikut susunan dalam ZIP.
    - Letak TOC di awal
    - Tambah muka surat (Page X of Y)
    - Append setiap .docx terus (tanpa heading/page break tambahan)
    """
    base = Document()
    add_toc_field(base, title="Table of Contents")
    add_page_numbers(base)

    composer = Composer(base)

    entries = zip_docx_entries_in_order(zip_bytes)
    if not entries:
        raise ValueError("ZIP tidak mengandungi sebarang .docx")

    for name, blob in entries:
        sub = Document(io.BytesIO(blob))
        composer.append(sub)  # append as-is; tiada perubahan pada kandungan

    bio = io.BytesIO()
    composer.save(bio)
    bio.seek(0)
    return bio.read()

st.subheader("Muat Naik")
zip_file = st.file_uploader(
    "Upload satu ZIP (mengandungi folder + .docx). Susunan akan ikut susunan dalam ZIP.",
    type=["zip"],
    accept_multiple_files=False
)

st.info("Jika anda muat naik fail individu .docx, susunan mungkin tidak tepat mengikut folder. Disaran gunakan ZIP.")

default_name = f"SPC_Proceedings_{datetime.now().strftime('%Y%m%d_%H%M')}.docx"
out_name = st.text_input("Nama fail output", value=default_name)

if st.button("🚀 Gabungkan Ikut Susunan Folder (ZIP)"):
    try:
        if not zip_file:
            st.warning("Sila upload satu fail ZIP.")
        else:
            with st.spinner("Menggabungkan dokumen ikut susunan folder dalam ZIP..."):
                combined_bytes = combine_in_zip_order(zip_file.read())
            st.success("Siap! Muat turun di bawah. (Di Microsoft Word, right-click TOC → Update Field → Update entire table)")
            st.download_button(
                "⬇️ Muat Turun Fail Gabungan",
                data=combined_bytes,
                file_name=out_name or "SPC_Proceedings.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )
    except Exception as e:
        st.error("Ralat semasa menggabungkan dokumen.")
        st.exception(e)
