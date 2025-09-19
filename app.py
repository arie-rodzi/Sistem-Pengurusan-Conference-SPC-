import io
import os
import zipfile
import tempfile
from datetime import datetime
from typing import List, Tuple

import streamlit as st
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docxcompose.composer import Composer

st.set_page_config(page_title="SPC — Sistem Pengurusan Conference", layout="wide")
st.title("📚 SPC — Sistem Pengurusan Conference")
st.caption("Gabung banyak DOCX → satu fail dengan TOC & nombor muka surat")

st.markdown(
    """
    **Fungsi:**
    1) Upload banyak **.docx** / 1 **.zip** berisi .docx  
    2) Gabung jadi **satu .docx**  
    3) Automatik tambah **Table of Contents (TOC)** & **nombor muka surat**  
    4) Sisip **page break** antara setiap dokumen  

    **Penting (Word/Office):** Selepas buka fail gabungan, _Right‑click_ TOC → **Update Field** → **Update entire table**.
    """
)

# ---------- Helpers ----------
def load_docx_from_mem(file_bytes: bytes) -> Document:
    return Document(io.BytesIO(file_bytes))

def add_heading(doc: Document, text: str, level: int = 1):
    p = doc.add_heading(text, level=level)
    return p

def add_page_break(doc: Document):
    doc.add_page_break()

def add_toc_field(doc: Document, title="Table of Contents"):
    # Tajuk TOC
    title_para = doc.add_paragraph()
    run = title_para.add_run(title)
    run.bold = True
    run.font.size = Pt(16)
    title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Field TOC: { TOC \o "1-3" \h \z \u }
    p = doc.add_paragraph()
    fldSimple = OxmlElement('w:fldSimple')
    fldSimple.set(qn('w:instr'), 'TOC \\o "1-3" \\h \\z \\u')
    p._p.append(fldSimple)

    # Page break selepas TOC
    add_page_break(doc)

def add_field_run(paragraph, field):
    """Sisip kod medan Word (cth. PAGE, NUMPAGES)."""
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
    """Tambah 'Page X of Y' di footer, tengah."""
    section = doc.sections[0]
    footer = section.footer
    paragraph = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph.add_run("Page ")
    add_field_run(paragraph, "PAGE")
    paragraph.add_run(" of ")
    add_field_run(paragraph, "NUMPAGES")

def sort_key_for_names(name: str) -> Tuple[int, str]:
    """Susun: nombor di depan nama fail (jika ada) diutamakan, selain itu abjad."""
    base = os.path.splitext(os.path.basename(name))[0]
    num = ''
    i = 0
    while i < len(base) and base[i].isdigit():
        num += base[i]
        i += 1
    if num:
        return (0, f"{int(num):08d}")
    return (1, base.lower())

def extract_zip_to_temp(zip_bytes: bytes):
    tmpdir = tempfile.mkdtemp()
    with zipfile.ZipFile(io.BytesIO(zip_bytes)) as zf:
        zf.extractall(tmpdir)
    docx_paths = []
    for root, _, files in os.walk(tmpdir):
        for fn in files:
            if fn.lower().endswith(".docx"):
                docx_paths.append(os.path.join(root, fn))
    return sorted(docx_paths, key=sort_key_for_names)

def combine_docx(files):
    """
    files: senarai (display_name, bytes)
    pulangkan bytes .docx gabungan
    """
    if not files:
        raise ValueError("Tiada fail .docx diberikan.")

    # Dokumen asas
    base = Document()
    add_toc_field(base, title="Table of Contents")
    add_page_numbers(base)
    composer = Composer(base)

    for idx, (display_name, blob) in enumerate(sorted(files, key=lambda x: sort_key_for_names(x[0]))):
        # Jadikan nama fail sebagai Heading 1 (untuk TOC)
        add_heading(base, os.path.splitext(display_name)[0], level=1)
        sub_doc = load_docx_from_mem(blob)
        composer.append(sub_doc)
        if idx < len(files) - 1:
            add_page_break(base)

    output = io.BytesIO()
    composer.save(output)
    output.seek(0)
    return output.read()

# ---------- UI ----------
st.subheader("Muat Naik Fail")
col1, col2 = st.columns(2)
with col1:
    many_files = st.file_uploader(
        "Upload banyak DOCX (boleh pilih lebih dari satu)",
        type=["docx"],
        accept_multiple_files=True,
        key="multi_docx",
    )
with col2:
    one_zip = st.file_uploader(
        "ATAU upload satu ZIP (mengandungi .docx)",
        type=["zip"],
        accept_multiple_files=False,
        key="zip_docx",
    )

st.caption("Saranan: Jika ada 50+ fail, guna ZIP untuk lebih stabil.")

default_name = f"SPC_Proceedings_{datetime.now().strftime('%Y%m%d_%H%M')}.docx"
out_name = st.text_input("Nama fail output", value=default_name)

if st.button("🚀 Gabungkan"):
    try:
        files_to_merge = []

        if many_files:
            for f in many_files:
                files_to_merge.append((f.name, f.read()))

        if one_zip is not None:
            paths = extract_zip_to_temp(one_zip.read())
            for p in paths:
                with open(p, "rb") as fp:
                    files_to_merge.append((os.path.basename(p), fp.read()))

        if not files_to_merge:
            st.warning("Sila upload sekurang-kurangnya satu .docx atau satu .zip.")
        else:
            with st.spinner("Menggabungkan dokumen..."):
                combined_bytes = combine_docx(files_to_merge)

            st.success("Siap! Muat turun di bawah. (Jangan lupa Update TOC dalam Word)")
            st.download_button(
                "⬇️ Muat Turun Fail Gabungan",
                data=combined_bytes,
                file_name=out_name or "SPC_Proceedings.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )
    except Exception as e:
        st.error("Ralat semasa menggabungkan dokumen.")
        st.exception(e)

st.markdown("---")
st.markdown(
    """
    **Amalan Terbaik**
    - Selepas buka fail gabungan di Word → _Right-click_ TOC → **Update Field** → **Update entire table**  
    - Gaya tajuk (Heading) asal dalam dokumen akan kekal dan menyumbang kepada TOC.  
    - Susunan fail: nombor di depan nama fail diutamakan, jika tiada ikut urutan abjad.  
    """
)
