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

# =========================
# App Config
# =========================
st.set_page_config(page_title="SPC — Folder Order + TOC + Page Numbers", layout="wide")
st.title("📚 SPC — Susun Ikut Folder (ZIP) + TOC + Muka Surat")
st.caption("Gabung .docx ikut susunan folder dalam ZIP. TOC di awal, muka surat di footer. Tiada perubahan pada format kandungan setiap fail.")

st.markdown(
    """
**Cara guna ringkas**
1) Upload satu **ZIP** yang mengandungi folder + fail **.docx**.  
2) Klik **Gabungkan** → muat turun satu fail .docx gabungan.  
3) Buka di Microsoft Word → **Right-click** pada TOC → **Update Field** → **Update entire table** (untuk paparan gaya laporan dengan dot leader & nombor muka surat kanan).

> **Nota TOC**: TOC bergantung pada **Heading 1/2/3** dalam dokumen asal. Tanpa heading, entri mungkin kosong.
"""
)

# =========================
# Utilities
# =========================
def zip_docx_entries_in_order(zip_bytes: bytes):
    """
    Pulangkan senarai (name_in_zip, bytes) mengikut susunan entri dalam ZIP (ZipInfo order).
    Hanya .docx diambil, folder diabaikan.
    """
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
    """
    Sisip TOC di awal. Gaya visual 'laporan' (dot leader + nombor kanan) akan dipaparkan oleh Word
    selepas pengguna 'Update Field'. Di sini kita sisip field standard TOC.
    """
    # Tajuk TOC
    title_para = doc.add_paragraph()
    run = title_para.add_run(title)
    run.bold = True
    run.font.size = Pt(16)
    title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Medan TOC: { TOC \o "1-3" \h \z \u }
    # \o "1-3" = Heading 1-3; \h = hyperlinked; \z = hide page numbers in web layout; \u = use applied outline levels
    p = doc.add_paragraph()
    fldSimple = OxmlElement('w:fldSimple')
    fldSimple.set(qn('w:instr'), 'TOC \\o "1-3" \\h \\z \\u')
    p._p.append(fldSimple)


def add_field_run(paragraph, field):
    """
    Sisip Word field (cth. PAGE / NUMPAGES) ke dalam paragraph sedia ada.
    """
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
    """
    Tambah 'Page X of Y' di footer, penjajaran tengah.
    """
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


def make_sample_conference_zip(n_docs: int = 10) -> bytes:
    """
    Jana ZIP in-memory yang mengandungi n_docs fail .docx contoh
    (tajuk Heading1, penulis Heading2, dan isi ringkas) untuk ujian cepat.
    """
    mem_zip = io.BytesIO()
    with zipfile.ZipFile(mem_zip, 'w', zipfile.ZIP_DEFLATED) as zf:
        for i in range(1, n_docs + 1):
            doc = Document()
            # Heading untuk TOC
            doc.add_heading(f"Extended Abstract Title {i}", level=1)
            doc.add_heading(f"Author A{i}, Author B{i}", level=2)
            doc.add_paragraph(
                "Abstract:\n"
                "This study investigates aspects of conference paper management. "
                "It outlines methodology, results, and implications. "
                "Future work will scale this preliminary study."
            )
            for j in range(3):
                doc.add_paragraph(f"Section {j+1}: Detailed discussion for document {i}.")

            # Simpan ke bytes lalu masuk ZIP
            buf = io.BytesIO()
            doc.save(buf)
            buf.seek(0)
            # Letak dalam folder 'papers/' untuk demo struktur folder
            zf.writestr(f"papers/paper{i:02d}.docx", buf.read())
    mem_zip.seek(0)
    return mem_zip.read()


# =========================
# UI — Sample Docs
# =========================
with st.expander("🧪 Jana & Muat Turun 10 Dokumen Contoh (Extended Abstract)"):
    st.write(
        "Klik butang di bawah untuk memuat turun ZIP contoh (mengandungi 10 fail .docx dalam folder `papers/`). "
        "Fail contoh sudah ada Heading supaya TOC akan memaparkan entri."
    )
    if st.button("🔧 Jana ZIP Contoh (10 DOCX)"):
        sample_zip = make_sample_conference_zip(10)
        st.download_button(
            "⬇️ Muat Turun ZIP Contoh",
            data=sample_zip,
            file_name="sample_conference_docs.zip",
            mime="application/zip",
        )

# =========================
# UI — Main Uploader
# =========================
st.subheader("Muat Naik ZIP Anda")
zip_file = st.file_uploader(
    "Upload satu ZIP (mengandungi folder + .docx). Susunan akan ikut susunan dalam ZIP.",
    type=["zip"],
    accept_multiple_files=False
)

st.info("Disaran guna ZIP supaya susunan ikut folder adalah tepat. Jika perlu, guna ZIP contoh di atas untuk ujian.")

default_name = f"SPC_Proceedings_{datetime.now().strftime('%Y%m%d_%H%M')}.docx"
out_name = st.text_input("Nama fail output", value=default_name)

# =========================
# Action
# =========================
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
