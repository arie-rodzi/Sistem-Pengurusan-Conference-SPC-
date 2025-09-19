import io
import os
import zipfile
from datetime import datetime

import streamlit as st
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.shared import Inches
from docxcompose.composer import Composer

st.set_page_config(page_title="SPC — PDF-like Merge + Manual TOC + Page Numbers", layout="wide")
st.title("📚 SPC — Gabung DOCX (PDF-like) + TOC Tajuk Kiri / Nombor Kanan + Muka Surat")
st.caption("TOC manual (tajuk kiri, nombor kanan dengan dot leaders). Setiap dokumen bermula di halaman baharu. Kandungan asal tidak diubah.")

# ----------------- Helpers -----------------

def zip_docx_entries_in_order(zip_bytes: bytes):
    """Pulangkan senarai (name_in_zip, bytes) ikut susunan ZIP (ZipInfo order)."""
    result = []
    with zipfile.ZipFile(io.BytesIO(zip_bytes)) as zf:
        for info in zf.infolist():
            if info.is_dir():
                continue
            name = info.filename
            if name.lower().endswith(".docx"):
                with zf.open(info, "r") as fp:
                    result.append((name, fp.read()))
    return result

def extract_title_from_doc_bytes(doc_bytes: bytes, fallback_name: str) -> str:
    """Ambil tajuk dari Heading 1 jika ada; jika tiada guna nama fail (tanpa .docx)."""
    try:
        d = Document(io.BytesIO(doc_bytes))
        for p in d.paragraphs:
            try:
                if p.style and p.style.name and str(p.style.name).lower().startswith("heading 1"):
                    t = (p.text or "").strip()
                    if t:
                        return t
            except Exception:
                pass
    except Exception:
        pass
    return os.path.splitext(os.path.basename(fallback_name))[0]

def add_field_run(paragraph, field_code: str):
    """Sisip Word field (PAGE / NUMPAGES / PAGEREF ...) ke dalam paragraph sedia ada."""
    r_begin = OxmlElement("w:r")
    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")
    r_begin.append(fld_begin)
    paragraph._p.append(r_begin)

    r_instr = OxmlElement("w:r")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = f" {field_code} "
    r_instr.append(instr)
    paragraph._p.append(r_instr)

    r_end = OxmlElement("w:r")
    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")
    r_end.append(fld_end)
    paragraph._p.append(r_end)

def add_page_numbers_all_sections(doc: Document):
    """Tambah 'Page X of Y' di footer tengah untuk semua seksyen."""
    for section in doc.sections:
        footer = section.footer
        p = footer.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.add_run("Page ")
        add_field_run(p, "PAGE")
        p.add_run(" of ")
        add_field_run(p, "NUMPAGES")

def add_bookmark(paragraph, name: str):
    """Buat bookmark tersembunyi pada paragraph ini."""
    # start
    start = OxmlElement("w:bookmarkStart")
    start.set(qn("w:id"), "0")
    start.set(qn("w:name"), name)
    paragraph._p.insert(0, start)
    # end
    end = OxmlElement("w:bookmarkEnd")
    end.set(qn("w:id"), "0")
    paragraph._p.append(end)

def add_manual_toc(doc: Document, entries):
    """
    Bina TOC manual bergaya laporan:
    - Kiri: tajuk
    - TAB
    - Kanan: nombor halaman (PAGEREF bookmark)
    - Dot leaders
    entries: list of dicts {title, bookmark}
    """
    # Tajuk "Table of Contents"
    title_para = doc.paragraphs[0] if doc.paragraphs else doc.add_paragraph()
    run = title_para.add_run("Table of Contents")
    run.bold = True
    title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Satu baris kosong
    doc.add_paragraph("")

    # Set tab stop kanan dgn dot leaders pada setiap entri
    for e in entries:
        p = doc.add_paragraph()
        pf = p.paragraph_format
        # Right tab stop ~ 6.5 in (ikut margin standard A4). Laras jika perlu.
        pf.tab_stops.add_tab_stop(Inches(6.5), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)

        # Tajuk
        p.add_run(e["title"])
        # Tab
        p.add_run("\t")
        # Nombor halaman → PAGEREF bookmark
        add_field_run(p, f'PAGEREF {e["bookmark"]} \\h')

def combine_pdf_like_with_manual_toc(zip_bytes: bytes) -> bytes:
    """
    Gabung dokumen DOCX:
    - TOC manual di atas (tajuk kiri / nombor kanan, dot leaders)
    - Setiap dokumen bermula HALAMAN BAHARU
    - Letak bookmark DOC_i di permulaan setiap dokumen
    - Tambah muka surat utk semua seksyen
    - Tidak ubah kandungan asal sub-docs
    """
    entries = zip_docx_entries_in_order(zip_bytes)
    if not entries:
        raise ValueError("ZIP tidak mengandungi sebarang .docx")

    # 1) Sediakan list tajuk & nama bookmark
    toc_entries = []
    for idx, (name, blob) in enumerate(entries, start=1):
        title = extract_title_from_doc_bytes(blob, name)
        toc_entries.append({"title": title, "bookmark": f"DOC_{idx}", "blob": blob})

    # 2) Mula dokumen asas dengan TOC placeholder (akan diisi sejurus selepas)
    base = Document()
    base.add_paragraph()  # tempat tajuk TOC nanti

    # 3) Halaman baharu selepas TOC
    base.add_page_break()

    # 4) Gabungkan dokumen; sebelum setiap sub-doc, letak paragraph bookmark + page break
    composer = Composer(base)
    for i, entry in enumerate(toc_entries, start=1):
        if i > 1:
            base.add_page_break()
        # paragraph untuk bookmark permulaan dokumen i
        bm_para = base.add_paragraph()
        add_bookmark(bm_para, entry["bookmark"])
        # append sub document
        sub = Document(io.BytesIO(entry["blob"]))
        composer.append(sub)

    # 5) Simpan sementara gabungan
    tmp = io.BytesIO()
    composer.save(tmp)
    tmp.seek(0)

    # 6) Buka semula & bina TOC manual di atas berdasarkan bookmark yang sudah wujud
    compiled = Document(tmp)
    # Sisip TOC manual pada mula-mula (selepas tajuk placeholder)
    add_manual_toc(compiled, [{"title": e["title"], "bookmark": e["bookmark"]} for e in toc_entries])

    # 7) Tambah muka surat untuk semua seksyen
    add_page_numbers_all_sections(compiled)

    out = io.BytesIO()
    compiled.save(out)
    out.seek(0)
    return out.read()

# ----------------- UI -----------------

st.subheader("Muat Naik ZIP Anda")
zip_file = st.file_uploader(
    "Upload satu ZIP (mengandungi folder + .docx). Susunan ikut folder (ZipInfo order).",
    type=["zip"],
    accept_multiple_files=False
)

st.info(
    "TOC manual: tajuk di kiri, nombor halaman di kanan dengan dot leaders. "
    "Setiap dokumen bermula di halaman baharu. "
    "Jika nombor halaman belum muncul betul, pilih semua (Ctrl+A) → F9 / Right-click → Update Field di Word."
)

default_name = f"SPC_Proceedings_{datetime.now().strftime('%Y%m%d_%H%M')}.docx"
out_name = st.text_input("Nama fail output", value=default_name)

if st.button("🚀 Gabungkan (TOC manual + setiap dokumen halaman baharu)"):
    try:
        if not zip_file:
            st.warning("Sila upload satu fail ZIP.")
        else:
            with st.spinner("Menggabungkan dokumen ikut susunan folder dalam ZIP..."):
                combined_bytes = combine_pdf_like_with_manual_toc(zip_file.read())
            st.success("Siap! Muat turun di bawah.")
            st.download_button(
                "⬇️ Muat Turun Fail Gabungan",
                data=combined_bytes,
                file_name=out_name or "SPC_Proceedings.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )
    except Exception as e:
        st.error("Ralat semasa menggabungkan dokumen.")
        st.exception(e)
