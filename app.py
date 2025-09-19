import io, os, zipfile, re
from datetime import datetime

import streamlit as st
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_SECTION
from docx.text.run import WD_BREAK          # <-- penting: PAGE break sebenar
from docx.shared import Pt
from docxcompose.composer import Composer

st.set_page_config(page_title="SPC — Merge DOCX + TOC (TC fields)", layout="wide")
st.title("📚 SPC — Merge DOCX (PDF-like) + TOC Tajuk & Nombor Halaman")
st.caption("Upload ZIP atau banyak DOCX terus. TOC di atas (tanpa page number), dokumen pertama mula page 1 & bersambung. Setiap dokumen di halaman baharu.")

# ----------------- Helpers -----------------

def natural_key(s: str):
    return [int(t) if t.isdigit() else t.lower() for t in re.split(r'(\d+)', s)]

def files_from_zip(zip_bytes: bytes):
    out = []
    with zipfile.ZipFile(io.BytesIO(zip_bytes)) as zf:
        for info in zf.infolist():
            if info.is_dir():
                continue
            name = info.filename
            if name.lower().endswith(".docx"):
                with zf.open(info, "r") as fp:
                    out.append((name, fp.read()))
    return out  # ZipInfo order (ikut folder)

def files_from_uploads(uploaded_files):
    out = []
    for uf in uploaded_files:
        out.append((uf.name, uf.read()))
    out.sort(key=lambda x: natural_key(x[0]))  # natural sort nama fail
    return out

def extract_title_from_doc_bytes(doc_bytes: bytes, fallback_name: str) -> str:
    try:
        d = Document(io.BytesIO(doc_bytes))
        for p in d.paragraphs:
            if getattr(p.style, "name", "").lower().startswith("heading 1"):
                t = (p.text or "").strip()
                if t:
                    return t
    except Exception:
        pass
    return os.path.splitext(os.path.basename(fallback_name))[0]

def add_field_run(paragraph, field_code: str):
    r1 = OxmlElement("w:r"); fc1 = OxmlElement("w:fldChar"); fc1.set(qn("w:fldCharType"), "begin")
    r1.append(fc1); paragraph._p.append(r1)
    r2 = OxmlElement("w:r"); it = OxmlElement("w:instrText"); it.set(qn("xml:space"), "preserve"); it.text = f" {field_code} "
    r2.append(it); paragraph._p.append(r2)
    r3 = OxmlElement("w:r"); fc3 = OxmlElement("w:fldChar"); fc3.set(qn("w:fldCharType"), "end")
    r3.append(fc3); paragraph._p.append(r3)

def add_hidden_tc_paragraph(doc: Document, title: str, level: int = 1):
    """
    Sisip TC tersembunyi (identifier X) di lokasi semasa:
    TC "Title" \l 1 \f X
    """
    p = doc.add_paragraph()

    r_begin = OxmlElement("w:r")
    rpr = OxmlElement("w:rPr"); rpr.append(OxmlElement("w:vanish"))
    fld_begin = OxmlElement("w:fldChar"); fld_begin.set(qn("w:fldCharType"), "begin")
    r_begin.append(rpr); r_begin.append(fld_begin); p._p.append(r_begin)

    r_instr = OxmlElement("w:r")
    rpr2 = OxmlElement("w:rPr"); rpr2.append(OxmlElement("w:vanish"))
    instr = OxmlElement("w:instrText"); instr.set(qn("xml:space"), "preserve")
    safe = title.replace('"', "'")
    instr.text = f' TC "{safe}" \\l {level} \\f X '
    r_instr.append(rpr2); r_instr.append(instr); p._p.append(r_instr)

    r_end = OxmlElement("w:r")
    rpr3 = OxmlElement("w:rPr"); rpr3.append(OxmlElement("w:vanish"))
    fld_end = OxmlElement("w:fldChar"); fld_end.set(qn("w:fldCharType"), "end")
    r_end.append(fld_end); r_end.append(rpr3); p._p.append(r_end)

def add_toc_from_tc_at_top(doc: Document):
    """
    Tajuk + medan TOC yang baca TC identifier X:
    { TOC \h \z \f X }  <-- tanpa petikan
    """
    title_para = doc.add_paragraph()
    run = title_para.add_run("Table of Contents"); run.bold = True; run.font.size = Pt(14)
    title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER

    p = doc.add_paragraph()
    fld = OxmlElement("w:fldSimple")
    fld.set(qn("w:instr"), 'TOC \\h \\z \\f X')
    p._p.append(fld)

def clear_pgnumtype_for_all_sections(doc: Document):
    for section in doc.sections:
        sectPr = section._sectPr
        pgNumType = sectPr.find(qn('w:pgNumType'))
        if pgNumType is not None:
            sectPr.remove(pgNumType)

def start_numbering_at_section(doc: Document, index: int, start_at: int = 1):
    if index < len(doc.sections):
        sectPr = doc.sections[index]._sectPr
        pgNumType = OxmlElement('w:pgNumType')
        pgNumType.set(qn('w:start'), str(start_at))
        sectPr.append(pgNumType)

def add_page_numbers_from_section(doc: Document, start_index: int = 1):
    """
    Letak 'Page X of Y' pada footer seksyen pertama kandungan,
    seksyen berikutnya link ke previous. Seksyen 0 (TOC) tiada nombor.
    """
    if len(doc.sections) == 0:
        return
    # TOC section no number
    doc.sections[0].different_first_page_header_footer = True

    if start_index < len(doc.sections):
        s = doc.sections[start_index]
        s.different_first_page_header_footer = False
        p = s.footer.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.add_run("Page "); add_field_run(p, "PAGE"); p.add_run(" of "); add_field_run(p, "NUMPAGES")
        # propagate to others
        for i in range(start_index + 1, len(doc.sections)):
            doc.sections[i].footer.is_linked_to_previous = True
            doc.sections[i].different_first_page_header_footer = False

def set_update_fields_on_open(doc: Document):
    settings = doc.settings.element
    upd = OxmlElement("w:updateFields"); upd.set(qn("w:val"), "true")
    settings.append(upd)

# ----------------- Core -----------------

def combine_with_tc(files: list) -> bytes:
    """
    files: list[(display_name, bytes)] — ikut susunan diberi
    Hasil:
      - Halaman 1: TOC (tanpa page number)
      - SECTION BREAK selepas TOC
      - **Setiap dokumen bermula HALAMAN BAHARU (PAGE BREAK sebenar)**
      - TC tersembunyi di awal setiap dokumen (tajuk untuk TOC)
      - Nombor muka surat bermula 1 pada dokumen pertama & sambung
    """
    if not files:
        raise ValueError("Tiada .docx diberikan")

    meta = [{"title": extract_title_from_doc_bytes(b, n), "blob": b} for n, b in files]

    # 1) Base: TOC + SECTION BREAK (Next Page)
    base = Document()
    add_toc_from_tc_at_top(base)
    base.add_section(WD_SECTION.NEW_PAGE)   # seksyen kandungan asing dari TOC

    composer = Composer(base)

    # 2) Untuk setiap paper: PAGE BREAK sebenar (kecuali paper pertama), TC tersembunyi, append subdoc
    for idx, item in enumerate(meta):
        if idx > 0:
            br_para = base.add_paragraph()
            br_para.add_run().add_break(WD_BREAK.PAGE)   # <-- PAGE break (bukan line break)
        add_hidden_tc_paragraph(base, item["title"], level=1)
        sub = Document(io.BytesIO(item["blob"]))
        composer.append(sub)

    # 3) Simpan, buka semula, normalkan numbering & footer
    buf = io.BytesIO(); composer.save(buf); buf.seek(0)
    doc = Document(buf)

    clear_pgnumtype_for_all_sections(doc)
    if len(doc.sections) >= 2:
        start_numbering_at_section(doc, 1, 1)  # mula 1 pada seksyen kandungan pertama

    add_page_numbers_from_section(doc, start_index=1)
    set_update_fields_on_open(doc)

    out = io.BytesIO(); doc.save(out); out.seek(0)
    return out.read()

# ----------------- UI -----------------

st.subheader("Muat Naik Fail")
col1, col2 = st.columns(2)
with col1:
    many_files = st.file_uploader(
        "Upload banyak DOCX (boleh pilih semua fail dalam folder)",
        type=["docx"],
        accept_multiple_files=True,
        key="multi_docx",
    )
with col2:
    one_zip = st.file_uploader(
        "ATAU upload satu ZIP (mengandungi .docx ikut susunan folder)",
        type=["zip"],
        accept_multiple_files=False,
        key="zip_docx",
    )

st.caption("Tanpa ZIP, app ikut natural sort nama fail. Untuk ikut struktur folder tepat, guna ZIP.")

default_name = f"SPC_Proceedings_{datetime.now().strftime('%Y%m%d_%H%M')}.docx"
out_name = st.text_input("Nama fail output", value=default_name)

if st.button("🚀 Gabungkan"):
    try:
        files = []
        if many_files:
            files = files_from_uploads(many_files)
        if one_zip is not None:
            files += files_from_zip(one_zip.read())

        if not files:
            st.warning("Sila upload sekurang-kurangnya satu .docx atau satu .zip.")
        else:
            with st.spinner("Menggabungkan dokumen..."):
                combined = combine_with_tc(files)
            st.success("Siap! Muat turun di bawah. Jika Word tidak auto-update: Ctrl+A → F9.")
            st.download_button(
                "⬇️ Muat Turun Fail Gabungan",
                data=combined,
                file_name=out_name or "SPC_Proceedings.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )
    except Exception as e:
        st.error("Ralat semasa menggabungkan dokumen.")
        st.exception(e)
