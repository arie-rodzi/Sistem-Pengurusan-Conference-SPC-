import io, os, zipfile
from datetime import datetime

import streamlit as st
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt
from docxcompose.composer import Composer

st.set_page_config(page_title="SPC — Merge DOCX + TOC (TC fields)", layout="wide")
st.title("📚 SPC — Merge DOCX (PDF-like) + TOC Tajuk Kiri/Nombor Kanan")
st.caption("TOC di atas (guna TC fields). TOC tidak bernombor; penomboran bermula 1 pada dokumen pertama dan bersambung.")

# ================= Helpers =================

def zip_docx_entries_in_order(zip_bytes: bytes):
    """Return [(name_in_zip, bytes)] preserving ZIP (ZipInfo) order."""
    out = []
    with zipfile.ZipFile(io.BytesIO(zip_bytes)) as zf:
        for info in zf.infolist():
            if info.is_dir():
                continue
            name = info.filename
            if name.lower().endswith(".docx"):
                with zf.open(info, "r") as fp:
                    out.append((name, fp.read()))
    return out

def extract_title_from_doc_bytes(doc_bytes: bytes, fallback_name: str) -> str:
    """Use Heading 1 as title if present; otherwise fallback to filename (without .docx)."""
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
    """Insert a Word field (PAGE / NUMPAGES / PAGEREF ...) into an existing paragraph."""
    r1 = OxmlElement("w:r")
    fc1 = OxmlElement("w:fldChar"); fc1.set(qn("w:fldCharType"), "begin")
    r1.append(fc1); paragraph._p.append(r1)

    r2 = OxmlElement("w:r")
    it = OxmlElement("w:instrText"); it.set(qn("xml:space"), "preserve"); it.text = f" {field_code} "
    r2.append(it); paragraph._p.append(r2)

    r3 = OxmlElement("w:r")
    fc3 = OxmlElement("w:fldChar"); fc3.set(qn("w:fldCharType"), "end")
    r3.append(fc3); paragraph._p.append(r3)

def add_hidden_tc_paragraph(doc: Document, title: str, level: int = 1):
    """
    Insert a hidden TC field paragraph at the current position.
    We tag entries with identifier X so the TOC can find them:  TC "Title" \l 1 \f X
    """
    p = doc.add_paragraph()

    # BEGIN (hidden)
    r_begin = OxmlElement("w:r")
    rpr = OxmlElement("w:rPr"); vanish = OxmlElement("w:vanish"); rpr.append(vanish)
    fld_begin = OxmlElement("w:fldChar"); fld_begin.set(qn("w:fldCharType"), "begin")
    r_begin.append(rpr); r_begin.append(fld_begin); p._p.append(r_begin)

    # INSTR (hidden)  -> include \f X identifier
    r_instr = OxmlElement("w:r")
    rpr2 = OxmlElement("w:rPr"); rpr2.append(OxmlElement("w:vanish"))
    instr = OxmlElement("w:instrText"); instr.set(qn("xml:space"), "preserve")
    safe = title.replace('"', "'")
    instr.text = f' TC "{safe}" \\l {level} \\f X '
    r_instr.append(rpr2); r_instr.append(instr); p._p.append(r_instr)

    # END (hidden)
    r_end = OxmlElement("w:r")
    rpr3 = OxmlElement("w:rPr"); rpr3.append(OxmlElement("w:vanish"))
    fld_end = OxmlElement("w:fldChar"); fld_end.set(qn("w:fldCharType"), "end")
    r_end.append(rpr3); r_end.append(fld_end); p._p.append(r_end)

def add_toc_from_tc_at_top(doc: Document, entries_count: int):
    """
    Add a 'Table of Contents' header and a TOC field that reads TC entries with identifier X.
    """
    # Title
    title_para = doc.add_paragraph()
    run = title_para.add_run("Table of Contents"); run.bold = True; run.font.size = Pt(14)
    title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # TOC field reading TC entries with identifier X: { TOC \h \z \f "X" }
    p = doc.add_paragraph()
    fld = OxmlElement("w:fldSimple")
    fld.set(qn("w:instr"), 'TOC \\h \\z \\f "X"')
    p._p.append(fld)

def clear_pgnumtype_for_all_sections(doc: Document):
    """Remove any existing page numbering start/restart on all sections."""
    for section in doc.sections:
        sectPr = section._sectPr
        pgNumType = sectPr.find(qn('w:pgNumType'))
        if pgNumType is not None:
            sectPr.remove(pgNumType)

def start_numbering_at_section(doc: Document, index: int, start_at: int = 1):
    """Force a given section to start numbering at 'start_at'."""
    if index < len(doc.sections):
        sectPr = doc.sections[index]._sectPr
        pgNumType = OxmlElement('w:pgNumType')
        pgNumType.set(qn('w:start'), str(start_at))
        sectPr.append(pgNumType)

def add_page_numbers_from_section(doc: Document, start_index: int = 1):
    """
    Add 'Page X of Y' to the footer of the first content section,
    then link all following sections' footers to previous.
    TOC section (0) remains without a page number.
    """
    if len(doc.sections) == 0:
        return
    # TOC section has no page number
    doc.sections[0].different_first_page_header_footer = True

    if start_index < len(doc.sections):
        s = doc.sections[start_index]
        s.different_first_page_header_footer = False
        p = s.footer.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.add_run("Page "); add_field_run(p, "PAGE"); p.add_run(" of "); add_field_run(p, "NUMPAGES")

        # Link all subsequent sections' footers to previous so numbering continues
        for i in range(start_index + 1, len(doc.sections)):
            doc.sections[i].footer.is_linked_to_previous = True
            doc.sections[i].different_first_page_header_footer = False

def set_update_fields_on_open(doc: Document):
    """Ask Word to update all fields (PAGE/NUMPAGES/TOC) when opening the document."""
    settings = doc.settings.element
    upd = OxmlElement("w:updateFields"); upd.set(qn("w:val"), "true")
    settings.append(upd)

# ================= Core =================

def combine_zip_with_toc_tc(zip_bytes: bytes) -> bytes:
    """
    Build a proceedings doc:
    - Top page: TOC built from TC fields (identifier X)
    - Each paper starts on a new page (PDF-like)
    - TOC has no page number; numbering starts at 1 on first paper and continues
    - No formatting of sub-docs is changed
    """
    files = zip_docx_entries_in_order(zip_bytes)
    if not files:
        raise ValueError("ZIP tidak mengandungi .docx")

    meta = []
    for i, (name, blob) in enumerate(files, start=1):
        meta.append({"title": extract_title_from_doc_bytes(blob, name), "blob": blob})

    # 1) Base doc: TOC page
    base = Document()
    add_toc_from_tc_at_top(base, len(meta))
    base.add_page_break()  # separate TOC from content

    composer = Composer(base)

    # 2) For each paper: insert a hidden TC entry at the paper start, then append the sub-doc
    for idx, item in enumerate(meta):
        if idx > 0:
            base.add_page_break()  # each paper on a new page
        # TC entry (hidden) at the start of this paper
        add_hidden_tc_paragraph(base, item["title"], level=1)
        # Append sub-doc unchanged
        sub = Document(io.BytesIO(item["blob"]))
        composer.append(sub)

    # 3) Save, reopen, normalize numbering & add page numbers
    buf = io.BytesIO(); composer.save(buf); buf.seek(0)
    doc = Document(buf)

    clear_pgnumtype_for_all_sections(doc)
    if len(doc.sections) >= 2:
        start_numbering_at_section(doc, 1, 1)  # start numbering at first content section
    add_page_numbers_from_section(doc, start_index=1)
    set_update_fields_on_open(doc)

    out = io.BytesIO(); doc.save(out); out.seek(0)
    return out.read()

# ================= UI =================

st.subheader("Muat Naik ZIP Anda")
zip_file = st.file_uploader("Upload satu ZIP (folder + .docx) — susunan ikut folder (ZipInfo order).",
                            type=["zip"], accept_multiple_files=False)

st.info("TOC guna TC fields di awal setiap paper. TOC tidak bernombor; halaman dokumen bermula 1 pada paper pertama dan bersambung. "
        "Jika Word tidak auto-kemas kini, tekan Ctrl+A → F9 selepas dibuka.")

default_name = f"SPC_Proceedings_{datetime.now().strftime('%Y%m%d_%H%M')}.docx"
out_name = st.text_input("Nama fail output", value=default_name)

if st.button("🚀 Gabungkan (TOC TC-fields + setiap dokumen halaman baharu)"):
    try:
        if not zip_file:
            st.warning("Sila upload satu fail ZIP.")
        else:
            with st.spinner("Menggabungkan dokumen..."):
                compiled = combine_zip_with_toc_tc(zip_file.read())
            st.success("Siap! Muat turun di bawah.")
            st.download_button("⬇️ Muat Turun Fail Gabungan",
                               data=compiled,
                               file_name=out_name or "SPC_Proceedings.docx",
                               mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    except Exception as e:
        st.error("Ralat semasa menggabungkan dokumen.")
        st.exception(e)
