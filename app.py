import io, os, zipfile
from datetime import datetime

import streamlit as st
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.shared import Inches, Pt
from docxcompose.composer import Composer

st.set_page_config(page_title="SPC — Merge DOCX + Manual TOC (TC fields)", layout="wide")
st.title("📚 SPC — Merge DOCX (PDF-like) + TOC Tajuk Kiri/Nombor Kanan")
st.caption("TOC di atas (guna TC fields), dokumen pertama mula page 1 & sambung seterusnya, TOC tiada page number.")

# ---------- Helpers ----------
EMU_PER_INCH = 914400.0

def zip_docx_entries_in_order(zip_bytes: bytes):
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
    try:
        d = Document(io.BytesIO(doc_bytes))
        for p in d.paragraphs:
            if getattr(p.style, "name", "").lower().startswith("heading 1"):
                t = (p.text or "").strip()
                if t: return t
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
    """Tambah 1 perenggan TC tersembunyi pada lokasi semasa (akan berada di permulaan kertas)."""
    p = doc.add_paragraph()
    # hide/vanish supaya tak nampak
    r_begin = OxmlElement("w:r"); rpr = OxmlElement("w:rPr"); vanish = OxmlElement("w:vanish"); rpr.append(vanish)
    fld_begin = OxmlElement("w:fldChar"); fld_begin.set(qn("w:fldCharType"), "begin")
    r_begin.append(rpr); r_begin.append(fld_begin); p._p.append(r_begin)

    r_instr = OxmlElement("w:r"); rpr2 = OxmlElement("w:rPr"); rpr2.append(OxmlElement("w:vanish"))
    instr = OxmlElement("w:instrText"); instr.set(qn("xml:space"), "preserve")
    safe = title.replace('"', "'")
    instr.text = f' TC "{safe}" \\l {level} '
    r_instr.append(rpr2); r_instr.append(instr); p._p.append(r_instr)

    r_end = OxmlElement("w:r"); rpr3 = OxmlElement("w:rPr"); rpr3.append(OxmlElement("w:vanish"))
    fld_end = OxmlElement("w:fldChar"); fld_end.set(qn("w:fldCharType"), "end")
    r_end.append(rpr3); r_end.append(fld_end); p._p.append(r_end)

def add_manual_toc_at_top(doc: Document, entries):
    """TOC: tajuk tengah, baris keluaran di bawahnya dengan dot leader + nombor kanan (PAGEREF TC handled by TOC field)."""
    # tajuk
    title_para = doc.add_paragraph()
    r = title_para.add_run("Table of Contents"); r.bold = True; r.font.size = Pt(14)
    title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # set tab kanan untuk baris manual (rujuk jika nak manual); untuk TOC berasaskan TC, tak perlu.
    # tapi kita akan biar Word format ikut styles apabila update.
    # sisip TOC field yang baca TC: { TOC \h \z \f "TC" }
    p = doc.add_paragraph()
    fld = OxmlElement("w:fldSimple")
    fld.set(qn("w:instr"), 'TOC \\h \\z \\f "TC"')
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
    """Letak footer 'Page X of Y' pada seksyen start_index, dan link seksyen seterusnya."""
    if len(doc.sections) == 0: return
    # TOC section tiada nombor
    doc.sections[0].different_first_page_header_footer = True
    # letak pada seksyen pertama kandungan
    if start_index < len(doc.sections):
        s = doc.sections[start_index]
        s.different_first_page_header_footer = False
        p = s.footer.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.add_run("Page "); add_field_run(p, "PAGE"); p.add_run(" of "); add_field_run(p, "NUMPAGES")
        # link others
        for i in range(start_index + 1, len(doc.sections)):
            doc.sections[i].footer.is_linked_to_previous = True
            doc.sections[i].different_first_page_header_footer = False

def set_update_fields_on_open(doc: Document):
    settings = doc.settings.element
    upd = OxmlElement("w:updateFields"); upd.set(qn("w:val"), "true")
    settings.append(upd)

# ---------- Core ----------
def combine_zip_with_toc_tc(zip_bytes: bytes) -> bytes:
    files = zip_docx_entries_in_order(zip_bytes)
    if not files:
        raise ValueError("ZIP tidak mengandungi .docx")

    meta = []
    for i, (name, blob) in enumerate(files, start=1):
        meta.append({"title": extract_title_from_doc_bytes(blob, name), "blob": blob})

    # 1) Base doc: TOC page
    base = Document()
    add_manual_toc_at_top(base, meta)
    base.add_page_break()  # pisah TOC dari kandungan

    composer = Composer(base)

    # 2) Untuk setiap doc:
    for idx, item in enumerate(meta):
        if idx > 0:
            base.add_page_break()  # setiap doc mula halaman baharu
        # letak TC tersembunyi di SINI (permulaan kertas idx)
        add_hidden_tc_paragraph(base, item["title"], level=1)
        # append sub-doc tanpa ubah format
        sub = Document(io.BytesIO(item["blob"]))
        composer.append(sub)

    # 3) Save sementara & buka semula untuk set page numbering & footer
    buf = io.BytesIO(); composer.save(buf); buf.seek(0)
    doc = Document(buf)

    # normalise numbering: mula 1 pada seksyen selepas TOC; lain-lain sambung
    clear_pgnumtype_for_all_sections(doc)
    if len(doc.sections) >= 2:
        start_numbering_at_section(doc, 1, 1)

    add_page_numbers_from_section(doc, start_index=1)
    set_update_fields_on_open(doc)

    out = io.BytesIO(); doc.save(out); out.seek(0)
    return out.read()

# ---------- UI ----------
st.subheader("Muat Naik ZIP Anda")
zip_file = st.file_uploader("Upload satu ZIP (folder + .docx) — susunan ikut folder (ZipInfo order).",
                            type=["zip"], accept_multiple_files=False)

st.info("TOC guna TC fields di permulaan setiap kertas (tajuk kiri, nombor kanan selepas Update Field). TOC tidak bernombor; nombor bermula 1 pada dokumen pertama dan bersambung.")
default_name = f"SPC_Proceedings_{datetime.now().strftime('%Y%m%d_%H%M')}.docx"
out_name = st.text_input("Nama fail output", value=default_name)

if st.button("🚀 Gabungkan (TOC TC-fields + setiap dokumen halaman baharu)"):
    try:
        if not zip_file:
            st.warning("Sila upload satu fail ZIP.")
        else:
            with st.spinner("Menggabungkan dokumen..."):
                compiled = combine_zip_with_toc_tc(zip_file.read())
            st.success("Siap! Muat turun di bawah. (Jika nombor belum muncul: Ctrl+A → F9 di Word)")
            st.download_button("⬇️ Muat Turun Fail Gabungan",
                               data=compiled,
                               file_name=out_name or "SPC_Proceedings.docx",
                               mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    except Exception as e:
        st.error("Ralat semasa menggabungkan dokumen.")
        st.exception(e)
