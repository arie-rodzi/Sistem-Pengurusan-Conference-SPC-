import io, os, zipfile
from datetime import datetime

import streamlit as st
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.shared import Inches, Pt
from docxcompose.composer import Composer

st.set_page_config(page_title="SPC — Merge DOCX + Manual TOC", layout="wide")
st.title("📚 SPC — Merge DOCX (PDF-like) + TOC Tajuk Kiri / Nombor Kanan")
st.caption("TOC di atas. Setiap dokumen bermula halaman baharu. TOC tiada nombor; nombor bermula 1 pada dokumen pertama dan bersambung.")

# -------- Helpers --------
def zip_docx_entries_in_order(zip_bytes: bytes):
    out = []
    with zipfile.ZipFile(io.BytesIO(zip_bytes)) as zf:
        for info in zf.infolist():
            if info.is_dir(): continue
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
    r1 = OxmlElement("w:r"); fc1 = OxmlElement("w:fldChar"); fc1.set(qn("w:fldCharType"), "begin"); r1.append(fc1); paragraph._p.append(r1)
    r2 = OxmlElement("w:r"); it = OxmlElement("w:instrText"); it.set(qn("xml:space"), "preserve"); it.text = f" {field_code} "; r2.append(it); paragraph._p.append(r2)
    r3 = OxmlElement("w:r"); fc3 = OxmlElement("w:fldChar"); fc3.set(qn("w:fldCharType"), "end"); r3.append(fc3); paragraph._p.append(r3)

def add_bookmark_to_paragraph(paragraph, name: str, bid: int):
    start = OxmlElement("w:bookmarkStart"); start.set(qn("w:id"), str(bid)); start.set(qn("w:name"), name)
    end   = OxmlElement("w:bookmarkEnd");   end.set(qn("w:id"), str(bid))
    paragraph._p.insert(0, start); paragraph._p.append(end)

def _new_para_after(doc: Document, anchor_para):
    new_para = doc.add_paragraph()
    anchor_para._p.addnext(new_para._p)
    return new_para

def set_update_fields_on_open(doc: Document):
    settings = doc.settings.element
    upd = OxmlElement("w:updateFields"); upd.set(qn("w:val"), "true")
    settings.append(upd)

def add_manual_toc_at_top(doc: Document, toc_entries):
    # pastikan perenggan pertama ada
    if len(doc.paragraphs) == 0:
        doc.add_paragraph()
    title_para = doc.paragraphs[0]
    title_para.clear()
    run = title_para.add_run("Table of Contents"); run.bold = True; run.font.size = Pt(14)
    title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # lebar boleh guna (EMU → inci)
    sec = doc.sections[0]
    usable_width_inch = (sec.page_width - sec.left_margin - sec.right_margin) / 914400.0

    anchor = _new_para_after(doc, title_para)
    last_para = anchor
    for e in toc_entries:
        p = _new_para_after(doc, last_para)
        p.paragraph_format.tab_stops.add_tab_stop(
            Inches(usable_width_inch), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS
        )
        p.add_run(e["title"])
        p.add_run("\t")
        add_field_run(p, f'PAGEREF {e["bookmark"]} \\h')
        last_para = p

def clear_pgnumtype_for_all_sections(doc: Document):
    for section in doc.sections:
        sectPr = section._sectPr
        pgNumType = sectPr.find(qn('w:pgNumType'))
        if pgNumType is not None:
            sectPr.remove(pgNumType)

def start_numbering_at_section(doc: Document, section_index: int, start_at: int = 1):
    if section_index < len(doc.sections):
        sectPr = doc.sections[section_index]._sectPr
        pgNumType = OxmlElement('w:pgNumType')
        pgNumType.set(qn('w:start'), str(start_at))
        sectPr.append(pgNumType)

def add_page_numbers_linking_from_section1(doc: Document):
    # Seksyen 0 = TOC (tiada nombor). Letak nombor pada seksyen 1 dan link seterusnya.
    if len(doc.sections) == 0: return
    doc.sections[0].different_first_page_header_footer = True

    if len(doc.sections) >= 2:
        s1 = doc.sections[1]
        s1.different_first_page_header_footer = False
        p = s1.footer.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.add_run("Page "); add_field_run(p, "PAGE"); p.add_run(" of "); add_field_run(p, "NUMPAGES")
        for i in range(2, len(doc.sections)):
            doc.sections[i].footer.is_linked_to_previous = True
            doc.sections[i].different_first_page_header_footer = False

# -------- Core merge --------
def combine_with_manual_toc(zip_bytes: bytes) -> bytes:
    files = zip_docx_entries_in_order(zip_bytes)
    if not files:
        raise ValueError("ZIP tidak mengandungi .docx")

    # sediakan tajuk + nama bookmark
    toc_meta = []
    for i, (name, blob) in enumerate(files, start=1):
        toc_meta.append({"title": extract_title_from_doc_bytes(blob, name),
                         "bookmark": f"DOC_{i}",
                         "blob": blob})

    # dokumen asas: p0 = tajuk TOC → page break
    base = Document()
    base.add_paragraph()
    base.add_page_break()

    composer = Composer(base)
    # penting: letak BOOKMARK DI DALAM SUB-DOC (bukan di base) supaya PAGEREF tunjuk halaman sebenar
    bookmark_id = 1
    for idx, item in enumerate(toc_meta, start=1):
        if idx > 1:
            base.add_page_break()  # setiap dokumen mula halaman baharu

        sub = Document(io.BytesIO(item["blob"]))
        # pastikan ada perenggan pertama
        if len(sub.paragraphs) == 0:
            sub.add_paragraph()
        # tambah bookmark DOC_i pada perenggan pertama sub-doc
        add_bookmark_to_paragraph(sub.paragraphs[0], item["bookmark"], bookmark_id)
        bookmark_id += 1

        composer.append(sub)

    # simpan gabungan
    buf = io.BytesIO(); composer.save(buf); buf.seek(0)

    # buka semula → TOC + normalise numbering + page numbers
    doc = Document(buf)
    add_manual_toc_at_top(doc, [{"title": x["title"], "bookmark": x["bookmark"]} for x in toc_meta])

    # buang restart, mula 1 di seksyen pertama selepas TOC
    clear_pgnumtype_for_all_sections(doc)
    if len(doc.sections) >= 2:
        start_numbering_at_section(doc, 1, 1)

    # letak nombor pada seksyen 1 dan link seksyen2+
    add_page_numbers_linking_from_section1(doc)

    set_update_fields_on_open(doc)

    out = io.BytesIO(); doc.save(out); out.seek(0)
    return out.read()

# -------- UI --------
st.subheader("Muat Naik ZIP Anda")
zip_file = st.file_uploader(
    "Upload satu ZIP (folder + .docx) — susunan ikut folder (ZipInfo order).",
    type=["zip"], accept_multiple_files=False
)

st.info("TOC manual: tajuk kiri, nombor kanan (dot leaders). TOC tidak bernombor; nombor bermula 1 pada dokumen pertama dan bersambung.")
default_name = f"SPC_Proceedings_{datetime.now().strftime('%Y%m%d_%H%M')}.docx"
out_name = st.text_input("Nama fail output", value=default_name)

if st.button("🚀 Gabungkan (TOC manual + setiap dokumen halaman baharu)"):
    try:
        if not zip_file:
            st.warning("Sila upload satu fail ZIP.")
        else:
            with st.spinner("Menggabungkan dokumen..."):
                compiled = combine_with_manual_toc(zip_file.read())
            st.success("Siap! Muat turun di bawah.")
            st.download_button(
                "⬇️ Muat Turun Fail Gabungan",
                data=compiled,
                file_name=out_name or "SPC_Proceedings.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )
    except Exception as e:
        st.error("Ralat semasa menggabungkan dokumen.")
        st.exception(e)
