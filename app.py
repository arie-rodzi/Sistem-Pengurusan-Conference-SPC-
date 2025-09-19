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
st.caption("TOC manual di atas. Setiap dokumen bermula halaman baharu. TOC tiada nombor; nombor bermula 1 pada dokumen pertama. Kandungan asal tidak diubah.")

# ================= helpers =================

def zip_docx_entries_in_order(zip_bytes: bytes):
    """Pulangkan [(name_in_zip, bytes)] ikut susunan dalam ZIP (ZipInfo order)."""
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
    """Ambil tajuk daripada Heading 1 jika wujud; jika tiada, guna nama fail (tanpa .docx)."""
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
    """Sisip Word field (PAGE / NUMPAGES / PAGEREF ...) ke dalam paragraph sedia ada."""
    r1 = OxmlElement("w:r")
    fc1 = OxmlElement("w:fldChar"); fc1.set(qn("w:fldCharType"), "begin")
    r1.append(fc1); paragraph._p.append(r1)

    r2 = OxmlElement("w:r")
    it = OxmlElement("w:instrText"); it.set(qn("xml:space"), "preserve"); it.text = f" {field_code} "
    r2.append(it); paragraph._p.append(r2)

    r3 = OxmlElement("w:r")
    fc3 = OxmlElement("w:fldChar"); fc3.set(qn("w:fldCharType"), "end")
    r3.append(fc3); paragraph._p.append(r3)

def add_bookmark(paragraph, name: str):
    """Letak bookmark pada paragraph (untuk rujukan PAGEREF dalam TOC)."""
    start = OxmlElement("w:bookmarkStart"); start.set(qn("w:id"), "0"); start.set(qn("w:name"), name)
    end   = OxmlElement("w:bookmarkEnd");   end.set(qn("w:id"), "0")
    paragraph._p.insert(0, start); paragraph._p.append(end)

def _new_para_after(doc: Document, anchor_para):
    """Buat perenggan baharu dan sisip SELEPAS perenggan 'anchor_para'."""
    new_para = doc.add_paragraph()          # create temporary
    anchor_para._p.addnext(new_para._p)     # move right after anchor
    return new_para

def set_update_fields_on_open(doc: Document):
    """Paksa Word auto-refresh semua field (PAGE/NUMPAGES/PAGEREF) ketika buka dokumen."""
    settings = doc.settings.element
    upd = OxmlElement("w:updateFields"); upd.set(qn("w:val"), "true")
    settings.append(upd)

def add_manual_toc_at_top(doc: Document, toc_entries):
    """
    Sisip TOC manual tepat di bawah tajuk:
    - Tajuk 'Table of Contents' (center)
    - Setiap entri: Tajuk ... (dot leaders) ... PAGEREF bookmark (right-aligned)
    """
    # Pastikan ada perenggan pertama sebagai tajuk
    if len(doc.paragraphs) == 0:
        doc.add_paragraph()
    title_para = doc.paragraphs[0]
    title_para.clear()
    run = title_para.add_run("Table of Contents"); run.bold = True; run.font.size = Pt(14)
    title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Lebar boleh guna (EMU) → tukar kepada inci
    sec = doc.sections[0]
    usable_width_emu = sec.page_width - sec.left_margin - sec.right_margin  # integer EMU
    usable_width_inch = usable_width_emu / 914400.0  # 1 inch = 914,400 EMU

    # Sisip satu baris kosong SELEPAS tajuk sebagai sauh
    anchor = _new_para_after(doc, title_para)

    # Tambah setiap baris TOC tepat selepas anchor (berturutan)
    last_para = anchor
    for e in toc_entries:
        p = _new_para_after(doc, last_para)
        # Tab kanan pada had lebar kandungan; dot leaders
        p.paragraph_format.tab_stops.add_tab_stop(
            Inches(usable_width_inch), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS
        )
        p.add_run(e["title"])
        p.add_run("\t")
        add_field_run(p, f'PAGEREF {e["bookmark"]} \\h')
        last_para = p

def add_page_numbers_selected_sections(doc: Document):
    """
    Tambah 'Page X of Y' di footer (tengah) untuk SEMUA seksyen KECUALI seksyen 0 (TOC).
    Juga set seksyen 0 'different first page' supaya tiada nombor pada halaman TOC.
    """
    # seksyen TOC: tiada page number
    if len(doc.sections) > 0:
        doc.sections[0].different_first_page_header_footer = True

    for idx, section in enumerate(doc.sections):
        if idx == 0:
            continue  # skip TOC section
        p = section.footer.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.add_run("Page "); add_field_run(p, "PAGE"); p.add_run(" of "); add_field_run(p, "NUMPAGES")

def normalize_page_numbering(doc: Document):
    """
    Pastikan page numbering:
    - restart = 1 pada seksyen pertama SELEPAS TOC,
    - seksyen seterusnya teruskan numbering (tiada restart).
    """
    # buang sebarang pgNumType sedia ada pada semua seksyen
    for section in doc.sections:
        sectPr = section._sectPr
        pgNumType = sectPr.find(qn('w:pgNumType'))
        if pgNumType is not None:
            sectPr.remove(pgNumType)

    # jika wujud seksyen selepas TOC, set start=1 di situ
    if len(doc.sections) > 1:
        sectPr1 = doc.sections[1]._sectPr
        pgNumType = OxmlElement('w:pgNumType')
        pgNumType.set(qn('w:start'), '1')
        sectPr1.append(pgNumType)

# --------------- core merge ---------------

def combine_with_manual_toc(zip_bytes: bytes) -> bytes:
    files = zip_docx_entries_in_order(zip_bytes)
    if not files:
        raise ValueError("ZIP tidak mengandungi .docx")

    # Sediakan senarai tajuk + bookmark
    toc = []
    for i, (name, blob) in enumerate(files, start=1):
        title = extract_title_from_doc_bytes(blob, name)
        toc.append({"title": title, "bookmark": f"DOC_{i}", "blob": blob})

    # Dokumen asas: p0 akan menjadi tajuk TOC
    base = Document()
    base.add_paragraph()      # p0 = tajuk TOC
    base.add_page_break()     # pisahkan TOC daripada kandungan

    composer = Composer(base)
    for i, item in enumerate(toc, start=1):
        if i > 1:
            base.add_page_break()  # setiap dokumen bermula halaman baharu
        # paragraph untuk bookmark di permulaan dokumen i
        bm_para = base.add_paragraph()
        add_bookmark(bm_para, item["bookmark"])
        # append sub-doc TANPA ubah format
        sub = Document(io.BytesIO(item["blob"]))
        composer.append(sub)

    # Simpan gabungan sementara
    buf = io.BytesIO(); composer.save(buf); buf.seek(0)

    # Buka semula → sisip TOC manual di atas + set page numbering
    doc = Document(buf)
    add_manual_toc_at_top(doc, [{"title": x["title"], "bookmark": x["bookmark"]} for x in toc])
    normalize_page_numbering(doc)           # <-- penting: elak semua '1', mula 1 selepas TOC
    add_page_numbers_selected_sections(doc) # <-- TOC tiada nombor; dokumen ada nombor
    set_update_fields_on_open(doc)

    out = io.BytesIO(); doc.save(out); out.seek(0)
    return out.read()

# ================= UI =================

st.subheader("Muat Naik ZIP Anda")
zip_file = st.file_uploader(
    "Upload satu ZIP (folder + .docx) — susunan ikut folder (ZipInfo order).",
    type=["zip"], accept_multiple_files=False
)

st.info("TOC manual: tajuk kiri, nombor kanan (dot leaders). TOC tidak bernombor; nombor bermula 1 pada dokumen pertama.")
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
