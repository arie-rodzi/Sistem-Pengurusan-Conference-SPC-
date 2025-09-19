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
