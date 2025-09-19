# Sistem Pengurusan Conference (SPC) — Streamlit

Aplikasi ini membolehkan anda memuat naik banyak fail **.docx** (atau satu **ZIP** berisi .docx),
kemudian menggabungkannya menjadi **satu fail .docx** lengkap dengan **Table of Contents (TOC)**
dan **nombor muka surat**.

## Cara Jalan
1. Pasang kebergantungan:
   ```bash
   pip install -r requirements.txt
   ```
2. Jalankan app:
   ```bash
   streamlit run app.py
   ```
3. Dalam UI, muat naik:
   - Banyak **.docx** sekali gus **ATAU**
   - Satu **.zip** berisi .docx
4. Klik **Gabungkan** dan muat turun fail akhir.
5. **Penting (Word/Office):** Selepas buka di Microsoft Word, **Right‑click** pada TOC → **Update Field** → **Update entire table**.

## Nota
- Hanya format **.docx** disokong (bukan .doc, .pdf).
- Format kompleks (rujuk silang, header/footer unik) mungkin perlu dilaras selepas gabungan.
- Susunan fail: nombor pada awal nama fail (jika ada) akan diutamakan, selain itu susunan abjad.
