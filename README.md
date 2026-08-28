# Ayub Laporan — Laporan Penyelenggaraan AI

Web MVP untuk menyusun laporan penyelenggaraan dari dokumen sumber.

Konsep utama: **upload dokumennya, bukan datanya**.

## Flow
1. User drag & drop dokumen sumber untuk satu kegiatan.
2. Aplikasi membaca PDF/DOCX/XLSX/CSV/TXT.
3. AI mengekstrak dan merekonsiliasi data.
4. User mereview field yang meragukan.
5. Aplikasi menghasilkan Source Mailmerge XLSX dan Laporan Penyelenggaraan DOCX.

## Deploy ke Streamlit Community Cloud
- Main file: `app.py`
- Dependencies: `requirements.txt`
- Tambahkan `OPENAI_API_KEY` di Streamlit Secrets bila ingin mode AI tanpa memasukkan key setiap kali.

MVP ini belum memiliki login, database, SSO, penyimpanan permanen, atau integrasi Nadine.