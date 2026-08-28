import streamlit as st
from pathlib import Path
from io import BytesIO
import os, re, json, zipfile, xml.etree.ElementTree as ET
from docx import Document
from pypdf import PdfReader
import xlsxwriter

APP_DIR = Path(__file__).parent
SCHEMA = json.loads((APP_DIR / "schema.json").read_text(encoding="utf-8"))

st.set_page_config(page_title="Ayub Laporan", page_icon="📑", layout="wide")

st.markdown("""
<style>
.block-container {padding-top: 1.2rem; max-width: 1280px;}
.hero {padding: 22px 26px; border: 1px solid #dfe3e8; border-radius: 18px; margin-bottom: 16px;}
.hero h1 {margin:0 0 4px 0;}
.muted {color:#667085;}
div[data-testid="stFileUploader"] {border:1.5px dashed #98a2b3; border-radius:16px; padding:8px;}
</style>
<div class="hero">
<h1>📑 Ayub Laporan</h1>
<div class="muted">Upload dokumennya, bukan datanya. Bentuk source mailmerge + laporan Word dari satu proses.</div>
</div>
""", unsafe_allow_html=True)


def extract_docx(data: bytes) -> str:
    doc = Document(BytesIO(data))
    blocks = [p.text for p in doc.paragraphs if p.text.strip()]
    for table in doc.tables:
        for row in table.rows:
            vals = [c.text.strip() for c in row.cells]
            if any(vals):
                blocks.append(" | ".join(vals))
    return "\n".join(blocks)


def extract_pdf(data: bytes) -> str:
    reader = PdfReader(BytesIO(data))
    return "\n".join((p.extract_text() or "") for p in reader.pages)


def _xlsx_shared_strings(z):
    ns = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    if "xl/sharedStrings.xml" not in z.namelist():
        return []
    root = ET.fromstring(z.read("xl/sharedStrings.xml"))
    out = []
    for si in root.findall("m:si", ns):
        out.append("".join(t.text or "" for t in si.iter("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t")))
    return out


def extract_xlsx(data: bytes, max_rows_per_sheet=250) -> str:
    ns = {
        "m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    }
    with zipfile.ZipFile(BytesIO(data)) as z:
        ss = _xlsx_shared_strings(z)
        wb = ET.fromstring(z.read("xl/workbook.xml"))
        rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        relmap = {r.attrib["Id"]: r.attrib["Target"] for r in rels}
        chunks = []
        for sh in wb.find("m:sheets", ns):
            name = sh.attrib["name"]
            rid = sh.attrib["{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"]
            target = relmap[rid]
            if not target.startswith("xl/"):
                target = "xl/" + target.lstrip("/")
            root = ET.fromstring(z.read(target))
            chunks.append(f"\n### SHEET: {name}")
            rows = root.findall(".//m:sheetData/m:row", ns)[:max_rows_per_sheet]
            for row in rows:
                vals = []
                for c in row.findall("m:c", ns):
                    typ = c.attrib.get("t")
                    v = c.find("m:v", ns)
                    val = "" if v is None else (v.text or "")
                    if typ == "s" and val:
                        try:
                            val = ss[int(val)]
                        except Exception:
                            pass
                    elif typ == "inlineStr":
                        val = "".join(t.text or "" for t in c.findall(".//m:t", ns))
                    if val != "":
                        vals.append(f"{c.attrib.get('r','')}: {val}")
                if vals:
                    chunks.append(" | ".join(vals))
        return "\n".join(chunks)


def extract_text(uploaded):
    data = uploaded.getvalue()
    suffix = Path(uploaded.name).suffix.lower()
    if suffix == ".pdf": return extract_pdf(data)
    if suffix == ".docx": return extract_docx(data)
    if suffix == ".xlsx": return extract_xlsx(data)
    if suffix in [".txt", ".csv"]: return data.decode("utf-8", errors="ignore")
    return ""


def heuristic_extract(text):
    fields = {k: "" for k in SCHEMA["legacy_mailmerge_fields"] if k}
    evidence = {}
    patterns = {
        "no_pengumuman": r"\b(?:PENG|PENGUMUMAN)[-/\s]*[A-Z0-9./-]+",
        "realisasi_peserta": r"(?:realisasi peserta|peserta yang mengikuti)[^\d]{0,40}(\d{1,5})",
        "rencana_peserta": r"(?:direncanakan|rencana peserta)[^\d]{0,40}(\d{1,5})",
        "peserta_lulus": r"(?:dinyatakan lulus|peserta lulus)[^\d]{0,40}(\d{1,5})",
    }
    for key, pat in patterns.items():
        m = re.search(pat, text, flags=re.I)
        if m:
            fields[key] = m.group(1) if m.lastindex else m.group(0)
            evidence[key] = {"source": "heuristic", "confidence": 0.55, "note": "Fallback regex"}
    return fields, evidence


def get_api_key():
    typed = st.session_state.get("api_key", "")
    if typed:
        return typed
    try:
        return st.secrets.get("OPENAI_API_KEY", "")
    except Exception:
        return os.getenv("OPENAI_API_KEY", "")


def ai_extract(combined_text, file_names, model):
    from openai import OpenAI
    client = OpenAI(api_key=get_api_key())
    wanted = [x for x in SCHEMA["legacy_mailmerge_fields"] if x]
    prompt = f"""
Anda adalah mesin ekstraksi data dokumen kedinasan penyelenggaraan pembelajaran.
Rekonsiliasi dokumen menjadi source mailmerge.

Aturan:
1. Jangan mengarang; data yang tidak ditemukan harus string kosong.
2. Bedakan rencana dan realisasi.
3. Jika dokumen konflik, pilih bukti yang paling otoritatif/relevan dan catat konfliknya.
4. Untuk field yang terisi, sertakan nama sumber dan cuplikan/alasan singkat.
5. confidence 0..1.
6. Keluarkan JSON valid saja tanpa markdown.

Format:
{{"fields":{{"field":"value"}},"evidence":{{"field":{{"source":"file","confidence":0.9,"note":"bukti"}}}}}}

Field: {json.dumps(wanted, ensure_ascii=False)}
Nama file: {json.dumps(file_names, ensure_ascii=False)}
Isi dokumen:\n{combined_text[:150000]}
"""
    resp = client.responses.create(model=model, input=prompt)
    raw = resp.output_text.strip()
    raw = re.sub(r"^```(?:json)?\s*", "", raw)
    raw = re.sub(r"\s*```$", "", raw)
    obj = json.loads(raw)
    fields = {k: "" for k in wanted}
    for k, v in obj.get("fields", {}).items():
        if k in fields:
            fields[k] = "" if v is None else str(v)
    return fields, obj.get("evidence", {})


def replace_in_paragraph(paragraph, mapping):
    full = "".join(r.text for r in paragraph.runs)
    if not full:
        return
    replaced = full
    for key, value in mapping.items():
        value = "" if value is None else str(value)
        replaced = replaced.replace(f"<<{key}>>", value).replace(f"[@{key}]", value)
    if replaced != full:
        if paragraph.runs:
            paragraph.runs[0].text = replaced
            for r in paragraph.runs[1:]: r.text = ""
        else:
            paragraph.text = replaced


def build_docx(mapping, template_bytes=None):
    if template_bytes:
        doc = Document(BytesIO(template_bytes))
        for p in doc.paragraphs:
            replace_in_paragraph(p, mapping)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        replace_in_paragraph(p, mapping)
    else:
        doc = Document()
        doc.add_heading("Laporan Penyelenggaraan", 0)
        title = mapping.get("nama_pelatihan") or "Kegiatan Pembelajaran"
        doc.add_heading(title, level=1)
        for key in SCHEMA["core_fields"]:
            value = mapping.get(key, "")
            if value:
                p = doc.add_paragraph()
                p.add_run(key.replace("_", " ").title() + ": ").bold = True
                p.add_run(str(value))
        doc.add_paragraph("Dokumen ini dihasilkan otomatis dan perlu direview sebelum diunggah ke Nadine.")
    out = BytesIO(); doc.save(out); out.seek(0); return out


def build_xlsx(mapping, evidence):
    out = BytesIO()
    wb = xlsxwriter.Workbook(out, {"in_memory": True})
    ws = wb.add_worksheet("source_mailmerge")
    ev = wb.add_worksheet("evidence_log")
    hdr = wb.add_format({"bold": True, "bg_color": "#EAF2F8", "border": 1, "text_wrap": True})
    cell = wb.add_format({"border": 1, "valign": "top", "text_wrap": True})
    warn = wb.add_format({"border": 1, "bg_color": "#FFF4E5", "text_wrap": True})
    headers = [h for h in SCHEMA["legacy_mailmerge_fields"] if h]
    for c, h in enumerate(headers):
        ws.write(0, c, h, hdr); ws.write(1, c, mapping.get(h, ""), cell)
    ws.freeze_panes(1, 2); ws.set_column(0, len(headers)-1, 16)
    ev_headers = ["field", "value", "source_document", "confidence", "note", "status_review"]
    for c, h in enumerate(ev_headers): ev.write(0, c, h, hdr)
    for r, field in enumerate(headers, start=1):
        e = evidence.get(field, {}); value = mapping.get(field, ""); conf = e.get("confidence", "")
        try: high_conf = conf == "" or float(conf) >= 0.75
        except Exception: high_conf = False
        status = "OK" if value and high_conf else ("REVIEW" if value else "MISSING")
        vals = [field, value, e.get("source", ""), conf, e.get("note", ""), status]
        for c, v in enumerate(vals): ev.write(r, c, v, cell if status == "OK" else warn)
    ev.set_column("A:A", 30); ev.set_column("B:B", 34); ev.set_column("C:C", 30); ev.set_column("E:E", 60)
    wb.close(); out.seek(0); return out


with st.sidebar:
    st.subheader("⚙️ Mesin AI")
    ai_enabled = st.toggle("Gunakan AI", value=True)
    st.session_state["api_key"] = st.text_input("OpenAI API Key", type="password", value=st.session_state.get("api_key", ""), help="Tidak disimpan oleh aplikasi.")
    model = st.text_input("Model API", value=st.session_state.get("model", "gpt-5"))
    st.session_state["model"] = model
    st.caption("Jika server belum diberi secret API, key bisa dimasukkan sementara di sini.")

tab1, tab2, tab3, tab4 = st.tabs(["1 · Upload", "2 · Review Data", "3 · Evidence", "4 · Output"])

with tab1:
    st.subheader("Upload dokumen sumber")
    st.write("Boleh campur PDF, Word, Excel, CSV, atau TXT: pemanggilan, KAP/jadwal, presensi, rekap nilai, evaluasi, dan pengumuman hasil.")
    uploads = st.file_uploader("Drop semua dokumen untuk satu kegiatan", type=["pdf", "docx", "xlsx", "csv", "txt"], accept_multiple_files=True)
    template = st.file_uploader("Template Word laporan (opsional)", type=["docx"], help="Jika diisi, placeholder <<nama_field>> atau [@nama_field] akan diganti otomatis.")
    c1, c2, c3 = st.columns(3)
    nomor_nd = c1.text_input("Nomor ND/Laporan (opsional)")
    tanggal_nd = c2.text_input("Tanggal ND (opsional)")
    angkatan = c3.text_input("Angkatan (opsional)")
    if st.button("✨ Proses dokumen", type="primary", use_container_width=True, disabled=not uploads):
        texts = []
        prog = st.progress(0)
        for i, u in enumerate(uploads):
            try: t = extract_text(u)
            except Exception as e: t = f"[ERROR EKSTRAKSI: {e}]"
            texts.append(f"\n===== FILE: {u.name} =====\n{t}")
            prog.progress((i+1)/len(uploads))
        combined = "\n".join(texts)
        try:
            if ai_enabled and get_api_key():
                fields, evidence = ai_extract(combined, [u.name for u in uploads], model)
                mode = "AI"
            else:
                fields, evidence = heuristic_extract(combined); mode = "DEMO"
            if nomor_nd: fields["NomorND"] = nomor_nd
            if tanggal_nd: fields["TanggalND"] = tanggal_nd
            if angkatan: fields["akt"] = angkatan
            if fields.get("nama_pelatihan"): fields["nama_pelatihan_upper"] = fields["nama_pelatihan"].upper()
            st.session_state["fields"] = fields
            st.session_state["evidence"] = evidence
            st.session_state["template_bytes"] = template.getvalue() if template else None
            st.success(f"Ekstraksi selesai · Mode {mode}. Lanjut ke Review Data.")
        except Exception as e:
            st.error(f"Proses gagal: {e}")

with tab2:
    fields = st.session_state.get("fields", {})
    if not fields:
        st.info("Belum ada data. Upload dan proses dokumen terlebih dahulu.")
    else:
        core = SCHEMA["core_fields"]
        rows = [{"Field": k, "Nilai": fields.get(k, "")} for k in core]
        edited = st.data_editor(rows, use_container_width=True, hide_index=True, num_rows="fixed", column_config={"Field": st.column_config.TextColumn(disabled=True), "Nilai": st.column_config.TextColumn(width="large")})
        for row in edited: fields[row["Field"]] = row["Nilai"]
        st.session_state["fields"] = fields
        missing = [k for k in core if not fields.get(k)]
        c1, c2, c3 = st.columns(3)
        c1.metric("Field inti", len(core)); c2.metric("Terisi", len(core)-len(missing)); c3.metric("Perlu dilengkapi", len(missing))
        if missing: st.warning("Masih kosong: " + ", ".join(missing[:15]) + (" ..." if len(missing) > 15 else ""))

with tab3:
    fields = st.session_state.get("fields", {}); evidence = st.session_state.get("evidence", {})
    if not fields:
        st.info("Evidence akan muncul setelah dokumen diproses.")
    else:
        rows = []
        for k, v in fields.items():
            if v:
                e = evidence.get(k, {})
                rows.append({"Field": k, "Nilai": v, "Sumber": e.get("source", ""), "Confidence": e.get("confidence", ""), "Catatan": e.get("note", "")})
        st.dataframe(rows, use_container_width=True, hide_index=True)

with tab4:
    fields = st.session_state.get("fields", {}); evidence = st.session_state.get("evidence", {})
    if not fields:
        st.info("Belum ada output. Proses dokumen terlebih dahulu.")
    else:
        xlsx = build_xlsx(fields, evidence)
        docx = build_docx(fields, st.session_state.get("template_bytes"))
        name = re.sub(r"[^A-Za-z0-9_-]+", "_", fields.get("nama_pelatihan", "Laporan"))[:60].strip("_") or "Laporan"
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("#### ① Source Mailmerge")
            st.download_button("⬇️ Download Source Mailmerge (.xlsx)", data=xlsx, file_name=f"Source_Mailmerge_{name}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        with c2:
            st.markdown("#### ② Laporan Word")
            st.download_button("⬇️ Download Laporan (.docx)", data=docx, file_name=f"Laporan_Penyelenggaraan_{name}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
        st.caption("Gate sebelum Nadine: field kosong atau confidence rendah tetap harus direview manusia.")
