import * as XLSX from 'xlsx';
import * as pdfjsLib from 'pdfjs-dist';
import pdfWorker from 'pdfjs-dist/build/pdf.worker.min.mjs?url';

pdfjsLib.GlobalWorkerOptions.workerSrc = pdfWorker;

const clean = (v) => String(v ?? '').replace(/\s+/g, ' ').trim();
const num = (v) => {
  if (typeof v === 'number' && Number.isFinite(v)) return v;
  const n = Number(String(v ?? '').replace(',', '.'));
  return Number.isFinite(n) ? n : null;
};
const avg = (arr) => arr.length ? arr.reduce((a,b)=>a+b,0)/arr.length : null;
const fmt = (n, digits=2) => n == null ? '' : Number(n.toFixed(digits));

export function blankReport() {
  return {
    nama_pelatihan:'', akt:'', PIC:'', metode:'',
    'non_tatap_muka_(mulai)':'', 'non_tatap_muka_(selesai)':'',
    'tatap_muka_(mulai)':'', 'tatap_muka_(selesai)':'',
    'action_learning_(mulai)':'', 'action_learning_(selesai)':'',
    tgl_mulai:'', tgl_selesai:'', asrama:'', tgl_pengumuman:'', no_pengumuman:'',
    'deadline (H+25)':'', warning:'', warning_IKU:'', Nilai_IKU:'', no_laporan:'', tgl_laporan:'',
    unit_asal:'', tempat_pelatihan:'', alamat_tempat:'', rencana_peserta:'', realisasi_peserta:'',
    tidak_mengikuti:'', '#dgn_ket':'', '#tanpa_ket':'', ket_abstain:'',
    rencana_laki:'', rencana_perempuan:'', realisasi_laki:'', realisasi_perempuan:'',
    S3:'', S2:'', D4_S1:'', D3:'', D1:'', SMA_sederajat:'', Unknown_School:'',
    Jenis_Ujian:'', peserta_ujian:'', peserta_lulus:'', tidak_lulus:'', ket_tidak_lulus:'',
    tidak_memenuhi_syarat:'', ket_tidak_memenuhi_syarat:'',
    Min_Kompre_Pretest:'', Max_Kompre_Pretest:'', Rata2_Kompre_Pretest:'',
    Min_NA_Posttest:'', Max_NA_Posttest:'', Rata2_NA_Posttest:'', Selisih_Pretest_Posttest:'',
    predikat_A:'', predikat_B:'', predikat_C:'', lulus_min_B:'', persen_kelulusan_minB:'',
    indeks_materi_pembelajaran_elearning:'', indeks_bahan_ajar_elearning:'', indeks_metode_pembelajaran_elearning:'',
    indeks_kesigapan_penyelenggara_elearning:'', indeks_menu_elearning_mudah_digunakan:'', indeks_fasilitas_elearning_dapat_diakses:'',
    indeks_rata2_evagara_elearning:'', kriteria_idx_evagara_elearning:'', indeks_kemampuan_mengajar_elearning:'',
    indeks_rata2_evajar_elearning:'', kriteria_idx_evajar_elearning:'',
    indeks_materi_pembelajaran_PJJ:'', indeks_bahan_ajar_PJJ:'', indeks_metode_pembelajaran_PJJ:'',
    indeks_waktu_penyelenggaraan_PJJ:'', indeks_kesigapan_penyelenggara_PJJ:'', indeks_waktu_ujian_PJJ:'',
    indeks_fasilitas_PJJ_mudah_diakses:'', indeks_fasilitas_PJJ_mudah_digunakan:'', indeks_rata2_evagara_PJJ:'',
    kriteria_idx_evagara_PJJ:'', indeks_kemampuan_mengajar:'', indeks_rata2_evajar:'', kriteria_idx_evajar:'',
    indeks_materi_pembelajaran_klasikal:'', indeks_bahan_ajar_klasikal:'', indeks_metode_pembelajaran_klasikal:'',
    indeks_kesigapan_penyelenggara_klasikal:'', indeks_konsumsi:'', indeks_ruang_kelas:'', indeks_asrama:'',
    indeks_rata2_evagara_klasikal:'', kriteria_idx_evagara_klasikal:'', indeks_pengetahuan_pengajar_klasikal:'',
    indeks_kemampuan_mengajar_klasikal:'', indeks_rata2_evajar_klasikal:'', kriteria_idx_evajar_klasikal:'',
    nama_pelatihan_upper:'', maksud_program:'', diharapkan_mampu:'', rec_pengajar:'', rec_renbang:'', rec_penye:'',
    rec_evalap:'', rec_tu:'', rec_penyelenggaraan_pelatihan:'', nama_pelatihan_lengkap:'', END:'', DS:'', Unggah_MsTeams:'',
    Keterangan:'', NomorND:'', TanggalND:''
  };
}

function findHeaderRow(rows, requiredGroups) {
  let best = {idx:-1, score:0};
  rows.forEach((row, idx) => {
    const cells = row.map(v=>clean(v).toUpperCase());
    let score = 0;
    requiredGroups.forEach(group => {
      if (group.some(k => cells.some(c => c === k || c.includes(k)))) score++;
    });
    if (score > best.score) best = {idx, score};
  });
  return best.score >= Math.max(2, Math.ceil(requiredGroups.length/2)) ? best.idx : -1;
}

function rowsToObjects(rows, headerIndex) {
  if (headerIndex < 0) return [];
  const headers = rows[headerIndex].map((h,i)=> clean(h) || `COL_${i}`);
  return rows.slice(headerIndex+1).map(row => {
    const o={}; headers.forEach((h,i)=>o[h]=row[i]); return o;
  });
}

function getByHeader(obj, variants) {
  const keys = Object.keys(obj);
  const key = keys.find(k => variants.some(v => clean(k).toUpperCase() === v.toUpperCase()))
    || keys.find(k => variants.some(v => clean(k).toUpperCase().includes(v.toUpperCase())));
  return key ? obj[key] : undefined;
}

function educationBucket(v) {
  const s=clean(v).toUpperCase().replace(/\./g,'').replace(/\s+/g,' ');
  if (!s) return 'Unknown_School';
  if (/S\s*3|DOKTOR/.test(s)) return 'S3';
  if (/S\s*2|MAGISTER/.test(s)) return 'S2';
  if (/D\s*IV|D4|S\s*1|S1|SARJANA/.test(s)) return 'D4_S1';
  if (/D\s*III|D3/.test(s)) return 'D3';
  if (/D\s*I\b|D1\b/.test(s)) return 'D1';
  if (/SMA|SMK|SLTA/.test(s)) return 'SMA_sederajat';
  return 'Unknown_School';
}

function classifyPredikat(score) {
  if (score == null) return null;
  if (score >= 90) return 'A';
  if (score >= 80) return 'B';
  if (score >= 70) return 'C';
  return 'D';
}

async function parseWorkbook(file) {
  const ab=await file.arrayBuffer();
  const wb=XLSX.read(ab,{type:'array',cellDates:false});
  let all=[];
  for (const sheetName of wb.SheetNames) {
    const ws=wb.Sheets[sheetName];
    const rows=XLSX.utils.sheet_to_json(ws,{header:1,defval:null,raw:true});
    if (rows.some(r=>r.some(v=>v!=null))) all.push({sheetName,rows});
  }

  const result={type:'xlsx', source:file.name, report:{}, details:{}};
  for (const {sheetName,rows} of all) {
    const participantHeader=findHeaderRow(rows,[['NO'],['NAMA'],['NIP','NIP/NRP/NIK'],['GENDER'],['PENDIDIKAN']]);
    if (participantHeader>=0) {
      const objs=rowsToObjects(rows,participantHeader).filter(o=>num(getByHeader(o,['NO']))!=null && clean(getByHeader(o,['NAMA'])));
      if (objs.length) {
        const genders=objs.map(o=>clean(getByHeader(o,['GENDER'])).toUpperCase());
        const edu=objs.map(o=>educationBucket(getByHeader(o,['PENDIDIKAN'])));
        const counts=(arr,val)=>arr.filter(x=>x===val).length;
        result.details.participants=objs;
        Object.assign(result.report,{
          realisasi_peserta:objs.length,
          realisasi_laki:counts(genders,'L'),
          realisasi_perempuan:counts(genders,'P'),
          S3:counts(edu,'S3'), S2:counts(edu,'S2'), D4_S1:counts(edu,'D4_S1'), D3:counts(edu,'D3'),
          D1:counts(edu,'D1'), SMA_sederajat:counts(edu,'SMA_sederajat'), Unknown_School:counts(edu,'Unknown_School')
        });
        const title=rows.slice(0,participantHeader).flat().map(clean).find(x=>/pelatihan|pjj/i.test(x));
        if (title) result.report.nama_pelatihan_lengkap=title;
        continue;
      }
    }

    const scoreHeader=findHeaderRow(rows,[['NO'],['NAMA'],['NIP'],['PRETEST'],['POSTTEST']]);
    if (scoreHeader>=0) {
      const objs=rowsToObjects(rows,scoreHeader).filter(o=>num(getByHeader(o,['NO']))!=null && clean(getByHeader(o,['NAMA'])));
      const pre=objs.map(o=>num(getByHeader(o,['PRETEST']))).filter(v=>v!=null);
      const post=objs.map(o=>num(getByHeader(o,['POSTTEST']))).filter(v=>v!=null);
      if (pre.length && post.length) {
        const pred=post.map(classifyPredikat);
        const a=pred.filter(x=>x==='A').length, b=pred.filter(x=>x==='B').length, c=pred.filter(x=>x==='C').length;
        const minB=post.filter(x=>x>=80).length;
        result.details.scores=objs;
        Object.assign(result.report,{
          Jenis_Ujian:'Pretest dan Posttest', peserta_ujian:objs.length,
          Min_Kompre_Pretest:Math.min(...pre), Max_Kompre_Pretest:Math.max(...pre), Rata2_Kompre_Pretest:fmt(avg(pre)),
          Min_NA_Posttest:Math.min(...post), Max_NA_Posttest:Math.max(...post), Rata2_NA_Posttest:fmt(avg(post)),
          Selisih_Pretest_Posttest:fmt(avg(post)-avg(pre)), predikat_A:a, predikat_B:b, predikat_C:c,
          lulus_min_B:minB, persen_kelulusan_minB:fmt(minB/objs.length*100)
        });
      }
    }
  }
  return result;
}

async function extractPdfText(file) {
  const data=new Uint8Array(await file.arrayBuffer());
  const pdf=await pdfjsLib.getDocument({data}).promise;
  const pages=[];
  for (let p=1;p<=pdf.numPages;p++) {
    const page=await pdf.getPage(p);
    const content=await page.getTextContent();
    const text=content.items.map(x=>x.str).join(' ');
    pages.push(clean(text));
  }
  return pages;
}

const first = (text, regex, group=1) => {
  const m=text.match(regex); return m ? clean(m[group]) : '';
};

function parsePdfData(file,pages) {
  const text=pages.join(' \n ');
  const lower=text.toLowerCase();
  const report={};
  let docType='pdf_lainnya';

  if (/kerangka acuan program/i.test(text)) docType='kap';
  else if (/rekomendasi perbaikan/i.test(text)) docType='rekomendasi_perbaikan';
  else if (/rekapitulasi kehadiran pengajar|rekapitulasi kehadiran.*peserta/i.test(text)) docType='rekap_kehadiran';
  else if (/pengumuman/i.test(text) && /yang dinyatakan telah mengikuti/i.test(text)) docType='pengumuman';

  const fullName=first(text,/(Pelatihan(?: Jarak Jauh)?[^\n]{5,180}?)(?= Tahun Anggaran| TA 20\d\d| yang diselenggarakan|\s+Penyelenggara\s*:|\s+Lokasi\/Tanggal)/i);
  if (fullName) {
    report.nama_pelatihan_lengkap=fullName.replace(/\s+/g,' ').trim();
    report.nama_pelatihan=report.nama_pelatihan_lengkap.replace(/\s+Angkatan\s+[IVXLC\d]+.*$/i,'').trim();
    report.nama_pelatihan_upper=report.nama_pelatihan.toUpperCase();
    const akt=first(report.nama_pelatihan_lengkap,/Angkatan\s+([IVXLC\d]+)/i); if(akt) report.akt=akt;
  }

  const startEnd=text.match(/(?:dilaksanakan pada tanggal|Waktu Penyelenggaraan\s*:?|Lokasi\/Tanggal Penyelenggaraan\s*:?[^\n]*?\/?)\s*(\d{1,2}\s+[A-Za-z]+\s+20\d{2})\s*(?:s\.d\.|sampai dengan|-)\s*(\d{1,2}\s+[A-Za-z]+\s+20\d{2})/i);
  if (startEnd) { report.tgl_mulai=clean(startEnd[1]); report.tgl_selesai=clean(startEnd[2]); }

  const manager=first(text,/Manajer Kelas\s*:?\s*([A-Za-z.' ]{3,80}?)(?=\s+(?:No\s+Mata Pelajaran|REKAPITULASI|\d+\s+Proses|$))/i);
  if (manager) report.PIC=manager;
  if (/pelatihan jarak jauh/i.test(text)) report.metode='Pelatihan Jarak Jauh';
  if (report.metode && report.tgl_mulai) { report['non_tatap_muka_(mulai)']=report.tgl_mulai; report['non_tatap_muka_(selesai)']=report.tgl_selesai; }

  if (docType==='pengumuman') {
    const no=first(text,/NOMOR\s+(PENG-[A-Z0-9.\/-]+)/i); if(no) report.no_pengumuman=no;
    const dt=first(text,/pada tanggal\s+(\d{1,2}\s+[A-Za-z]+\s+20\d{2})/i); if(dt) report.tgl_pengumuman=dt;
    const declared=(text.match(/\b\d+\s+[A-Z][A-Za-z.' -]+\s+\d{8}\s+\d{6}/g)||[]).length;
    if(declared) report.realisasi_peserta=declared;
  }

  if (docType==='rekomendasi_perbaikan') {
    const pcount=first(text,/Jumlah Peserta Pembelajaran\s*:?\s*(\d+)/i); if(pcount) report.realisasi_peserta=Number(pcount);
    const evalAvg=first(text,/nilai rata-rata kinerja sebesar\s*([0-9]+[,.][0-9]+)/i);
    if(evalAvg) { report.indeks_rata2_evagara_PJJ=Number(evalAvg.replace(',','.')); report.kriteria_idx_evagara_PJJ=/sangat baik/i.test(text)?'Sangat Baik':''; }
    const teacherAvg=first(text,/kemampuan pengajar[\s\S]{0,180}?sebesar\s*([0-9]+[,.][0-9]+)/i);
    if(teacherAvg) { report.indeks_kemampuan_mengajar=Number(teacherAvg.replace(',','.')); report.indeks_rata2_evajar=Number(teacherAvg.replace(',','.')); report.kriteria_idx_evajar=/sangat baik/i.test(text)?'Sangat Baik':''; }
    const rec=first(text,/butir yang perlu diperbaiki,?\s*([^\n]{20,500}?)(?=\s*2\.|Untuk hasil evaluasi pengajar|Jakarta,)/i);
    if(rec) report.rec_penyelenggaraan_pelatihan=rec;
  }

  if (docType==='kap') {
    const tujuan=first(text,/TUJUAN PROGRAM\s+([\s\S]*?)(?=KEBUTUHAN STRATEGIS UNIT PENGGUNA YANG AKAN DICAPAI)/i);
    if(tujuan) report.maksud_program=tujuan;
    const sasaran=first(text,/SASARAN \(TARGET LEARNERS\)\s+([\s\S]*?)(?=MODEL PEMBELAJARAN)/i);
    if(sasaran) report.diharapkan_mampu=sasaran;
    const totalJP=first(text,/TOTAL JP\s+(\d+)/i); if(totalJP) report.Keterangan=`Total ${totalJP} JP`;
  }

  if (/pusdiklat pajak/i.test(text)) report.tempat_pelatihan='Pusdiklat Pajak';
  if (/direktorat jenderal pajak/i.test(text)) report.unit_asal='Direktorat Jenderal Pajak';

  return {type:docType, source:file.name, report, details:{pages:pages.length, textPreview:text.slice(0,1200)}};
}

export async function parseFile(file) {
  const ext=file.name.split('.').pop().toLowerCase();
  if (['xlsx','xls'].includes(ext)) return parseWorkbook(file);
  if (ext==='pdf') return parsePdfData(file,await extractPdfText(file));
  throw new Error(`Format .${ext} belum didukung`);
}

function chooseValue(current,incoming) {
  if (incoming===undefined || incoming===null || incoming==='') return current;
  if (current===undefined || current===null || current==='') return incoming;
  return current;
}

export function aggregate(parsedDocs) {
  const report=blankReport();
  for (const doc of parsedDocs) {
    for (const [k,v] of Object.entries(doc.report||{})) report[k]=chooseValue(report[k],v);
  }
  if (!report.nama_pelatihan && report.nama_pelatihan_lengkap)
    report.nama_pelatihan=report.nama_pelatihan_lengkap.replace(/\s+Angkatan\s+[IVXLC\d]+.*$/i,'').trim();
  if (!report.nama_pelatihan_upper && report.nama_pelatihan) report.nama_pelatihan_upper=report.nama_pelatihan.toUpperCase();
  if (!report.akt && report.nama_pelatihan_lengkap) report.akt=first(report.nama_pelatihan_lengkap,/Angkatan\s+([IVXLC\d]+)/i);
  if (report.metode==='Pelatihan Jarak Jauh') {
    report['non_tatap_muka_(mulai)'] ||= report.tgl_mulai;
    report['non_tatap_muka_(selesai)'] ||= report.tgl_selesai;
  }
  return report;
}

export function reportStats(report) {
  const entries=Object.entries(report);
  const filled=entries.filter(([,v])=>v!=='' && v!==null && v!==undefined).length;
  return {filled,total:entries.length,percent:Math.round(filled/entries.length*100)};
}
