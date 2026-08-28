import React, {useMemo, useState} from 'react';
import {aggregate, parseFile, reportStats} from './parser.js';

const LABELS={
  nama_pelatihan:'Nama pelatihan', akt:'Angkatan', PIC:'PIC / Manajer Kelas', metode:'Metode',
  tgl_mulai:'Tanggal mulai', tgl_selesai:'Tanggal selesai', tgl_pengumuman:'Tanggal pengumuman', no_pengumuman:'Nomor pengumuman',
  unit_asal:'Unit asal', tempat_pelatihan:'Tempat pelatihan', rencana_peserta:'Rencana peserta', realisasi_peserta:'Realisasi peserta',
  realisasi_laki:'Laki-laki', realisasi_perempuan:'Perempuan', S3:'S3', S2:'S2', D4_S1:'D4/S1', D3:'D3', D1:'D1', SMA_sederajat:'SMA/sederajat', Unknown_School:'Pendidikan tidak diketahui',
  Jenis_Ujian:'Jenis ujian', peserta_ujian:'Peserta ujian', Min_Kompre_Pretest:'Min pretest', Max_Kompre_Pretest:'Max pretest', Rata2_Kompre_Pretest:'Rata-rata pretest',
  Min_NA_Posttest:'Min posttest', Max_NA_Posttest:'Max posttest', Rata2_NA_Posttest:'Rata-rata posttest', Selisih_Pretest_Posttest:'Selisih rata-rata pre/post',
  predikat_A:'Predikat A', predikat_B:'Predikat B', predikat_C:'Predikat C', lulus_min_B:'Minimal B', persen_kelulusan_minB:'% minimal B',
  indeks_rata2_evagara_PJJ:'Rata-rata evaluasi penyelenggaraan PJJ', indeks_rata2_evajar:'Rata-rata evaluasi pengajar',
  maksud_program:'Maksud/tujuan program', diharapkan_mampu:'Sasaran peserta', rec_penyelenggaraan_pelatihan:'Rekomendasi penyelenggaraan',
  nama_pelatihan_lengkap:'Nama pelatihan lengkap', Keterangan:'Keterangan'
};

const DOC_NAMES={kap:'KAP', rekomendasi_perbaikan:'Rekomendasi Perbaikan', rekap_kehadiran:'Rekap Kehadiran', pengumuman:'Pengumuman', xlsx:'Excel', pdf_lainnya:'PDF lain'};

function Section({title,keys,report,onChange}){
  return <section className="card">
    <h2>{title}</h2>
    <div className="grid">
      {keys.map(k=><label className={['maksud_program','diharapkan_mampu','rec_penyelenggaraan_pelatihan'].includes(k)?'wide':''} key={k}>
        <span>{LABELS[k]||k}</span>
        {['maksud_program','diharapkan_mampu','rec_penyelenggaraan_pelatihan'].includes(k)
          ? <textarea rows="4" value={report[k]??''} onChange={e=>onChange(k,e.target.value)}/>
          : <input value={report[k]??''} onChange={e=>onChange(k,e.target.value)}/>} 
      </label>)}
    </div>
  </section>
}

export default function App(){
  const [parsed,setParsed]=useState([]);
  const [manual,setManual]=useState({});
  const [busy,setBusy]=useState(false);
  const [errors,setErrors]=useState([]);
  const report=useMemo(()=>({...aggregate(parsed),...manual}),[parsed,manual]);
  const stats=reportStats(report);

  async function handleFiles(files){
    setBusy(true); setErrors([]);
    const next=[]; const errs=[];
    for(const file of [...files]){
      try{ next.push(await parseFile(file)); }
      catch(e){ errs.push(`${file.name}: ${e.message}`); }
    }
    setParsed(next); setManual({}); setErrors(errs); setBusy(false);
  }

  const setField=(k,v)=>setManual(m=>({...m,[k]:v}));
  const downloadJson=()=>{
    const blob=new Blob([JSON.stringify(report,null,2)],{type:'application/json'});
    const a=document.createElement('a'); a.href=URL.createObjectURL(blob); a.download='hasil-ekstraksi-laporan.json'; a.click(); URL.revokeObjectURL(a.href);
  };

  return <main>
    <header>
      <div><p className="eyebrow">AYUB LAPORAN</p><h1>Generator Laporan Penyelenggaraan Pelatihan</h1><p>Upload semua dokumen sekaligus. Sistem mendeteksi jenis dokumen, membaca Excel dengan header dinamis, lalu menggabungkan dan menghitung field laporan.</p></div>
      <div className="score"><strong>{stats.percent}%</strong><span>{stats.filled}/{stats.total} field terisi</span></div>
    </header>

    <section className="upload card">
      <input id="files" type="file" multiple accept=".pdf,.xlsx,.xls" onChange={e=>handleFiles(e.target.files)}/>
      <label htmlFor="files"><b>{busy?'Sedang membaca file…':'Pilih / tarik paket dokumen'}</b><span>PDF, XLSX atau XLS — boleh banyak sekaligus</span></label>
      {errors.length>0 && <div className="error">{errors.map(e=><div key={e}>{e}</div>)}</div>}
    </section>

    {parsed.length>0 && <section className="card">
      <div className="section-title"><h2>Dokumen terdeteksi</h2><button onClick={downloadJson}>Unduh JSON</button></div>
      <div className="docs">{parsed.map((d,i)=><div className="doc" key={i}><b>{DOC_NAMES[d.type]||d.type}</b><span>{d.source}</span><small>{Object.values(d.report||{}).filter(v=>v!==''&&v!=null).length} field ditemukan</small></div>)}</div>
    </section>}

    <Section title="Identitas Pelatihan" keys={['nama_pelatihan','akt','PIC','metode','tgl_mulai','tgl_selesai','tgl_pengumuman','no_pengumuman','unit_asal','tempat_pelatihan','rencana_peserta','realisasi_peserta']} report={report} onChange={setField}/>
    <Section title="Profil Peserta" keys={['realisasi_laki','realisasi_perempuan','S3','S2','D4_S1','D3','D1','SMA_sederajat','Unknown_School']} report={report} onChange={setField}/>
    <Section title="Pretest & Posttest" keys={['Jenis_Ujian','peserta_ujian','Min_Kompre_Pretest','Max_Kompre_Pretest','Rata2_Kompre_Pretest','Min_NA_Posttest','Max_NA_Posttest','Rata2_NA_Posttest','Selisih_Pretest_Posttest','predikat_A','predikat_B','predikat_C','lulus_min_B','persen_kelulusan_minB']} report={report} onChange={setField}/>
    <Section title="Evaluasi" keys={['indeks_rata2_evagara_PJJ','indeks_rata2_evajar','rec_penyelenggaraan_pelatihan']} report={report} onChange={setField}/>
    <Section title="Program" keys={['nama_pelatihan_lengkap','maksud_program','diharapkan_mampu','Keterangan']} report={report} onChange={setField}/>

    <details className="card"><summary>Semua field output</summary><div className="raw">{Object.entries(report).map(([k,v])=><div key={k}><code>{k}</code><span>{String(v??'')}</span></div>)}</div></details>
  </main>
}
