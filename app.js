'use strict';
(function(){
const KEY = 'officelab_offline_v2_1';
const OLD_KEY = 'officelab_offline_v2';
const OLDEST_KEY = 'officelab_offline_v1';
const PASS_SCORE = 80;
let storageEnabled=true,storageWarned=false;
const $ = (sel, root=document) => root.querySelector(sel);
const $$ = (sel, root=document) => [...root.querySelectorAll(sel)];
const esc = value => String(value == null ? '' : value).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const own = (o,k) => Object.prototype.hasOwnProperty.call(o,k);
const clone = v => JSON.parse(JSON.stringify(v));
const fmt = num => Number(num).toLocaleString('id-ID',{maximumFractionDigits:3});
const initialState=()=>({theme:'light',learnerName:'',role:'keuangan',excelCaseId:'keuangan',caseChecks:Object.fromEntries(CAMPUS_CASES.map(x=>[x.id,[]])),course:'excel',section:'home',completedLessons:{excel:[],word:[],powerpoint:[]},excelCells:clone(EXCEL_INITIAL),excelDone:[],wordHTML:WORD_INITIAL,wordDone:[],pptSlides:clone(PPT_INITIAL),pptTheme:'blue',pptDone:[],pptPresented:false,quizBest:{excel:0,word:0,powerpoint:0},quizLast:{excel:null,word:null,powerpoint:null},passedChecks:{excel:[],word:[],powerpoint:[]},selectedLevel:{excel:'Basic',word:'Basic',powerpoint:'Basic'},levelBest:{excel:{},word:{},powerpoint:{}},levelLast:{excel:{},word:{},powerpoint:{}},levelAttempts:{excel:{},word:{},powerpoint:{}}});
function prepareState(source,fromV1=false){
  const d=initialState();
  if(!source||typeof source!=='object')return d;
  for(const k of Object.keys(d))if(own(source,k))d[k]=source[k];
  if(!OFFICE[d.course])d.course='excel';
  if(!COURSE_SECTIONS.some(s=>s.id===d.section&&(!s.only||s.only===d.course)))d.section='home';
  if(!['light','dark'].includes(d.theme))d.theme='light';
  if(!CASE_BY_ID[d.role])d.role='keuangan';
  if(!EXCEL_CASES[d.excelCaseId])d.excelCaseId='keuangan';
  if(!d.caseChecks||typeof d.caseChecks!=='object')d.caseChecks={};
  for(const c of CAMPUS_CASES){const a=Array.isArray(d.caseChecks[c.id])?d.caseChecks[c.id]:[];d.caseChecks[c.id]=[...new Set(a.filter(x=>c.tasks.some(t=>t.id===x)))];}
  d.learnerName=typeof d.learnerName==='string'?d.learnerName.slice(0,80):'';
  for(const id of Object.keys(OFFICE)){
    const indices=Array.isArray(d.completedLessons?.[id])?d.completedLessons[id]:[];
    d.completedLessons[id]=[...new Set(indices.map(i=>fromV1?V1_INDEX_TO_V2[id]?.[i]:i).filter(i=>Number.isInteger(i)&&i>=0&&i<OFFICE[id].lessons.length))];
    if(!d.passedChecks||typeof d.passedChecks!=='object')d.passedChecks={};
    const checks=Array.isArray(d.passedChecks[id])?d.passedChecks[id]:[];
    d.passedChecks[id]=[...new Set(checks.filter(i=>Number.isInteger(i)&&i>=0&&i<OFFICE[id].lessons.length))];
    if(!d.levelBest||typeof d.levelBest!=='object')d.levelBest={};
    if(!d.levelLast||typeof d.levelLast!=='object')d.levelLast={};
    if(!d.levelAttempts||typeof d.levelAttempts!=='object')d.levelAttempts={};
    for(const field of ['levelBest','levelLast','levelAttempts']){
      if(!d[field][id]||typeof d[field][id]!=='object'||Array.isArray(d[field][id]))d[field][id]={};
      for(const level of LEVELS){const v=Number(d[field][id][level.id]||0);d[field][id][level.id]=Number.isFinite(v)?Math.max(0,Math.min(field==='levelAttempts'?99999:100,v)):0;}
    }
    if(!d.selectedLevel||typeof d.selectedLevel!=='object')d.selectedLevel={};
    if(!LEVELS.some(l=>l.id===d.selectedLevel[id]))d.selectedLevel[id]='Basic';
    if(!d.quizBest||typeof d.quizBest!=='object')d.quizBest={};
    if(!d.quizLast||typeof d.quizLast!=='object')d.quizLast={};
    if(!Number.isFinite(d.quizBest[id]))d.quizBest[id]=0;
    if(!Number.isFinite(d.quizLast[id]))d.quizLast[id]=null;
  }
  if(!d.excelCells||typeof d.excelCells!=='object'||Array.isArray(d.excelCells))d.excelCells=clone(EXCEL_INITIAL);
  else d.excelCells=Object.fromEntries(Object.entries(d.excelCells).filter(([key,val])=>/^[A-G][1-9]$/.test(key)&&['string','number'].includes(typeof val)).map(([key,val])=>[key,String(val).slice(0,900)]));
  for(const key of ['excelDone','wordDone','pptDone'])if(!Array.isArray(d[key]))d[key]=[];
  if(typeof d.wordHTML!=='string')d.wordHTML=WORD_INITIAL;
  d.wordHTML=d.wordHTML.slice(0,250000);
  if(!Array.isArray(d.pptSlides)||!d.pptSlides.length)d.pptSlides=clone(PPT_INITIAL);
  d.pptSlides=d.pptSlides.slice(0,100).map(slide=>({title:String(slide.title||'').slice(0,180),body:String(slide.body||'').slice(0,10000),notes:String(slide.notes||'').slice(0,10000),layout:['content','cover','split','quote'].includes(slide.layout)?slide.layout:'content'}));
  if(!['blue','purple','green','dark','sunset'].includes(d.pptTheme))d.pptTheme='blue';
  if(fromV1){
    d.completedLessons={excel:[],word:[],powerpoint:[]};
    d.passedChecks={excel:[],word:[],powerpoint:[]};
    d.selectedLevel={excel:'Basic',word:'Basic',powerpoint:'Basic'};
    d.levelBest={excel:{},word:{},powerpoint:{}};
    d.levelLast={excel:{},word:{},powerpoint:{}};
    d.levelAttempts={excel:{},word:{},powerpoint:{}};
    d.quizBest={excel:0,word:0,powerpoint:0};
    d.quizLast={excel:null,word:null,powerpoint:null};
    d.excelCells=clone(EXCEL_INITIAL);d.excelCaseId='keuangan';d.excelDone=[];
    d.wordDone=[];d.pptDone=[];d.pptPresented=false;
  }
  return d;
}
function loadState(){
  try{
    const current=localStorage.getItem(KEY);
    if(current)return prepareState(JSON.parse(current),false);
    const v2=localStorage.getItem(OLD_KEY);
    const v1=localStorage.getItem(OLDEST_KEY);
    if(v2||v1){const migrated=prepareState(JSON.parse(v2||v1),true);localStorage.setItem(KEY,JSON.stringify(migrated));return migrated;}
  }catch(_){storageEnabled=false;}
  return initialState();
}
let state=loadState();
const ui={lessonIndex:{excel:0,word:0,powerpoint:0},lessonChoice:null,lessonFeedback:null,shortcutSearch:'',shortcutCategory:'Semua',formulaSearch:'',formulaCategory:'Semua',flashIndex:0,flashFlip:false,activeCell:'E8',activeTask:0,practiceMessage:null,pptIndex:0,presentationIndex:0,quiz:null,wordRange:null,toastTimer:null};
function save(){if(!storageEnabled){if(!storageWarned){storageWarned=true;toast('Penyimpanan otomatis tidak tersedia di browser ini. Ekspor progres untuk membuat cadangan.');}return;}try{localStorage.setItem(KEY,JSON.stringify(state));}catch(_){storageEnabled=false;if(!storageWarned){storageWarned=true;toast('Penyimpanan otomatis tidak tersedia. Ekspor progres untuk membuat cadangan.');}}}
function toast(message){const el=$('#toast');el.textContent=message;el.classList.add('visible');clearTimeout(ui.toastTimer);ui.toastTimer=setTimeout(()=>el.classList.remove('visible'),3200);}
function go(course=state.course,section='home'){
  if(!OFFICE[course])return;
  if(!COURSE_SECTIONS.some(s=>s.id===section&&(!s.only||s.only===course)))section='home';
  if(course!==state.course){ui.quiz=null;ui.shortcutSearch='';ui.formulaSearch='';ui.shortcutCategory='Semua';ui.formulaCategory='Semua';ui.flashIndex=0;ui.flashFlip=false;}
  state.course=course;state.section=section;ui.practiceMessage=null;save();render();window.scrollTo({top:0,behavior:'instant'});
}
function lessonsCompleted(id){return state.completedLessons[id].filter(i=>Number.isInteger(i)&&i>=0&&i<OFFICE[id].lessons.length).length;}
function practiceItems(id){return id==='excel'?EXCEL_TASKS:id==='word'?WORD_TASKS:PPT_TASKS;}
function practiceDone(id){return id==='excel'?state.excelDone:id==='word'?state.wordDone:state.pptDone;}
function levelIndices(id,level){return OFFICE[id].lessons.map((l,i)=>({l,i})).filter(({l})=>l.level===level).map(({i})=>i);}
function levelStats(id,level){const indices=levelIndices(id,level),read=indices.filter(i=>state.completedLessons[id].includes(i)).length,checks=indices.filter(i=>state.passedChecks[id].includes(i)).length,best=state.levelBest[id][level]||0;return {read,checks,best,total:indices.length,passed:best>=PASS_SCORE,complete:read===5&&checks===5&&best>=PASS_SCORE,percent:Math.round(100*(read+checks+(best>=PASS_SCORE?1:0))/11)};}
function progressFor(id){const total=OFFICE[id].lessons.length*2+practiceItems(id).length+LEVELS.length+1;
  const exams=LEVELS.filter(l=>(state.levelBest[id][l.id]||0)>=PASS_SCORE).length;
  const done=lessonsCompleted(id)+state.passedChecks[id].length+practiceDone(id).length+exams+(state.quizBest[id]>=PASS_SCORE?1:0);
  return Math.min(100,Math.round(100*done/total));}
function overallProgress(){return Math.round(Object.keys(OFFICE).reduce((sum,id)=>sum+progressFor(id),0)/3);}
function chooseLevel(course,level,section='lessons'){
  if(!OFFICE[course]||!LEVELS.some(l=>l.id===level))return;
  state.selectedLevel[course]=level;
  const items=levelIndices(course,level);
  ui.lessonIndex[course]=items.find(i=>!state.completedLessons[course].includes(i))??items[0];
  ui.lessonChoice=null;ui.lessonFeedback=null;
  ui.quiz=null;save();go(course,section);
}
function headerBlock(eyebrow,title,subtitle,actions=''){
  return `<div class="module-title"><div><span class="eyebrow">${esc(eyebrow)}</span><h1 class="page-title">${esc(title)}</h1><p class="sub-title">${esc(subtitle)}</p></div>${actions?`<div class="inline-actions">${actions}</div>`:''}</div>`;
}
function renderNavigation(){
  $('#course-nav').innerHTML=Object.entries(OFFICE).map(([id,c])=>`<button type="button" class="course-link ${state.course===id?'active':''}" data-action="navigate" data-course="${id}" data-section="home" ${state.course===id?'aria-current="page"':''}><span class="app-icon ${id}">${c.short}</span><span class="label-stack">${c.name}<em>${id==='excel'?'Spreadsheet':id==='word'?'Dokumen':'Presentasi'}</em></span></button>`).join('');
  $('#section-nav').innerHTML=COURSE_SECTIONS.filter(s=>!s.only||s.only===state.course).map(s=>`<button type="button" class="section-link ${s.id===state.section?'active':''}" data-action="navigate" data-course="${state.course}" data-section="${s.id}" ${state.section===s.id?'aria-current="page"':''}><span class="section-symbol" aria-hidden="true">${s.icon}</span><span>${s.label}</span></button>`).join('');
  const o=overallProgress();$('#overall-percent').textContent=o+'%';$('#overall-bar').style.width=o+'%';$('.side-progress small').textContent=storageEnabled?'Progres tersimpan otomatis di browser ini.':'Penyimpanan dibatasi. Ekspor progres sebagai cadangan.';
  $('#breadcrumb-course').textContent=OFFICE[state.course].name;
  $('#breadcrumb-section').textContent=(COURSE_SECTIONS.find(s=>s.id===state.section)?.label||'Beranda')+(state.section==='lessons'?' · '+state.selectedLevel[state.course]:'');
  document.body.dataset.course=state.course;document.body.dataset.theme=state.theme;
  $('#theme-toggle').textContent=state.theme==='light'?'☾':'☀';
  $('#theme-toggle').title=state.theme==='light'?'Aktifkan mode gelap':'Aktifkan mode terang';
}
function render(){renderNavigation();
  const view=$('#main-view');
  try{
    view.innerHTML=({home:renderHome,lessons:renderLessons,cases:renderCases,shortcuts:renderShortcuts,formulas:renderFormulas,practice:renderPractice,quiz:renderQuiz,settings:renderSettings,about:renderAbout}[state.section]||renderHome)();
    if(state.section==='shortcuts')updateShortcutRows();
    if(state.section==='formulas')updateFormulaCards();
    if(state.section==='practice'&&state.course==='word')refreshWordProgress();
    if(state.section==='practice'&&state.course==='powerpoint')refreshPptGoals();
  }catch(err){console.error(err);view.innerHTML=`<div class="panel"><h2>Terjadi kendala menampilkan materi.</h2><p>${esc(err.message||'Silakan muat ulang aplikasi.')}</p><button class="primary-btn" data-action="navigate" data-section="home">Ke Beranda</button></div>`;}
}
function renderAbout(){
  const roles=CAMPUS_CASES.map(c=>`<span class="about-role">${esc(c.name)}</span>`).join('');
  const levelLabels=['Basic','Beginner','Intermediate','Advanced','Expert'];
  return `<div class="about-page">
    <div class="about-hero">
      <div class="about-hero-copy">
        <span class="about-kicker"><span class="about-kicker-dot"></span> MENGENAL OFFICELAB · V2.1.2</span>
        <h1>Lebih terampil mengelola<br> <em>pekerjaan administrasi kampus.</em></h1>
        <p>OfficeLab adalah ruang belajar interaktif untuk mengembangkan keterampilan Microsoft Excel, Word, dan PowerPoint melalui materi bertahap, latihan langsung, dan studi kasus yang dekat dengan kegiatan administrasi perguruan tinggi.</p>
        <div class="about-hero-actions">
          <button class="about-white-button" type="button" data-action="navigate" data-section="lessons">Jelajahi materi <span aria-hidden="true">↗</span></button>
          <button class="about-ghost-button" type="button" data-action="navigate" data-section="cases">Lihat studi kasus unit →</button>
        </div>
      </div>
      <div class="about-hero-symbol" aria-hidden="true"><span class="about-symbol-grid"><b> X </b><b> W </b><b> P </b><b> ✦ </b></span><span class="about-symbol-caption">LEARN · PRACTICE · IMPROVE</span></div>
    </div>
    <div class="about-facts" aria-label="Ringkasan OfficeLab">
      <div><strong>03</strong><span>Aplikasi Office</span></div>
      <div><strong>05</strong><span>Level kompetensi</span></div>
      <div><strong>75</strong><span>Modul pembelajaran</span></div>
      <div><strong>07</strong><span>Jalur administrasi</span></div>
      <div><strong>100%</strong><span>Dapat digunakan offline</span></div>
    </div>

    <div class="about-section-head"><div><span class="about-overline">01 / LATAR BELAKANG</span><h2>Dirancang untuk kebutuhan nyata di kampus</h2><p>Pembelajaran berangkat dari pekerjaan yang ditemui staf administratif setiap hari, bukan semata-mata hafalan perintah aplikasi.</p></div></div>
    <div class="about-feature-grid">
      <article class="about-feature"><span class="about-feature-icon green">▦</span><h3>Belajar sesuai pekerjaan</h3><p>Contoh dan latihan mencakup rekap anggaran, administrasi mahasiswa, surat dinas, pengelolaan arsip, pelaporan program, dan presentasi kinerja.</p></article>
      <article class="about-feature"><span class="about-feature-icon blue">◫</span><h3>Dari konsep ke praktik</h3><p>Materi dikembangkan melalui tujuan belajar, contoh, langkah praktikum, tugas, latihan mikro, kuis, serta pemeriksaan hasil berbasis rubrik.</p></article>
      <article class="about-feature"><span class="about-feature-icon orange">⌁</span><h3>Fleksibel dan mandiri</h3><p>Dapat dibuka di browser tanpa instalasi tambahan, tanpa login, dan tanpa koneksi internet. Progres belajar disimpan di browser perangkat yang digunakan.</p></article>
    </div>

    <div class="about-grid-split">
      <section class="about-card about-howto" aria-labelledby="about-howto-heading">
        <span class="about-overline">02 / CARA MENGGUNAKAN</span><h2 id="about-howto-heading">Mulai dalam lima langkah</h2>
        <ol class="about-steps">
          <li><span>01</span><div><strong>Pilih aplikasi</strong><p>Masuk ke tab Excel, Word, atau PowerPoint sesuai kebutuhan pekerjaan Anda.</p></div></li>
          <li><span>02</span><div><strong>Tentukan tingkat kemampuan</strong><p>Pelajari Basic, Beginner, Intermediate, Advanced, hingga Expert secara bertahap; gunakan tes penempatan bila diperlukan.</p></div></li>
          <li><span>03</span><div><strong>Pilih jalur unit kerja</strong><p>Gunakan studi kasus berdasarkan bidang administrasi untuk membuka data dan tugas yang relevan.</p></div></li>
          <li><span>04</span><div><strong>Praktik dan evaluasi</strong><p>Kerjakan simulator, latihan mikro, serta tes per level. Untuk fitur lanjutan, lanjutkan praktik pada Microsoft Office asli.</p></div></li>
          <li><span>05</span><div><strong>Periksa dan simpan progres</strong><p>Pantau hasil latihan, buat cadangan JSON melalui Pengaturan, dan diskusikan proyek akhir bersama fasilitator.</p></div></li>
        </ol>
        <div class="about-inline-buttons"><button class="primary-btn" type="button" data-action="navigate" data-section="lessons">Buka materi →</button><button class="outline-btn" type="button" data-action="navigate" data-section="settings">Pengaturan & progres</button></div>
      </section>
      <section class="about-card about-purpose" aria-labelledby="about-purpose-heading">
        <span class="about-overline">03 / TUJUAN PEMBELAJARAN</span><h2 id="about-purpose-heading">Apa yang ingin dicapai?</h2>
        <div class="about-objectives">
          <p><span>✓</span> Meningkatkan ketelitian pengolahan data dan penyusunan laporan administratif.</p>
          <p><span>✓</span> Mempercepat pekerjaan rutin melalui shortcut, fungsi, template, dan teknik produktivitas Office.</p>
          <p><span>✓</span> Menyeragamkan kualitas dokumen, pelaporan, dan komunikasi kerja lintas unit.</p>
          <p><span>✓</span> Mengembangkan keterampilan pemecahan masalah dengan tugas yang dapat diperiksa hasilnya.</p>
          <p><span>✓</span> Mendukung pembelajaran mandiri maupun pelatihan yang dipandu fasilitator.</p>
        </div>
        <div class="about-learning-path"><span>JALUR KOMPETENSI</span><div>${levelLabels.map((l,i)=>`<span class="about-path-step"><b>${i+1}</b>${l}</span>`).join('')}</div></div>
      </section>
    </div>

    <section class="about-roles-card" aria-labelledby="about-roles-heading"><div><span class="about-overline">04 / PENGGUNA</span><h2 id="about-roles-heading">Untuk staf di berbagai unit</h2><p>Studi kasus disiapkan untuk mendukung tujuh bidang administrasi di lingkungan perguruan tinggi.</p></div><div class="about-role-list">${roles}</div></section>

    <section class="about-developer" aria-labelledby="about-developer-heading">
      <div class="about-developer-photo"><img src="assets/developer-profile.png" alt="Foto profil pengembang OfficeLab" loading="lazy" width="340" height="340"><span class="about-photo-caption">PENGEMBANG OFFICELAB</span></div>
      <div class="about-developer-details"><span class="about-overline">05 / DI BALIK OFFICELAB</span><h2 id="about-developer-heading">Tentang Pengembang</h2><p class="about-developer-intro">OfficeLab dikembangkan sebagai media pembelajaran digital yang membantu tenaga kependidikan mempelajari perangkat produktivitas secara kontekstual dan menerapkannya pada pekerjaan administrasi perguruan tinggi.</p>
      <div class="about-person-name">Mahardika I. Takaendengan, SST., M.Si</div>
      <div class="about-person-affiliation">Program Studi Sistem Informasi - FMIPA UNSRAT</div>
      <a class="about-person-mail" href="mailto:mahardika@unsrat.ac.id" aria-label="Kirim email kepada Mahardika I. Takaendengan"><span aria-hidden="true">✉</span> mahardika@unsrat.ac.id <span aria-hidden="true">↗</span></a>
      <div class="about-dev-signature"><span>OFFICELAB</span><small>Praktik yang relevan. Keterampilan yang berkembang.</small></div>
      </div>
    </section>

    <div class="about-notes"><div><strong>ⓘ Tentang penggunaan dan privasi</strong><p>OfficeLab adalah simulator edukasi, bukan produk resmi Microsoft atau aplikasi operasional kampus. Contoh dataset bersifat fiktif. Progres disimpan secara lokal di browser; lakukan ekspor cadangan sebelum mengganti perangkat atau membersihkan data browser. Hindari memasukkan data pribadi dan dokumen dinas yang bersifat rahasia ke dalam latihan.</p></div><span class="about-version-pill">OfficeLab v2.1.2</span></div>
  </div>`;
}

function renderHome(){const id=state.course,c=OFFICE[id],pct=progressFor(id),n=lessonsCompleted(id),p=practiceDone(id).length;
  const visual=id==='excel'?`<div class="art-window"><div class="dots"><span></span><span></span><span></span></div><div class="art-grid">${Array.from({length:20},()=>'<span></span>').join('')}</div></div>`:id==='word'?`<div class="art-window" style="transform:rotate(5deg);padding:22px"><div style="height:10px;width:55%;background:#8cb6f1;border-radius:6px;margin-bottom:17px"></div>${[90,95,70,94,85,56,90].map((x,i)=>`<div style="height:6px;width:${x}%;background:${i===3?'#b5cef1':'#dee8f5'};border-radius:5px;margin:9px 0"></div>`).join('')}</div>`:`<div class="art-window" style="padding:13px"><div style="aspect-ratio:16/9;background:linear-gradient(125deg,#efb08c,#e87d47);border-radius:7px;padding:22px 14px;color:#fff"><div style="font-size:20px;font-weight:800;line-height:1.2">Your next<br>big idea.</div><div style="width:65px;height:4px;background:#fff8;border-radius:4px;margin-top:12px"></div></div></div>`;
  const cards=LEVELS.map((lev,i)=>{const st=levelStats(id,lev.id);return `<button class="level-card ${state.selectedLevel[id]===lev.id?'current':''}" data-action="level-select" data-level="${lev.id}" type="button"><span class="level-card-top"><span class="level-number" style="--level-color:${lev.color}">${lev.label}</span><span class="level-status ${st.complete?'complete':''}">${st.complete?'✓ Tuntas':`${st.percent}% selesai`}</span></span><strong>${lev.name}</strong><span class="level-desc">${lev.caption}</span><span class="level-meter"><span style="width:${st.percent}%;background:${lev.color}"></span></span><span class="level-card-bottom">${st.read}/5 materi · Tes ${st.best?st.best+'%':'belum'}</span></button>`;}).join('');
  return `<div class="hero"><div class="hero-left"><span class="eyebrow inverse">✦ OFFICELAB 2.1.2 · ${c.name.toUpperCase()}</span><h1>${id==='excel'?'Rekap administrasi, lebih akurat.':id==='word'?'Persuratan dan laporan lebih tertib.':'Paparan kinerja lebih terarah.'}</h1><p>${esc(c.description)} Latihan berbasis pekerjaan administrasi kampus: keuangan, mahasiswa, perencanaan, persuratan, kearsipan, dan unit lainnya. Pilih lima level serta studi kasus unit, semuanya offline.</p><div class="hero-actions"><button class="primary-btn" data-action="navigate" data-section="lessons">Mulai / lanjutkan belajar →</button><button class="outline-btn" data-action="navigate" data-section="practice">Buka simulator ◇</button></div></div><div class="hero-art" aria-hidden="true">${visual}</div></div>
  <div class="stat-grid"><div class="stat-card"><div><small>Materi selesai</small><strong>${n}<span style="color:var(--muted);font-size:15px"> / 25</span></strong></div><span class="stat-icon">▤</span></div><div class="stat-card"><div><small>Latihan singkat lolos</small><strong>${state.passedChecks[id].length}<span style="color:var(--muted);font-size:15px"> / 25</span></strong></div><span class="stat-icon">☑</span></div><div class="stat-card"><div><small>Progres ${c.name}</small><strong>${pct}<span style="color:var(--muted);font-size:15px">%</span></strong></div><span class="stat-icon">◈</span></div></div>
  <div class="section-heading"><div><h2>Jalur belajar 5 level</h2><p>Setiap level berisi 5 modul, 5 latihan cek pemahaman, dan tes akhir level. Pilih level untuk memulai.</p></div><button class="text-link" data-action="navigate" data-section="quiz">Lihat evaluasi →</button></div>
  <div class="level-grid">${cards}</div>
  <div class="practice-intro" style="margin-top:17px"><b>Jalur aktif:</b> ${esc(CASE_BY_ID[state.role]?.name||'Keuangan')} · ${esc(CASE_BY_ID[state.role]?.tagline||'Praktik unit')}. <button class="text-link" data-action="navigate" data-section="cases">Buka studi kasus unit →</button></div>
  <div class="mini-grid-2" style="margin-top:18px"><div class="mini-guide"><strong>⚙ Praktik langsung</strong><p>Simulasi worksheet Excel, editor Word, dan slide PowerPoint tetap tersedia.</p><button class="text-link" style="padding:10px 0 0" data-action="navigate" data-section="practice">Masuk ke laboratorium →</button></div><div class="mini-guide"><strong>⌨ Referensi produktivitas</strong><p>Shortcut interaktif, kartu hafalan, serta pustaka rumus Excel dapat dibuka kapan saja.</p><button class="text-link" style="padding:10px 0 0" data-action="navigate" data-section="shortcuts">Buka shortcut →</button></div></div>
  <div class="pill-note"><b>Catatan:</b> Level Advanced dan Expert mencakup materi yang membutuhkan Microsoft Office asli. OfficeLab menyediakan pembelajaran dan evaluasi konsep, bukan emulasi lengkap seluruh fitur Office.</div>`;
}
function renderLessons(){
  const id=state.course,c=OFFICE[id],level=state.selectedLevel[id]||'Basic',indices=levelIndices(id,level);
  if(!indices.includes(ui.lessonIndex[id]))ui.lessonIndex[id]=indices[0];
  const i=ui.lessonIndex[id],l=c.lessons[i],done=state.completedLessons[id].includes(i),passed=state.passedChecks[id].includes(i),st=levelStats(id,level),check=l.check;
  const levTabs=LEVELS.map(lev=>{const x=levelStats(id,lev.id);return `<button class="level-tab ${lev.id===level?'active':''}" type="button" data-action="level-select" data-level="${lev.id}"><span class="dot" style="background:${lev.color}"></span><span>${lev.name}</span><small>${x.complete?'✓':x.read+'/5'}</small></button>`;}).join('');
  const choices=check.options.map((opt,j)=>{let cls=ui.lessonChoice===j?'selected':'';if(ui.lessonFeedback?.index===i&&ui.lessonFeedback.correct&&j===check.answer)cls+=' correct';if(ui.lessonFeedback?.index===i&&!ui.lessonFeedback.correct&&ui.lessonChoice===j)cls+=' incorrect';return `<button class="micro-option ${cls}" data-action="lesson-answer" data-index="${j}" type="button"><span>${'ABCD'[j]}</span>${esc(opt)}</button>`;}).join('');
  return `${headerBlock('KURIKULUM V2.1','Materi '+c.name,`25 modul administrasi kampus dalam 5 level, lengkap dengan prosedur, contoh, tugas, rubrik penilaian dan kuis.`, `<span class="chip">${lessonsCompleted(id)} / 25 materi selesai</span>`)}
  <nav class="level-tabs" aria-label="Level pembelajaran">${levTabs}</nav>
  <div class="level-overview"><div><span class="eyebrow">LEVEL ${LEVELS.findIndex(x=>x.id===level)+1} / 5</span><h2>${level} <span>— ${LEVELS.find(x=>x.id===level).caption}</span></h2><p>${LEVELS.find(x=>x.id===level).description}</p></div><div class="level-summary"><strong>${st.percent}%</strong><span>${st.read}/5 modul · ${st.checks}/5 latihan<br>Tes akhir: ${st.best?st.best+'%':'belum dikerjakan'}</span><button class="outline-btn" data-action="level-exam" data-level="${level}">Tes level ${level} →</button>${st.complete?`<button class="primary-btn" data-action="certificate-download" data-level="${level}">↓ Bukti tuntas</button>`:''}</div></div>
  <div class="two-column"><div class="panel no-pad"><div class="lesson-list">${indices.map((j,n)=>{const x=c.lessons[j];return `<button class="lesson-item ${i===j?'active':''}" data-action="lesson-select" data-index="${j}"><span class="lesson-item-index">${state.completedLessons[id].includes(j)?'✓':n+1}</span><span><strong>${esc(x.title)}</strong><small>${x.min} menit · ${state.passedChecks[id].includes(j)?'✓ Tes mikro lulus':'Latihan tersedia'}</small></span></button>`;}).join('')}</div><div class="lesson-list-help">✓ Tandai modul selesai setelah dipelajari.<br>○ Jawab latihan mikro untuk menguji konsep.</div></div>
  <article class="panel lesson-content"><span class="level-pill">${esc(l.level)} · Modul ${indices.indexOf(i)+1}/5 · ${l.min} menit</span><h2>${esc(l.title)}</h2><p>${esc(l.lead)}</p><div class="learning-box"><b>Tujuan pembelajaran</b><p>${esc(l.objective)}</p></div><h3>Materi inti dan prinsip kerja</h3><ul>${l.points.map(t=>`<li>${esc(t)}</li>`).join('')}</ul><h3>Langkah praktikum terstruktur</h3><ol class="lesson-steps">${(l.steps||[]).map((t,j)=>`<li><strong>Langkah ${j+1}.</strong> ${esc(t)}</li>`).join('')}</ol><h3>Contoh berbasis administrasi kampus</h3><div class="code-example">${esc(l.example)}</div><h3>Tugas mandiri / kelompok</h3><p>${esc(l.exercise)}</p><div class="learning-box"><b>Output yang dikumpulkan</b><p>${esc(l.output||'Bukti latihan yang dapat diperiksa.')}</p></div><h3>Rubrik penilaian hasil praktik</h3><ul>${(l.rubric||[]).map(t=>`<li>${esc(t)}</li>`).join('')}</ul><h3>Catatan risiko</h3><ul>${(l.cautions||[]).map(t=>`<li>${esc(t)}</li>`).join('')}</ul><div class="lesson-role">Relevan untuk: ${(l.roles||[]).map(r=>`<span>${esc(CASE_BY_ID[r]?.name||'Lintas unit')}</span>`).join('')}</div>${l.requiresOffice?`<div class="office-required">ⓘ <b>Praktik di Office asli.</b> Praktik fitur ini perlu dilakukan di Microsoft Office desktop yang mendukungnya. Cek konsepnya melalui tes di bawah.</div>`:''}<div class="learning-box"><b>💡 Tips</b><p>${esc(l.tip)}</p></div>
  <div class="micro-challenge"><div class="micro-head"><div><span class="eyebrow">LATIHAN MIKRO · INTERAKTIF</span><h3>Cek pemahaman</h3></div><span class="${passed?'mini-pass':'mini-pending'}">${passed?'✓ Pernah lulus':'Belum lulus'}</span></div><p class="micro-question">${esc(check.question)}</p><div class="micro-options">${choices}</div><div class="micro-actions"><button class="primary-btn" data-action="lesson-check" ${ui.lessonChoice==null?'disabled':''}>Periksa jawaban</button><button class="outline-btn" data-action="navigate" data-section="practice">Buka simulator ↗</button></div>${ui.lessonFeedback?.index===i?`<div class="micro-feedback ${ui.lessonFeedback.correct?'ok':'bad'}"><b>${ui.lessonFeedback.correct?'✓ Jawaban benar':'Coba lagi'}</b> · ${esc(check.why)}</div>`:''}</div>
  <div class="lesson-foot"><button class="outline-btn" data-action="lesson-select" data-index="${i-1}" ${i===indices[0]?'disabled':''}>← Sebelumnya</button><button class="${done?'soft-btn':'primary-btn'}" data-action="lesson-complete" data-index="${i}">${done?'✓ Sudah selesai':'✓ Tandai selesai'}</button><button class="outline-btn" data-action="lesson-select" data-index="${i+1}" ${i===indices[indices.length-1]?'disabled':''}>Berikutnya →</button></div></article></div>`;
}
function renderCases(){
 const c=CASE_BY_ID[state.role]||CAMPUS_CASES[0],done=state.caseChecks[c.id]||[],pct=Math.round(done.length/c.tasks.length*100),summary=EXCEL_CASES[c.id];
 const roles=CAMPUS_CASES.map(unit=>`<button type="button" class="case-role ${unit.id===c.id?'active':''}" data-action="case-select" data-case="${unit.id}"><span class="role-emoji">${esc(unit.emoji)}</span><strong>${esc(unit.name)}</strong><small>${esc(unit.tagline)}</small></button>`).join('');
 const stages=c.tasks.map(t=>`<label class="case-task ${done.includes(t.id)?'done':''}"><input type="checkbox" data-action="case-task" data-case="${c.id}" data-task="${t.id}" ${done.includes(t.id)?'checked':''}><span>${esc(t.text)}</span></label>`).join('');
 const rubrics=c.rubric.map(r=>`<div class="rubric-item"><span><b>${esc(r.name)}</b><small>${esc(r.desc)}</small></span><strong>${r.weight}%</strong></div>`).join('');
 return `${headerBlock('PRAKTIK BERBASIS PEKERJAAN','Studi kasus administrasi kampus','Pilih unit kerja dan selesaikan rangkaian tugas Excel → Word → PowerPoint menggunakan data contoh yang aman.')}<div class="safety-banner"><b>Data pelatihan sepenuhnya fiktif.</b> Dilarang menggunakan daftar mahasiswa, pegawai, rekening, atau dokumen rahasia asli dalam latihan tanpa izin dan pengamanan institusi. Template bukan dokumen resmi.</div>
 <div class="case-role-grid">${roles}</div>
 <div class="case-hero" style="--rolecolor:${c.theme}"><div><span class="eyebrow">JALUR KOMPETENSI UNIT · ${esc(c.unit)}</span><h2>${esc(c.name)}</h2><p>${esc(c.scenario)}</p><div class="case-kpi"><span>Total target/pagu: <b>${fmt(c.metricSummary.target)}</b></span><span>Realisasi: <b>${fmt(c.metricSummary.actual)}</b></span><span>Selisih: <b>${fmt(c.metricSummary.gap)}</b></span><small>Angka ${c.kind==='Rupiah'?'dalam rupiah':'dalam satuan '+esc(c.kind).toLowerCase()} · simulasi</small></div></div><div class="case-score"><strong>${pct}%</strong><span>${done.length}/${c.tasks.length} bukti kerja dilaporkan selesai</span><div class="track"><div class="track-fill" style="width:${pct}%"></div></div></div></div>
 <div class="case-application-grid"><article class="case-app-card"><div class="case-app-top"><span class="app-icon excel">X</span><b>1. Excel · Data dan perhitungan</b></div><p>Gunakan enam baris data contoh. Hitung total, selisih, kategori, dan indikator yang diverifikasi silang.</p><div class="case-actions"><button class="primary-btn" data-action="case-load-excel">Buka simulator Excel</button><button class="outline-btn" data-action="case-download-csv" data-case="${c.id}">↓ Dataset CSV</button></div></article>
 <article class="case-app-card"><div class="case-app-top"><span class="app-icon word">W</span><b>2. Word · Dokumen tindak lanjut</b></div><p>Muat template ${esc(WORD_TEMPLATES[c.id].title)} lalu isi dengan angka dan kesimpulan dari Excel.</p><button class="primary-btn" data-action="case-load-word">Muat template Word</button></article>
 <article class="case-app-card"><div class="case-app-top"><span class="app-icon powerpoint">P</span><b>3. PowerPoint · Paparan hasil</b></div><p>Bangun paparan ringkas dengan KPI, sumber, interpretasi, serta rencana tindakan yang dapat diperiksa.</p><button class="primary-btn" data-action="case-load-ppt">Muat 4 slide contoh</button></article></div>
 <div class="case-two-col"><section class="panel"><div class="section-heading small-heading"><div><h2>Checklist kerja peserta</h2><p>Centang hanya setelah bukti tersedia. Checklist bersifat pernyataan mandiri, bukan koreksi otomatis.</p></div></div><div class="case-checklist">${stages}</div><div class="case-actions" style="margin-top:16px"><button class="outline-btn" data-action="case-sheet">↓ Lembar kerja / penilaian</button><button class="outline-btn" data-action="case-reset">↺ Reset checklist unit</button></div></section><aside class="panel"><h3>Rubrik penilaian fasilitator · Total 100%</h3><p class="sub-title">Bobot ini adalah rancangan pelatihan, bukan standar sertifikasi formal.</p><div class="rubric-list">${rubrics}</div><div class="learning-box"><b>Prosedur pelatihan</b><p>Pelajari materi sesuai level → selesaikan kasus unit → uji silang angka dan naskah → minta umpan balik fasilitator → simpan hasil latihan.</p></div><div class="learning-box"><b>Sumber rujukan dan batasan</b><p>Gunakan pedoman tata naskah, klasifikasi arsip, jadwal retensi, dan pengamanan data yang telah disahkan di institusi. OfficeLab tidak menetapkan ketentuan resmi tersebut.</p></div></aside></div>`;
}
function caseCsv(id){const c=EXCEL_CASES[id];if(!c)return;const v=c.cells,rows=[];for(let r=1;r<=7;r++){const row=[];for(let k=1;k<=5;k++){const key=colName(k)+r;let val=v[key]||'';if(k===5&&r>1)val=Number(v['C'+r])-Number(v['D'+r]);row.push('"'+String(val).replace(/"/g,'""')+'"');}rows.push(row.join(','));}download('OfficeLab_v21_Dataset_'+id+'.csv','\ufeff'+rows.join('\r\n'),'text/csv;charset=utf-8');}
function caseSheet(){const c=CASE_BY_ID[state.role],list=state.caseChecks[c.id]||[],date=new Date().toLocaleDateString('id-ID');const html=`<!doctype html><html lang="id"><meta charset="utf-8"><title>Lembar Penilaian OfficeLab</title><style>body{font:14px/1.6 Arial,sans-serif;color:#1b293f;padding:40px;max-width:900px;margin:auto}h1{color:#166b8c}table{border-collapse:collapse;width:100%;margin:20px 0}td,th{border:1px solid #bcc7d4;padding:10px;vertical-align:top;text-align:left}small{color:#667}th{background:#f3f7f9}@media print{button{display:none}}</style><h1>OfficeLab v2.1 · Lembar Praktikum Unit</h1><p><b>Peserta:</b> ${esc(state.learnerName||'_________________')}<br><b>Jalur:</b> ${esc(c.name)}<br><b>Tanggal cetak:</b> ${esc(date)}<br><b>Keterangan:</b> Seluruh data latihan fiktif.</p><h2>Checklist bukti kerja</h2><table><tr><th>No</th><th>Bukti</th><th>Dilaporkan</th><th>Paraf/komentar fasilitator</th></tr>${c.tasks.map((t,i)=>`<tr><td>${i+1}</td><td>${esc(t.text)}</td><td>${list.includes(t.id)?'Ya':'Belum'}</td><td></td></tr>`).join('')}</table><h2>Rubrik</h2><table><tr><th>Kriteria</th><th>Bobot</th><th>Nilai (0–100)</th></tr>${c.rubric.map(r=>`<tr><td>${esc(r.name)}</td><td>${r.weight}%</td><td></td></tr>`).join('')}</table><p><small>Lembar ini alat evaluasi pelatihan internal. Status mandiri dan skor kuis OfficeLab bukan sertifikasi kompetensi atau persetujuan institusi.</small></p><button onclick="print()">Cetak / PDF</button></html>`;download('OfficeLab_v21_Lembar_Praktikum_'+c.id+'.html',html,'text/html;charset=utf-8');}

function renderShortcuts(){let c=OFFICE[state.course],cats=['Semua',...new Set(c.shortcuts.map(s=>s[2]))];return `${headerBlock('REFERENSI','Shortcut keyboard',`Pelajari ${c.shortcuts.length} pintasan ${c.name} untuk Microsoft Office desktop Windows. Coba kartu hafalan di bawah tabel.`)}
  <div class="search-filter"><input class="filter-input" type="search" id="shortcut-filter" placeholder="Cari fungsi atau kombinasi, misalnya Ctrl + S..." value="${esc(ui.shortcutSearch)}" aria-label="Cari shortcut"><select id="shortcut-category" class="filter-select" aria-label="Filter kategori">${cats.map(cat=>`<option ${ui.shortcutCategory===cat?'selected':''}>${esc(cat)}</option>`).join('')}</select></div>
  <div class="panel no-pad" style="overflow:hidden"><div style="overflow-x:auto"><table class="shortcut-table"><thead><tr><th>Fungsi</th><th>Kombinasi tombol</th><th>Kategori</th></tr></thead><tbody id="shortcut-rows"></tbody></table></div></div><div id="shortcut-count" class="progress-info" style="margin-top:9px"></div>
  <div class="flash-card"><div class="lab-header"><h3>⌨ Kartu hafalan</h3><small>Ketahui aksi, tebak tombolnya.</small></div><div id="flashcard-slot"></div></div><div class="pill-note">Shortcut di halaman ini ditujukan terutama untuk aplikasi Office desktop di Windows. Perilaku browser, versi aplikasi, keyboard, dan sistem operasi dapat berbeda.</div>`;}
function updateShortcutRows(){const c=OFFICE[state.course],q=ui.shortcutSearch.trim().toLowerCase(),filtered=c.shortcuts.filter(s=>(ui.shortcutCategory==='Semua'||s[2]===ui.shortcutCategory)&&(!q||s.join(' ').toLowerCase().includes(q)));const holder=$('#shortcut-rows');if(!holder)return;holder.innerHTML=filtered.map(s=>`<tr><td>${esc(s[1])}</td><td>${s[0].split(' + ').map(k=>`<kbd>${esc(k)}</kbd>`).join(' <span style="color:var(--muted)">+</span> ')}</td><td class="category-cell">${esc(s[2])}</td></tr>`).join('')||`<tr><td colspan="3" class="empty-state">Tidak ada shortcut yang cocok dengan pencarian.</td></tr>`;$('#shortcut-count').textContent=`Menampilkan ${filtered.length} dari ${c.shortcuts.length} shortcut.`;renderFlashcard();}
function renderFlashcard(){const c=OFFICE[state.course],slot=$('#flashcard-slot');if(!slot)return;ui.flashIndex=(ui.flashIndex+c.shortcuts.length)%c.shortcuts.length;const sc=c.shortcuts[ui.flashIndex];slot.innerHTML=`<div class="flash-inner"><small>${ui.flashFlip?'JAWABAN':'PERTANYAAN'} · ${ui.flashIndex+1} / ${c.shortcuts.length}</small><h3>${ui.flashFlip?`<span class="shortcut-preview">${sc[0].split(' + ').map(k=>`<kbd style="font-size:17px;padding:9px 12px">${esc(k)}</kbd>`).join(' + ')}</span>`:`${esc(sc[1])}`}</h3>${ui.flashFlip?`<small>${esc(sc[1])}</small>`:'<small>Ingat kombinasi tombolnya, lalu buka jawaban.</small>'}</div><div class="flash-actions"><button class="outline-btn" data-action="flash-prev">←</button><button class="primary-btn" data-action="flash-flip">${ui.flashFlip?'Tutup jawaban':'Lihat jawaban'}</button><button class="outline-btn" data-action="flash-next">→</button></div>`;}
function formulaForCase(f){
 const c=EXCEL_CASES[state.excelCaseId]||EXCEL_CASES.keuangan,criterion=c.criterion,lookup=c.lookup;
 const dict={
 SUM:['=SUM(D2:D7)','D8'],AVERAGE:['=AVERAGE(D2:D7)','C9'],
 MIN:['=MIN(D2:D7)','G6'],MAX:['=MAX(D2:D7)','G6'],COUNT:['=COUNT(D2:D7)','G7'],COUNTA:['=COUNTA(A2:A7)','G7'],
 IF:['=IF(D2>C2;"Melebihi";"Dalam batas")','F2'],AND:['=AND(C2>=100;D2>=100)','G6'],OR:['=OR(D2>C2;D3>C3)','G6'],
 COUNTIF:[`=COUNTIF(B2:B7;"${criterion}")`,'D9'],SUMIF:[`=SUMIF(B2:B7;"${criterion}";D2:D7)`,'E9'],
 COUNTIFS:[`=COUNTIFS(B2:B7;"${criterion}";D2:D7;">=100")`,'G7'],
 SUMIFS:[`=SUMIFS(D2:D7;B2:B7;"${criterion}";C2:C7;">=100")`,'G8'],
 ROUND:['=ROUND(D2/C2*100;1)','G2'],CONCAT:['=CONCAT(A3;" - ";B3)','G3'],
 LEFT:['=LEFT(A2;4)','G6'],RIGHT:['=RIGHT(A2;4)','G6'],LEN:['=LEN(A2)','G6'],
 TRIM:['=TRIM("  Contoh   Layanan  ")','G6'],
 VLOOKUP:[`=VLOOKUP("${lookup}";A2:D7;4;FALSE)`,'G5'],
 XLOOKUP:[`=XLOOKUP("${lookup}";A2:A7;D2:D7)`,'G4'],IFERROR:['=IFERROR(10/0;"Perlu periksa")','G9']
 };
 const value=dict[f.name]||[f.example,f.target];return {example:value[0],target:value[1]};
}

function renderFormulas(){if(state.course!=='excel')return renderHome();const cats=['Semua',...new Set(OFFICE.excel.formulas.map(f=>f.category))];return `${headerBlock('PUSTAKA RUMUS','Rumus & fungsi Excel','Pelajari sintaks, kegunaan, dan contoh penggunaan. Klik “Coba di lembar” untuk mengisi contoh pada simulator Excel.')}<div class="practice-intro"><b>Format penulisan:</b> Contoh memakai nama fungsi berbahasa Inggris dan pemisah titik koma (;). Excel dengan pengaturan regional tertentu memakai koma (,). Simulator OfficeLab menerima keduanya.</div>
<div class="search-filter"><input class="filter-input" type="search" id="formula-filter" placeholder="Cari SUM, IF, lookup, teks..." value="${esc(ui.formulaSearch)}" aria-label="Cari rumus"><select class="filter-select" id="formula-category" aria-label="Filter kategori fungsi">${cats.map(x=>`<option ${ui.formulaCategory===x?'selected':''}>${esc(x)}</option>`).join('')}</select></div><div class="formula-grid" id="formula-cards"></div><div id="formula-count" class="progress-info" style="margin-top:11px"></div>`;}
function updateFormulaCards(){const slot=$('#formula-cards');if(!slot)return;const q=ui.formulaSearch.trim().toLowerCase(),list=OFFICE.excel.formulas.filter(f=>(ui.formulaCategory==='Semua'||f.category===ui.formulaCategory)&&(!q||[f.name,f.desc,f.syntax,f.example,f.category].join(' ').toLowerCase().includes(q)));slot.innerHTML=list.map(f=>`<article class="formula-card"><div class="formula-title"><h3>${esc(f.name)}</h3><span class="level-pill">${esc(f.category)}</span></div><p>${esc(f.desc)}</p><small>Penulisan umum</small><div style="margin-top:7px"><code>${esc(f.syntax)}</code></div><pre>${esc(formulaForCase(f).example)}</pre><div style="display:flex;align-items:center;justify-content:space-between;gap:9px"><small>Contoh kasus: <b style="color:var(--ink)">${esc(EXCEL_CASES[state.excelCaseId].title)}</b></small><button class="text-link" data-action="try-formula" data-name="${esc(f.name)}">Coba di lembar →</button></div></article>`).join('')||'<div class="panel empty-state">Tidak ada rumus yang cocok.</div>';$('#formula-count').textContent=`${list.length} dari ${OFFICE.excel.formulas.length} fungsi ditampilkan.`;}
// Mesin formula ringan untuk simulator. Rumus dievaluasi dengan parser sendiri,
// TANPA eval / Function / pengiriman data ke server.
function addressParts(id){const m=String(id).toUpperCase().match(/^\$?([A-Z]{1,3})\$?([1-9]\d*)$/);if(!m)return null;let col=0;for(const c of m[1])col=col*26+c.charCodeAt(0)-64;return {col,row:Number(m[2])};}
function colName(num){let text='';while(num>0){num--;text=String.fromCharCode(65+num%26)+text;num=Math.floor(num/26);}return text;}
function rangeObj(a,b,stack){const p=addressParts(a),q=addressParts(b);if(!p||!q)throw Error('#REF!');const loCol=Math.min(p.col,q.col),hiCol=Math.max(p.col,q.col),loRow=Math.min(p.row,q.row),hiRow=Math.max(p.row,q.row);if((hiCol-loCol+1)*(hiRow-loRow+1)>1500)throw Error('#RANGE!');const rows=[];for(let r=loRow;r<=hiRow;r++){const row=[];for(let c=loCol;c<=hiCol;c++)row.push(getCellValue(`${colName(c)}${r}`,stack));rows.push(row);}return {__range:true,rows,values:rows.flat()};}
function getCellValue(ref,stack=new Set()){const id=String(ref).replace(/\$/g,'').toUpperCase();if(!addressParts(id))throw Error('#REF!');if(stack.has(id))throw Error('#CIRC!');const raw=String(state.excelCells[id]??'').trim();if(!raw)return '';if(raw[0]!=='='){if(/^-?(\d+(\.\d+)?|\.\d+)$/.test(raw))return Number(raw);return raw;}stack.add(id);try{return calcFormula(raw.slice(1),stack);}finally{stack.delete(id);}}
function parseTokens(str){const result=[];let i=0;while(i<str.length){const s=str.slice(i);if(/^\s/.test(s)){i++;continue;}const quoted=s.match(/^"(?:[^"]|"")*"/);if(quoted){result.push({t:'str',v:quoted[0].slice(1,-1).replace(/""/g,'"')});i+=quoted[0].length;continue;}const number=s.match(/^(?:\d+(?:\.\d*)?|\.\d+)/);if(number){result.push({t:'num',v:Number(number[0])});i+=number[0].length;continue;}const ident=s.match(/^[A-Za-z_$][A-Za-z0-9_.$]*/);if(ident){result.push({t:'id',v:ident[0]});i+=ident[0].length;continue;}const op=s.match(/^(>=|<=|<>|[+\-*/^&=<>():;,%])/);if(op){result.push({t:'op',v:op[0]});i+=op[0].length;continue;}throw Error('#NAME?');}result.push({t:'end',v:'END'});return result;}
function numberValue(x){if(x&&x.__range)throw Error('#VALUE!');if(typeof x==='number')return x;if(typeof x==='boolean')return x?1:0;if(x===''||x==null)return 0;const n=Number(x);if(!Number.isFinite(n))throw Error('#VALUE!');return n;}
function compareValue(a,b){const na=(typeof a==='number'||typeof b==='number')&&a!==''&&b!==''?Number(a):NaN,nb=Number(b);if(Number.isFinite(na)&&Number.isFinite(nb))return na===nb?0:na>nb?1:-1;const x=String(a??'').toLocaleLowerCase('id'),y=String(b??'').toLocaleLowerCase('id');return x===y?0:x>y?1:-1;}
function opValue(op,a,b){const av=typeof a==='string'?a:a??'',bv=typeof b==='string'?b:b??'';if(op==='&')return String(av)+String(bv);if(['=','<>','>','>=','<','<='].includes(op)){const cmp=compareValue(a,b);return op==='='?cmp===0:op==='<>'?cmp!==0:op==='>'?cmp>0:op==='>='?cmp>=0:op==='<'?cmp<0:cmp<=0;}const x=numberValue(a),y=numberValue(b);switch(op){case '+':return x+y;case '-':return x-y;case '*':return x*y;case '/':if(y===0)throw Error('#DIV/0!');return x/y;case '^':return Math.pow(x,y);}throw Error('#NAME?');}
function asRange(x){return x&&x.__range?x:{__range:true,rows:[[x]],values:[x]};}
function numericRange(vals){return vals.flatMap(x=>asRange(x).values).filter(x=>typeof x==='number'&&Number.isFinite(x));}
function strValue(x){return x==null?'':String(x);}
function matchesCriterion(value,crit){const text=String(crit??'');const m=text.match(/^(>=|<=|<>|>|<|=)(.*)$/),op=m?m[1]:'=',rhs=m?m[2]:text;
  if(op==='='&&/[?*]/.test(rhs)){const regex=new RegExp('^'+rhs.replace(/[.+^${}()|[\]\\]/g,'\\$&').replace(/\*/g,'.*').replace(/\?/g,'.')+'$','i');return regex.test(String(value??''));}
  const cmp=compareValue(value,/^-?(\d+(\.\d+)?|\.\d+)$/.test(rhs)?Number(rhs):rhs);
  return op==='='?cmp===0:op==='<>'?cmp!==0:op==='>'?cmp>0:op==='>='?cmp>=0:op==='<'?cmp<0:cmp<=0;
}
function calculateFn(name,args){const fn=name.toUpperCase();if(fn==='IFERROR')return args[0]&&args[0].__error?args[1]??'':args[0];const firstError=args.find(x=>x&&x.__error);if(firstError)throw Error(firstError.__error);
  const numbers=numericRange(args),flat=args.flatMap(a=>asRange(a).values);
  if(fn==='SUM')return numbers.reduce((s,n)=>s+n,0);
  if(fn==='AVERAGE'){if(!numbers.length)throw Error('#DIV/0!');return numbers.reduce((s,n)=>s+n,0)/numbers.length;}
  if(fn==='COUNT')return numbers.length;
  if(fn==='COUNTA')return flat.filter(v=>v!==''&&v!=null).length;
  if(fn==='MIN'||fn==='MAX'){if(!numbers.length)return 0;return fn==='MIN'?Math.min(...numbers):Math.max(...numbers);}
  if(fn==='IF')return !!args[0]?args[1]??'':args[2]??'';
  if(fn==='AND')return flat.every(Boolean);
  if(fn==='OR')return flat.some(Boolean);
  if(fn==='NOT')return !args[0];
  if(fn==='COUNTIF'){const range=asRange(args[0]);return range.values.filter(v=>matchesCriterion(v,args[1])).length;}
  if(fn==='SUMIF'){const range=asRange(args[0]).values,values=asRange(args.length>=3?args[2]:args[0]).values;return range.reduce((s,v,i)=>s+(matchesCriterion(v,args[1])?((typeof values[i]==='number')?values[i]:0):0),0);}
  if(fn==='COUNTIFS'||fn==='SUMIFS'){const shift=fn==='SUMIFS'?1:0,howMany=asRange(args[shift]).values.length;let sum=0;for(let i=0;i<howMany;i++){let okay=true;for(let k=shift;k+1<args.length;k+=2){if(!matchesCriterion(asRange(args[k]).values[i],args[k+1])){okay=false;break;}}if(okay)sum+=shift?(numberValue(asRange(args[0]).values[i])):1;}return sum;}
  if(fn==='ROUND'){const n=numberValue(args[0]),digits=Math.trunc(numberValue(args[1]??0));if(Math.abs(digits)>100)throw Error('#NUM!');const factor=Math.pow(10,digits);return Math.sign(n)*Math.round((Math.abs(n)+Number.EPSILON)*factor)/factor;}
  if(fn==='CONCAT'||fn==='CONCATENATE')return flat.map(strValue).join('');
  if(fn==='LEN')return strValue(args[0]).length;
  if(fn==='LEFT'||fn==='RIGHT'){const src=strValue(args[0]),count=Math.trunc(numberValue(args[1]??1));if(count<0)throw Error('#VALUE!');return fn==='LEFT'?src.slice(0,count):src.slice(Math.max(0,src.length-count));}
  if(fn==='TRIM')return strValue(args[0]).trim().replace(/\s+/g,' ');
  if(fn==='VLOOKUP'){const grid=asRange(args[1]).rows,idx=Math.trunc(numberValue(args[2]));if(idx<1||idx>grid[0].length)throw Error('#REF!');const row=grid.find(r=>compareValue(r[0],args[0])===0);if(!row)throw Error('#N/A');return row[idx-1];}
  if(fn==='XLOOKUP'){const look=asRange(args[1]).values,values=asRange(args[2]).values;const idx=look.findIndex(v=>compareValue(v,args[0])===0);if(idx===-1){if(args.length>=4)return args[3];throw Error('#N/A');}return values[idx]??'';}
  throw Error('#NAME?');
}
function calcFormula(expr,stack){const tokens=parseTokens(expr);let pos=0;
  const peek=()=>tokens[pos],eat=val=>{if(peek()?.v===val){pos++;return true;}return false;};
  const expected=val=>{if(!eat(val))throw Error('#FORMULA!');};
  const prec={'=':1,'<>':1,'>':1,'>=':1,'<':1,'<=':1,'&':2,'+':3,'-':3,'*':4,'/':4,'^':5};
  function primary(){const t=peek();if(eat('+'))return numberValue(primary());if(eat('-'))return -numberValue(primary());if(eat('(')){const v=parseExpr(0);expected(')');return v;}
    if(t.t==='num'||t.t==='str'){pos++;return t.v;}
    if(t.t==='id'){pos++;const id=t.v;if(eat('(')){const args=[];while(!eat(')')){let val;try{val=parseExpr(0);}catch(err){val={__error:err.message||'#ERROR!'};let depth=0;while(peek().t!=='end'){if(peek().v==='('){depth++;}else if(peek().v===')'){if(depth===0)break;depth--;}if(depth===0&&[',',';'].includes(peek().v))break;pos++;}}
        args.push(val);if(eat(')'))break;if(eat(';')||eat(','))continue;throw Error('#FORMULA!');}return calculateFn(id,args);}
      if(eat(':')){const end=peek();if(end.t!=='id')throw Error('#REF!');pos++;return rangeObj(id,end.v,stack);}
      if(id.toUpperCase()==='TRUE')return true;if(id.toUpperCase()==='FALSE')return false;
      if(addressParts(id))return getCellValue(id,stack);throw Error('#NAME?');}
    throw Error('#FORMULA!');
  }
  function parseExpr(min){let left=primary();while(eat('%'))left=numberValue(left)/100;while(peek()&&own(prec,peek().v)&&prec[peek().v]>=min){const op=peek().v,priority=prec[op];pos++;const right=parseExpr(priority+(op==='^'?0:1));left=opValue(op,left,right);}return left;}
  const result=parseExpr(0);if(peek().t!=='end')throw Error('#FORMULA!');if(typeof result==='number'&&!Number.isFinite(result))throw Error('#NUM!');if(result&&result.__range)throw Error('#VALUE!');return result;
}
function safeCell(id){try{return getCellValue(id,new Set());}catch(err){return err.message&&err.message[0]==='#'?err.message:'#ERROR!';}}
function sheetDisplay(id){const raw=String(state.excelCells[id]??'');if(raw.startsWith('=')){const val=safeCell(id);return typeof val==='boolean'?(val?'TRUE':'FALSE'):val==null?'':String(typeof val==='number'?Math.round((val+Number.EPSILON)*1e8)/1e8:val);}return raw;}
function outputValue(id){const val=safeCell(id);return typeof val==='number'?fmt(val):String(val);}
function renderPractice(){return state.course==='excel'?renderExcelPractice():state.course==='word'?renderWordPractice():renderPowerpointPractice();}
function renderExcelPractice(){const sel=ui.activeCell||'E8';const current=EXCEL_CASES[state.excelCaseId]||EXCEL_CASES.keuangan;return `${headerBlock('LABORATORIUM EXCEL','Simulasi register administrasi kampus','Pilih dataset unit, edit data sintetis, masukkan rumus, lalu periksa hasil tantangan.',`<button class="outline-btn" data-action="excel-export">↓ Ekspor CSV</button><button class="outline-btn" data-action="excel-reset">↺ Reset worksheet</button>`)}
  <div class="case-switcher"><label for="excel-case-select"><b>Pilih dataset unit kerja</b></label><select id="excel-case-select">${CAMPUS_CASES.map(c=>`<option value="${c.id}" ${state.excelCaseId===c.id?'selected':''}>${esc(c.name)}</option>`).join('')}</select><button class="outline-btn" data-action="case-download-csv" data-case="${esc(state.excelCaseId)}">↓ Unduh dataset CSV</button><span>Data simulasi · tidak memuat transaksi/identitas riil</span></div><div class="practice-intro">🟢 <b>Lembar kerja aktif:</b> Klik sebuah sel atau pilih tantangan. Ketik rumus langsung di sel, atau tulis pada bilah <b>fx</b> lalu tekan Enter / Terapkan. Lembar ini menghitung ulang sel yang saling merujuk secara otomatis. Pemisah argumen rumus boleh <b>;</b> atau <b>,</b>.</div>
  <div class="lab-grid"><section class="panel lab-panel"><div class="lab-header"><h3>▦ Dataset ${esc(current.title)} · Sheet 1</h3><small>7 kolom × 9 baris</small></div><div class="sheet-toolbar"><span class="cell-ref" id="active-cell-ref">${esc(sel)}</span><span class="fx">fx</span><input class="formula-input" id="formula-bar" type="text" spellcheck="false" aria-label="Bilah rumus" placeholder="Ketik =SUM(E2:E7)" value="${esc(state.excelCells[sel]??'')}"><button class="primary-btn" data-action="excel-commit">Terapkan</button></div><div class="sheet-wrap"><table class="sheet-table" aria-label="Simulator lembar kerja Excel"><thead><tr><th class="row-num"></th>${['A','B','C','D','E','F','G'].map(x=>`<th>${x}</th>`).join('')}</tr></thead><tbody>${Array.from({length:9},(_,n)=>{const row=n+1;return `<tr class="${row===1?'head-row':''}"><th class="row-num">${row}</th>${['A','B','C','D','E','F','G'].map(letter=>{const ref=`${letter}${row}`,raw=String(state.excelCells[ref]??'');return `<td><input class="sheet-input ${raw[0]==='='?'formula-cell':''}" data-cell="${ref}" aria-label="Sel ${ref}" autocomplete="off" spellcheck="false" value="${esc(sheetDisplay(ref))}" title="${esc(ref+': '+raw)}"></td>`;}).join('')}</tr>`;}).join('')}</tbody></table></div><p class="sheet-hint">Contoh: <code>=C2*D2</code>, <code>=SUM(E2:E7)</code>, <code>=IF(D2>=5;"Banyak";"Sedikit")</code>. Geser tabel ke kanan untuk melihat kolom G. Hasil simulator bukan pengganti mesin kalkulasi Excel lengkap.</p><div id="excel-feedback">${ui.practiceMessage?renderPracticeMessage():''}</div></section>
  <aside class="panel"><div class="lab-header"><h3>🎯 Tantangan interaktif</h3><small>${state.excelDone.length} / ${EXCEL_TASKS.length} selesai</small></div><div class="task-list">${EXCEL_TASKS.map((t,i)=>`<div class="task-row ${ui.activeTask===i?'active':''} ${state.excelDone.includes(t.id)?'done':''}"><div class="task-row-top"><span class="task-index">${state.excelDone.includes(t.id)?'✓':i+1}</span><strong>${esc(t.title)}</strong></div><p class="task-desc">${esc(taskHintText(t.desc))}</p><div class="task-foot"><code>${t.cell}</code><button class="mini-button" data-action="excel-select-task" data-index="${i}">Pilih</button><button class="mini-button" data-action="excel-hint" data-index="${i}">Petunjuk</button></div></div>`).join('')}</div><div style="display:flex;gap:8px;margin-top:17px;flex-wrap:wrap"><button class="primary-btn" data-action="excel-check">Periksa tantangan</button><button class="outline-btn" data-action="excel-check-all">Periksa semua</button></div></aside></div>`;
}
function renderPracticeMessage(){const m=ui.practiceMessage;return `<div class="challenge-msg ${m?.type||''}">${esc(m?.text||'')}</div>`;}
function setExcelCell(id,raw){const ref=String(id).toUpperCase();if(!addressParts(ref))return;state.excelCells[ref]=String(raw);ui.activeCell=ref;save();refreshExcelSheet();}
function refreshExcelSheet(){const block=$('#main-view');if(!block||state.section!=='practice'||state.course!=='excel')return;$$('.sheet-input',block).forEach(el=>{const addr=el.dataset.cell;if(document.activeElement!==el){el.value=sheetDisplay(addr);el.title=addr+': '+(state.excelCells[addr]??'');el.classList.toggle('formula-cell',String(state.excelCells[addr]??'').startsWith('='));}});$('#active-cell-ref').textContent=ui.activeCell;const bar=$('#formula-bar');if(bar&&document.activeElement!==bar)bar.value=state.excelCells[ui.activeCell]??'';}
function selectExcelCell(ref,focus=true){ui.activeCell=ref;const bar=$('#formula-bar');if(bar)bar.value=state.excelCells[ref]??'';const label=$('#active-cell-ref');if(label)label.textContent=ref;if(focus){const input=$(`.sheet-input[data-cell="${ref}"]`);input?.focus();input?.select();}}
function excelCommit(){const bar=$('#formula-bar');if(!bar)return;setExcelCell(ui.activeCell,bar.value);const input=$(`.sheet-input[data-cell="${ui.activeCell}"]`);if(input)input.value=sheetDisplay(ui.activeCell);toast(`Sel ${ui.activeCell} diperbarui.`);}
function taskHintText(text){const c=EXCEL_CASES[state.excelCaseId]||EXCEL_CASES.keuangan;return String(text).replaceAll('{criterion}',c.criterion).replaceAll('{lookup}',c.lookup);}
function taskHint(task){return taskHintText(task.hint);}
function expectedTask(t){
 const get=id=>safeCell(id),num=id=>Number(get(id)),rows=[2,3,4,5,6,7];
 const src=EXCEL_CASES[state.excelCaseId]||EXCEL_CASES.keuangan,crit=src.criterion.toLowerCase(),lookup=src.lookup.toLowerCase();
 const subset=rows.filter(i=>String(get('B'+i)).toLowerCase()===crit);
 const found=rows.find(i=>String(get('A'+i)).toLowerCase()===lookup);
 switch(t.id){
  case 'gap':return rows.reduce((s,i)=>s+num('E'+i),0);
  case 'sumTarget':return rows.reduce((s,i)=>s+num('C'+i),0);
  case 'sumActual':return rows.reduce((s,i)=>s+num('D'+i),0);
  case 'if':return num('D2')>num('C2')?'Melebihi':'Dalam batas';
  case 'average':return rows.reduce((s,i)=>s+num('D'+i),0)/rows.length;
  case 'countif':return subset.length;
  case 'sumif':return subset.reduce((s,i)=>s+num('D'+i),0);
  case 'round':return Math.round((num('D2')/num('C2')*100+Number.EPSILON)*10)/10;
  case 'concat':return String(get('A3'))+' - '+String(get('B3'));
  case 'xlookup':return found?num('D'+found):undefined;
  case 'vlookup':return found?num('D'+found):undefined;
  case 'max':return Math.max(...rows.map(i=>num('D'+i)));
  case 'countifs':return subset.filter(i=>num('D'+i)>=100).length;
  case 'sumifs':return subset.filter(i=>num('C'+i)>=100).reduce((s,i)=>s+num('D'+i),0);
  default:return undefined;
 }
}
function validateTask(t){const raw=String(state.excelCells[t.cell]??'').trim(),actual=safeCell(t.cell),expected=expectedTask(t);if(!new RegExp('^=\\s*'+t.fn+'\\s*\\(','i').test(raw))return {ok:false,message:`Gunakan fungsi ${t.fn} pada sel ${t.cell}.`};if(typeof actual==='string'&&actual.startsWith('#'))return {ok:false,message:`Rumus menghasilkan ${actual}. Periksa referensi dan argumen.`};const match=typeof actual==='number'&&typeof expected==='number'?Number.isFinite(expected)&&Math.abs(actual-expected)<1e-7:String(actual).toLowerCase()===String(expected).toLowerCase();if(match)return {ok:true,message:`Benar! ${t.cell} = ${typeof actual==='number'?fmt(actual):actual}.`};return {ok:false,message:`Belum sesuai. Hasil ${t.cell}: ${typeof actual==='number'?fmt(actual):actual||'(kosong)'}. Periksa rentang dan kondisi rumus.`};}
function checkTask(index,all=false){const tasks=all?EXCEL_TASKS:[EXCEL_TASKS[index]],result=tasks.map(t=>({task:t,...validateTask(t)}));let changed=false;for(const x of result)if(x.ok&&!state.excelDone.includes(x.task.id)){state.excelDone.push(x.task.id);changed=true;}if(changed)save();if(all){const good=result.filter(x=>x.ok).length;ui.practiceMessage={type:good===result.length?'success':good?'':'error',text:`${good} dari ${result.length} rumus sudah benar. Pilih tantangan yang belum selesai dan periksa lagi.`};}else{ui.practiceMessage={type:result[0].ok?'success':'error',text:result[0].message};}render();}
function exportExcelCSV(){const rows=[];for(let r=1;r<=9;r++){const row=[];for(let c=1;c<=7;c++){const v=safeCell(colName(c)+r);row.push('"'+String(v==null?'':v).replace(/"/g,'""')+'"');}rows.push(row.join(','));}download('OfficeLab_v21_Excel_'+state.excelCaseId+'.csv','\ufeff'+rows.join('\r\n'),'text/csv;charset=utf-8');}
function wordTaskChecks(){const editor=$('#word-editor');if(!editor)return {};return {
  heading:!!editor.querySelector('h1'),bold:!!editor.querySelector('b,strong'),italic:!!editor.querySelector('i,em'),bullets:editor.querySelectorAll('ul li').length>=2,table:!!editor.querySelector('table')
};}
function cleanWordHTML(html){const t=document.createElement('template');t.innerHTML=String(html??'');t.content.querySelectorAll('script,iframe,object,embed,style,link,meta,form,video,audio,svg').forEach(x=>x.remove());t.content.querySelectorAll('*').forEach(x=>{for(const attr of [...x.attributes])if(/^on/i.test(attr.name)||['srcdoc','contenteditable'].includes(attr.name)||(['href','src'].includes(attr.name)&&!/^https?:|^#|^data:image\/(?:png|jpeg|gif);base64,/i.test(attr.value)))x.removeAttribute(attr.name);});return t.innerHTML;}
function refreshWordProgress(){const editor=$('#word-editor');if(!editor)return;const checks=wordTaskChecks();let changed=false;WORD_TASKS.forEach(t=>{if(checks[t.id]&&!state.wordDone.includes(t.id)){state.wordDone.push(t.id);changed=true;}});if(changed)save();const slot=$('#word-checks');if(slot)slot.innerHTML=WORD_TASKS.map(t=>`<div class="check-row ${state.wordDone.includes(t.id)?'done':''}"><span class="check-icon">${state.wordDone.includes(t.id)?'✓':'○'}</span><span><b>${esc(t.title)}</b><br>${esc(t.desc)}</span></div>`).join('');const total=$('#word-score');if(total)total.textContent=`${state.wordDone.length} / ${WORD_TASKS.length}`;const bar=$('#word-bar');if(bar)bar.style.width=(100*state.wordDone.length/WORD_TASKS.length)+'%';if(changed)renderNavigation();}
function renderWordPractice(){return `${headerBlock('LABORATORIUM WORD','Editor dokumen administrasi','Sorot teks kemudian pilih format pada toolbar. Buat heading, daftar berpoin, tabel, dan rapikan isi dokumen.',`<button class="outline-btn" data-action="word-export">↓ Ekspor HTML</button><button class="outline-btn" data-action="word-print">⎙ Cetak</button><button class="outline-btn" data-action="word-reset">↺ Reset dokumen</button>`)}
<div class="practice-intro">🔵 <b>Latihan Word:</b> Dokumen di bawah dapat diketik dan diformat. Gunakan tombol <b>B</b>, <b>I</b>, <b>U</b>, daftar, atau pilih <b>Heading 1</b>. Perubahan tersimpan di browser. Fitur ini berupa editor latihan, bukan Microsoft Word penuh.</div>
<div class="editor-grid"><section class="panel no-pad" style="overflow:hidden"><div class="word-toolbar"><select class="style-select" id="word-style" aria-label="Gaya teks"><option value="p">Normal</option><option value="h1">Heading 1</option><option value="h2">Heading 2</option></select><span class="tool-separator"></span><button class="tool-btn" data-action="word-command" data-cmd="bold" title="Bold (Ctrl+B)" aria-label="Tebal"><b>B</b></button><button class="tool-btn" data-action="word-command" data-cmd="italic" title="Italic (Ctrl+I)" aria-label="Miring"><i>I</i></button><button class="tool-btn" data-action="word-command" data-cmd="underline" title="Underline (Ctrl+U)" aria-label="Garis bawah"><u>U</u></button><span class="tool-separator"></span><button class="tool-btn" data-action="word-command" data-cmd="justifyLeft" title="Rata kiri">≡</button><button class="tool-btn" data-action="word-command" data-cmd="justifyCenter" title="Rata tengah">≣</button><button class="tool-btn" data-action="word-command" data-cmd="justifyRight" title="Rata kanan">☷</button><button class="tool-btn" data-action="word-command" data-cmd="justifyFull" title="Rata kiri-kanan">▤</button><span class="tool-separator"></span><button class="tool-btn" data-action="word-command" data-cmd="insertUnorderedList" title="Bullet list">• Daftar</button><button class="tool-btn" data-action="word-command" data-cmd="insertOrderedList" title="Numbering">1. Daftar</button><button class="tool-btn" data-action="word-table" title="Sisipkan tabel 2 × 2">▦ Tabel</button></div><div class="word-canvas-wrap"><div class="word-paper" id="word-editor" contenteditable="true" role="textbox" aria-label="Editor dokumen Word" aria-multiline="true" spellcheck="true">${cleanWordHTML(state.wordHTML)}</div></div></section>
<aside class="panel"><h3 style="margin:0 0 7px">🎯 Target latihan</h3><p class="sub-title" style="margin-bottom:15px">Lakukan lima tindakan berikut di dokumen. Checklist akan tercentang otomatis.</p><div class="scorebar"><strong id="word-score">${state.wordDone.length} / ${WORD_TASKS.length}</strong><div class="track"><div id="word-bar" class="track-fill" style="width:${100*state.wordDone.length/WORD_TASKS.length}%"></div></div></div><div class="word-checklist" id="word-checks"></div><div class="pill-note" style="margin-top:17px"><b>Tip:</b> Untuk membuat daftar dua butir, pilih “• Daftar”, ketik item pertama, tekan Enter, lalu ketik item kedua. Simpan latihan ke HTML jika ingin menyimpan salinan terpisah.</div></aside></div>`;}
function restoreWordSelection(){const editor=$('#word-editor');if(!editor)return;editor.focus();if(ui.wordRange){try{const sel=window.getSelection();sel.removeAllRanges();sel.addRange(ui.wordRange);}catch(_){}}}
function wordCommand(cmd){restoreWordSelection();try{document.execCommand(cmd,false,null);}catch(_){toast('Perintah ini tidak tersedia di browser Anda.');}const e=$('#word-editor');if(e){state.wordHTML=cleanWordHTML(e.innerHTML);save();refreshWordProgress();}}
function insertWordTable(){restoreWordSelection();try{document.execCommand('insertHTML',false,'<table><tbody><tr><th>Kolom 1</th><th>Kolom 2</th></tr><tr><td>Isi 1</td><td>Isi 2</td></tr></tbody></table><p><br></p>');}catch(_){toast('Sisipkan tabel tidak didukung browser ini.');}const e=$('#word-editor');if(e){state.wordHTML=cleanWordHTML(e.innerHTML);save();refreshWordProgress();}}
function exportWord(){const page=`<!doctype html><html lang="id"><head><meta charset="utf-8"><title>Latihan Word - OfficeLab</title><style>body{font:16px/1.7 Georgia,serif;max-width:760px;margin:48px auto;padding:0 25px;color:#213047}h1,h2{line-height:1.3}table{border-collapse:collapse;width:100%}td,th{border:1px solid #777;padding:9px}@media print{body{margin:0;max-width:none}}</style></head><body>${cleanWordHTML(state.wordHTML)}</body></html>`;download('OfficeLab_Latihan_Word.html',page,'text/html;charset=utf-8');}
function printWord(){const pop=window.open('','_blank','width=840,height=850');if(!pop){toast('Popup cetak diblokir oleh browser. Izinkan popup atau gunakan ekspor HTML.');return;}pop.document.write(`<!doctype html><html><head><title>Cetak - Latihan Word OfficeLab</title><style>body{max-width:760px;margin:45px auto;padding:0 25px;color:#222;font:15px/1.7 Georgia,serif}table{border-collapse:collapse;width:100%}td,th{border:1px solid #666;padding:8px}@media print{body{margin:0;max-width:none}}</style></head><body>${cleanWordHTML(state.wordHTML)}</body></html>`);pop.document.close();pop.focus();setTimeout(()=>{pop.print();},250);}
function checkPptGoals(){const slides=state.pptSlides;const titles=slides.map(s=>(s.title||'').trim()).filter(Boolean);return {slides:slides.length>=3,titles:new Set(titles.map(x=>x.toLowerCase())).size>=3,bullets:slides.some(s=>(s.body||'').split('\n').filter(x=>x.trim()).length>=3),theme:state.pptTheme!=='blue',notes:slides.some(s=>(s.notes||'').trim().length>=8),present:!!state.pptPresented};}
function refreshPptGoals(){if(state.course!=='powerpoint')return;const checks=checkPptGoals();let changed=false;PPT_TASKS.forEach(t=>{if(checks[t.id]&&!state.pptDone.includes(t.id)){state.pptDone.push(t.id);changed=true;}});if(changed){save();renderNavigation();}const slot=$('#ppt-checks');if(slot)slot.innerHTML=PPT_TASKS.map(t=>`<div class="check-row ${state.pptDone.includes(t.id)?'done':''}"><span class="check-icon">${state.pptDone.includes(t.id)?'✓':'○'}</span><span><b>${esc(t.title)}</b><br>${esc(t.desc)}</span></div>`).join('');const score=$('#ppt-score');if(score)score.textContent=`${state.pptDone.length} / ${PPT_TASKS.length}`;}
function slideHTML(slide,theme,index){const body=(slide.body||'').split('\n').filter(x=>x.trim()).map(x=>esc(x)).join('\n');return `<div class="slide-screen theme-${esc(theme)} layout-${esc(slide.layout||'content')}"><h2>${esc(slide.title||'Tanpa judul')}</h2><p>${body}</p><span class="slide-foot">OFFICELAB · ${index+1}</span></div>`;}
function renderPowerpointPractice(){ui.pptIndex=Math.max(0,Math.min(ui.pptIndex,state.pptSlides.length-1));const index=ui.pptIndex,slide=state.pptSlides[index],themes=[['blue','Biru'],['purple','Ungu'],['green','Hijau'],['dark','Gelap'],['sunset','Senja']];return `${headerBlock('LABORATORIUM POWERPOINT','Studio paparan kinerja unit','Sunting judul, isi, layout, tema dan catatan pembicara. Tambah slide lalu jalankan presentasi secara langsung.',`<button class="primary-btn" data-action="ppt-present">▶ Presentasikan</button><button class="outline-btn" data-action="ppt-export">↓ Ekspor HTML</button><button class="outline-btn" data-action="ppt-reset">↺ Reset</button>`)}
<div class="practice-intro">🟠 <b>Latihan PowerPoint:</b> Pilih thumbnail slide di kiri, sunting pada panel kanan, dan lihat pratinjau secara langsung. Presentasi berjalan sepenuhnya offline. Hasil ekspor berupa HTML mandiri, bukan file .pptx native.</div>
<div class="ppt-editor"><aside><div class="lab-header"><h3>Slide</h3><small>${state.pptSlides.length} buah</small></div><div class="slides-list">${state.pptSlides.map((s,i)=>`<button type="button" class="slide-thumb ${i===index?'active':''}" data-action="ppt-select" data-index="${i}" aria-label="Buka slide ${i+1}"><span class="thumb-num">${i+1}</span><span class="thumb-preview"><strong>${esc((s.title||'Slide baru').slice(0,35))}</strong></span></button>`).join('')}</div><div class="mini-tools"><button class="outline-btn" style="width:100%" data-action="ppt-add">＋ Tambah slide</button></div></aside>
<section><div class="lab-header"><h3>Pratinjau slide ${index+1}</h3><small>Rasio 16:9</small></div><div class="slide-stage" id="ppt-stage">${slideHTML(slide,state.pptTheme,index)}</div><div class="slide-mini-actions"><button data-action="ppt-up" ${index===0?'disabled':''}>↑ Pindah atas</button><button data-action="ppt-down" ${index===state.pptSlides.length-1?'disabled':''}>↓ Pindah bawah</button><button data-action="ppt-duplicate">▣ Duplikasi</button><button data-action="ppt-delete" ${state.pptSlides.length<=1?'disabled':''}>✕ Hapus</button></div><div class="panel ppt-goals"><div class="lab-header"><h3>🎯 Tantangan slide</h3><small id="ppt-score">${state.pptDone.length} / ${PPT_TASKS.length}</small></div><div class="word-checklist" id="ppt-checks"></div></div></section>
<aside class="panel slide-edit-panel"><h3 style="margin:0 0 17px">⚙ Pengaturan slide</h3><div class="field"><label for="ppt-title">Judul slide</label><input id="ppt-title" type="text" value="${esc(slide.title)}" maxlength="180" placeholder="Tulis judul ..."></div><div class="field"><label for="ppt-body">Isi / poin utama</label><textarea id="ppt-body" rows="5" placeholder="Pisahkan setiap poin dengan Enter">${esc(slide.body)}</textarea><small>Enter membuat baris baru pada slide.</small></div><div class="field"><label for="ppt-layout">Layout</label><select id="ppt-layout"><option value="content" ${slide.layout==='content'?'selected':''}>Judul + isi</option><option value="cover" ${slide.layout==='cover'?'selected':''}>Judul besar</option><option value="split" ${slide.layout==='split'?'selected':''}>Dua kolom</option><option value="quote" ${slide.layout==='quote'?'selected':''}>Kutipan</option></select></div><div class="field"><label>Tema warna (seluruh slide)</label><div class="theme-swatches">${themes.map(([t,label])=>`<button class="swatch theme-${t} ${state.pptTheme===t?'selected':''}" data-action="ppt-theme" data-theme="${t}" type="button" title="Tema ${label}" aria-label="Tema ${label}" aria-pressed="${state.pptTheme===t}"></button>`).join('')}</div></div><div class="field"><label for="ppt-notes">Catatan pembicara</label><textarea id="ppt-notes" rows="4" placeholder="Tulis detail yang tidak muncul di slide...">${esc(slide.notes)}</textarea><small>Catatan tidak ditampilkan pada mode presentasi.</small></div></aside></div>`;}
function pptLiveUpdate(){const slide=state.pptSlides[ui.pptIndex],stage=$('#ppt-stage');if(!slide||!stage)return;stage.innerHTML=slideHTML(slide,state.pptTheme,ui.pptIndex);const thumb=$(`.slide-thumb[data-index="${ui.pptIndex}"] .thumb-preview strong`);if(thumb)thumb.textContent=(slide.title||'Slide baru').slice(0,35);refreshPptGoals();}
function changePpt(action){const i=ui.pptIndex;switch(action){
  case 'ppt-add':state.pptSlides.push({title:'Slide baru',body:'',layout:'content',notes:''});ui.pptIndex=state.pptSlides.length-1;break;
  case 'ppt-duplicate':state.pptSlides.splice(i+1,0,{...clone(state.pptSlides[i]),title:state.pptSlides[i].title+' (salinan)'});ui.pptIndex=i+1;break;
  case 'ppt-delete':if(state.pptSlides.length<=1)return;if(!confirm('Hapus slide '+(i+1)+'?'))return;state.pptSlides.splice(i,1);ui.pptIndex=Math.min(i,state.pptSlides.length-1);break;
  case 'ppt-up':if(i<=0)return;[state.pptSlides[i-1],state.pptSlides[i]]=[state.pptSlides[i],state.pptSlides[i-1]];ui.pptIndex=i-1;break;
  case 'ppt-down':if(i>=state.pptSlides.length-1)return;[state.pptSlides[i+1],state.pptSlides[i]]=[state.pptSlides[i],state.pptSlides[i+1]];ui.pptIndex=i+1;break;
}save();render();}
function startPresentation(){state.pptPresented=true;save();refreshPptGoals();ui.presentationIndex=ui.pptIndex;$('#presentation-overlay').classList.remove('hidden');document.body.style.overflow='hidden';refreshPresentation();}
function refreshPresentation(){const slides=state.pptSlides;ui.presentationIndex=Math.max(0,Math.min(ui.presentationIndex,slides.length-1));$('#presentation-content').innerHTML=slideHTML(slides[ui.presentationIndex],state.pptTheme,ui.presentationIndex);$('#presentation-count').textContent=`Slide ${ui.presentationIndex+1} dari ${slides.length}`;$('#presentation-prev').disabled=ui.presentationIndex===0;$('#presentation-next').disabled=ui.presentationIndex>=slides.length-1;}
function closePresentation(){$('#presentation-overlay').classList.add('hidden');document.body.style.overflow='';}
function exportPpt(){const slides=state.pptSlides.map((s,i)=>slideHTML(s,state.pptTheme,i));const page=`<!doctype html><html lang="id"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Presentasi OfficeLab</title><style>*{box-sizing:border-box}body{margin:0;background:#101623;color:#fff;font-family:Segoe UI,Arial,sans-serif;height:100vh;display:flex;flex-direction:column}.screen{flex:1;display:grid;place-items:center;padding:20px;min-height:0}.slide-screen{width:min(100%,calc((100vh - 90px)*1.777));aspect-ratio:16/9;position:relative;display:flex;flex-direction:column;justify-content:center;padding:6% 8%;background:linear-gradient(120deg,#21468a,#5786d5);overflow:hidden}.slide-screen h2{font-size:clamp(30px,5vw,72px);line-height:1.15;max-width:90%;margin:0 0 3%}.slide-screen p{font-size:clamp(17px,2.3vw,34px);line-height:1.5;max-width:90%;white-space:pre-line;margin:0}.slide-foot{position:absolute;left:8%;bottom:6%;opacity:.6;font-size:12px}.theme-purple{background:linear-gradient(125deg,#39277c,#8d66cf)}.theme-green{background:linear-gradient(125deg,#0d6556,#2bbd97)}.theme-dark{background:linear-gradient(125deg,#16223b,#344761)}.theme-sunset{background:linear-gradient(125deg,#bb4b26,#ed9559)}.layout-cover h2{font-size:clamp(37px,5.4vw,90px)}.layout-split{flex-direction:row;justify-content:space-around;align-items:center;gap:5%}.layout-split h2{max-width:44%}.layout-split p{max-width:50%}.layout-quote{text-align:center;align-items:center}.layout-quote h2{font-style:italic}.controls{height:60px;display:flex;align-items:center;justify-content:center;gap:15px;font-size:13px}button{background:#ffffff24;color:#fff;border:1px solid #ffffff44;padding:9px 15px;border-radius:8px;cursor:pointer}</style></head><body><main class="screen" id="screen"></main><div class="controls"><button id="prev">←</button><span id="counter"></span><button id="next">→</button></div><script>const SLIDES=${JSON.stringify(slides).replace(/</g,'\\u003c')};let i=0;function show(){document.getElementById('screen').innerHTML=SLIDES[i];document.getElementById('counter').textContent=(i+1)+' / '+SLIDES.length;}document.getElementById('prev').onclick=()=>{i=Math.max(0,i-1);show()};document.getElementById('next').onclick=()=>{i=Math.min(SLIDES.length-1,i+1);show()};document.addEventListener('keydown',e=>{if(['ArrowRight',' ','PageDown'].includes(e.key)){i=Math.min(SLIDES.length-1,i+1);show();}if(['ArrowLeft','PageUp'].includes(e.key)){i=Math.max(0,i-1);show();}});show();<\/script></body></html>`;
  download('OfficeLab_Presentasi_Offline.html',page,'text/html;charset=utf-8');}
function shuffle(items){const out=[...items];for(let i=out.length-1;i>0;i--){const j=Math.floor(Math.random()*(i+1));[out[i],out[j]]=[out[j],out[i]];}return out;}
function startQuiz(mode='level',level=state.selectedLevel[state.course]){
  const id=state.course;
  let questions=[],levels=[];
  if(mode==='general')questions=shuffle(OFFICE[id].quiz).slice(0,10);
  else if(mode==='placement'){
    for(const lev of LEVELS){const items=shuffle(levelIndices(id,lev.id));for(const i of items.slice(0,2)){const c=OFFICE[id].lessons[i].check;questions.push([c.question,c.options,c.answer,c.why]);levels.push(lev.id);}}
    // Urutkan level untuk memudahkan interpretasi (masing-masing dua soal).
  }else{
    if(!LEVELS.some(l=>l.id===level))level='Basic';
    questions=shuffle(levelIndices(id,level).map(i=>{const c=OFFICE[id].lessons[i].check;return [c.question,c.options,c.answer,c.why];}));
    mode='level';
  }
  ui.quiz={course:id,mode,level,levels,questions,index:0,selected:null,checked:false,correct:0,correctByLevel:{},finished:false};render();
}
function quizCurrent(){return ui.quiz&&ui.quiz.course===state.course?ui.quiz:null;}
function finishQuiz(){const q=quizCurrent();if(!q)return;q.finished=true;
  const score=Math.round(q.correct/q.questions.length*100),id=state.course;
  if(q.mode==='level'){
    state.levelLast[id][q.level]=score;
    state.levelBest[id][q.level]=Math.max(state.levelBest[id][q.level]||0,score);
    state.levelAttempts[id][q.level]=(state.levelAttempts[id][q.level]||0)+1;
  }else if(q.mode==='general'){
    state.quizLast[id]=score;state.quizBest[id]=Math.max(state.quizBest[id],score);
  }else{
    const suggested=LEVELS.find(lev=>(q.correctByLevel[lev.id]||0)<2)?.id||'Expert';
    q.recommendation=suggested;
  }
  save();render();
}
function renderQuiz(){const id=state.course,c=OFFICE[id],q=quizCurrent();
  if(!q){
    const cards=LEVELS.map(lev=>{const st=levelStats(id,lev.id),attempts=state.levelAttempts[id][lev.id]||0;return `<article class="exam-card"><div class="exam-top"><span class="level-number" style="--level-color:${lev.color}">${lev.label}</span><span class="${st.passed?'mini-pass':'mini-pending'}">${st.passed?'✓ Lulus':'Belum lulus'}</span></div><h3>${lev.name}</h3><p>5 soal pilihan ganda dari 5 modul level ${lev.name}.</p><div class="exam-card-meta">Nilai terbaik <strong>${attempts?st.best+'%':'—'}</strong> · ${attempts} percobaan</div><button class="${st.passed?'outline-btn':'primary-btn'}" type="button" data-action="quiz-start" data-mode="level" data-level="${lev.id}">${attempts?'Ulangi evaluasi':'Mulai evaluasi'} →</button></article>`;}).join('');
    return `${headerBlock('EVALUASI PER LEVEL','Kuis '+c.name,`Setiap level memiliki tes berisi lima pertanyaan. Nilai kelulusan ${PASS_SCORE}%. Anda juga dapat mengerjakan tes penempatan dan kuis umum.`)}<div class="exam-grid">${cards}</div>
     <div class="mini-grid-2" style="margin-top:20px"><article class="mini-guide"><strong>◎ Tes penempatan (diagnostik)</strong><p>10 pertanyaan: 2 dari setiap level. Hasil memberi saran titik mulai, bukan sertifikasi kemampuan.</p><button class="text-link" data-action="quiz-start" data-mode="placement">Mulai tes penempatan →</button></article><article class="mini-guide"><strong>◈ Kuis umum ${c.name}</strong><p>10 soal acak dari bank soal OfficeLab. Skor terbaik: <b>${state.quizBest[id]}%</b>.</p><button class="text-link" data-action="quiz-start" data-mode="general">Mulai kuis umum →</button></article></div><div class="pill-note">Kelulusan level ditentukan oleh skor tes minimal 80%; status <b>Level Tuntas</b> juga memerlukan 5 materi dan 5 latihan mikro selesai. Tidak ada batas percobaan.</div>`;
  }
  const label=q.mode==='level'?`Tes ${q.level}`:q.mode==='placement'?'Tes penempatan':'Kuis umum';
  const side=`<aside class="panel quiz-side"><span class="eyebrow">EVALUASI OFFLINE</span><h3 style="font-size:18px;margin:13px 0 9px">${esc(label)}</h3><p class="sub-title">${q.mode==='level'?'Lima pertanyaan menguji konsep inti setiap modul level ini.':q.mode==='placement'?'Dua pertanyaan per level sebagai indikasi awal pengalaman Anda.':'Sepuluh pertanyaan acak dari bank soal umum.'}</p><div class="pill-note">Nilai batas lulus tes level adalah <b>${PASS_SCORE}%</b>. Jawaban diperiksa satu per satu dan dilengkapi pembahasan.</div><button class="outline-btn" style="width:100%;margin-top:16px" data-action="quiz-home">← Pilih evaluasi lain</button></aside>`;
  if(q.finished){const score=Math.round(q.correct/q.questions.length*100),passes=score>=PASS_SCORE,summary=q.mode==='level'?`<p>Nilai terbaik level ${q.level}: <b>${state.levelBest[id][q.level]}%</b>. ${passes?'Tes level dinyatakan lulus.':'Belum mencapai nilai 80%; ulangi materi dan coba kembali.'}</p>`:q.mode==='placement'?`<p>Rekomendasi awal: mulai dari <b>${q.recommendation}</b>. Hanya 2 pertanyaan per level, sehingga ini bukan ukuran kompetensi menyeluruh.</p><div class="placement-results">${LEVELS.map(lev=>`<span>${lev.id}: <b>${q.correctByLevel[lev.id]||0}/2</b></span>`).join('')}</div>`:`<p>Nilai terbaik kuis umum: <b>${state.quizBest[id]}%</b>.</p>`;
    return `${headerBlock('HASIL EVALUASI',`Hasil ${label}`,'Penilaian dilakukan di perangkat Anda; skor disimpan secara lokal untuk tes per level dan kuis umum.')}<div class="quiz-layout"><div class="panel quiz-result"><span class="eyebrow">EVALUASI SELESAI</span><div class="big-score">${score}%</div><h2>${passes?'Target nilai tercapai':'Lanjutkan latihan untuk memperkuat pemahaman'}</h2><p>Jawaban benar ${q.correct} dari ${q.questions.length} pertanyaan.</p>${summary}<div class="quiz-actions" style="justify-content:center;margin-top:23px"><button class="outline-btn" data-action="quiz-home">Kembali ke daftar evaluasi</button><button class="primary-btn" data-action="quiz-start" data-mode="${q.mode}" data-level="${q.level}">Kerjakan ulang ↺</button>${q.mode==='placement'?`<button class="outline-btn" data-action="level-select" data-level="${q.recommendation}">Buka ${q.recommendation} →</button>`:''}</div></div>${side}</div>`;
  }
  const question=q.questions[q.index],chosen=q.selected,checked=q.checked,pct=(q.index+1)/q.questions.length*100;
  return `${headerBlock('EVALUASI',label,'Pilih satu jawaban, periksa, lalu pelajari pembahasannya sebelum melanjutkan.')}
  <div class="quiz-layout"><section class="panel quiz-card"><div class="quiz-topline"><span>PERTANYAAN ${q.index+1} / ${q.questions.length}</span><span>${q.correct} jawaban benar</span></div><div class="track" style="margin-bottom:23px"><div class="track-fill" style="width:${pct}%"></div></div><h2 class="quiz-question">${esc(question[0])}</h2><div class="option-list">${question[1].map((opt,i)=>{let cls=chosen===i?'selected':'';if(checked&&i===question[2])cls='correct';else if(checked&&chosen===i&&i!==question[2])cls='incorrect';return `<button class="quiz-option ${cls}" data-action="quiz-select" data-index="${i}" ${checked?'disabled':''}><span class="option-letter">${'ABCD'[i]}</span>${esc(opt)}</button>`;}).join('')}</div>${checked?`<div class="quiz-feedback"><strong>${chosen===question[2]?'✓ Tepat!':'✕ Belum tepat.'}</strong> ${esc(question[3])}</div>`:''}<div class="quiz-actions">${checked?`<button class="primary-btn" data-action="quiz-next">${q.index===q.questions.length-1?'Lihat hasil →':'Soal berikutnya →'}</button>`:`<button class="primary-btn" data-action="quiz-check" ${chosen==null?'disabled':''}>Periksa jawaban</button>`}</div></section>${side}</div>`;
}
function renderSettings(){return `${headerBlock('PREFERENSI','Pengaturan & data belajar','Kelola tema aplikasi dan cadangan progres. Semua data disimpan di browser perangkat ini, bukan pada server.')}
<div class="settings-grid"><div class="settings-card"><h3>▦ Jalur unit kerja</h3><p>Pilih bidang pelatihan Anda. Pilihan ini mengatur studi kasus yang direkomendasikan, tanpa mengubah materi umum yang tersedia.</p><select class="filter-select" style="width:100%" id="learner-role">${CAMPUS_CASES.map(c=>`<option value="${c.id}" ${state.role===c.id?'selected':''}>${esc(c.name)}</option>`).join('')}</select><button class="text-link" data-action="navigate" data-section="cases" style="margin-top:12px">Buka jalur praktik →</button></div><div class="settings-card"><h3>👤 Identitas peserta (opsional)</h3><p>Nama ini digunakan pada Bukti Belajar Mandiri setelah menyelesaikan seluruh modul, latihan mikro, dan tes satu level. Tidak dikirim ke server.</p><input id="learner-name" class="filter-input" type="text" maxlength="80" value="${esc(state.learnerName)}" placeholder="Nama peserta (opsional)" aria-label="Nama peserta"><p style="font-size:11px;margin-top:10px">Bukti belajar mandiri bukan sertifikasi resmi.</p></div><div class="settings-card"><h3>☾ Preferensi tampilan</h3><p>Ganti mode terang dan gelap. Pilihan Anda tersimpan saat browser ditutup.</p><button class="outline-btn" data-action="toggle-theme">${state.theme==='light'?'Aktifkan mode gelap':'Aktifkan mode terang'}</button></div><div class="settings-card"><h3>⇩ Cadangan progres</h3><p>Unduh progres materi, latihan, kuis, serta konten latihan sebagai berkas JSON. Mendukung impor cadangan OfficeLab v1, v2, dan v2.1. Nilai lama tidak dianggap kelulusan kurikulum baru.</p><div style="display:flex;gap:8px;flex-wrap:wrap"><button class="primary-btn" data-action="progress-export">↓ Ekspor progres</button><button class="outline-btn" data-action="progress-import">↑ Impor progres</button><input id="import-progress" type="file" accept=".json,application/json" style="display:none" aria-label="Pilih cadangan OfficeLab"></div></div><div class="settings-card"><h3>↺ Atur ulang data</h3><p>Hapus progres kuis, tugas, data worksheet, dokumen, dan slide latihan di browser ini. Tidak dapat dibatalkan.</p><button class="danger-btn" data-action="progress-reset">Reset semua progres</button></div><div class="settings-card"><h3>ⓘ Tentang OfficeLab</h3><p>Webapp edukasi mandiri untuk Excel, Word dan PowerPoint. Versi 2.1.2: 75 modul terstruktur dan 7 jalur studi kasus administrasi kampus. Tidak memerlukan login, server, atau internet.</p><button class="text-link" data-action="navigate" data-section="about">Buka halaman About →</button><div class="small-meta">Konten berfokus pada Microsoft Office desktop untuk Windows dan fitur umum. Simulasi memiliki keterbatasan yang berbeda dari perangkat lunak aslinya.</div></div></div>
<div class="section-heading"><div><h2>Ringkasan progres</h2><p>Persentase menggabungkan modul, latihan mikro, praktik, evaluasi level, dan kuis umum ≥80%.</p></div></div><div class="stat-grid">${Object.keys(OFFICE).map(id=>`<div class="stat-card"><div><small>${OFFICE[id].name}</small><strong>${progressFor(id)}%</strong></div><span class="app-icon ${id}">${OFFICE[id].short}</span></div>`).join('')}</div>`;}
function exportCertificate(id,level){
 const st=levelStats(id,level);
 if(!st.complete){toast('Bukti tuntas tersedia setelah 5 materi, 5 latihan mikro, dan tes level ≥80%.');return;}
 const name=(state.learnerName||'Peserta OfficeLab').trim().slice(0,80),course=OFFICE[id].name,
       date=new Date().toLocaleDateString('id-ID',{day:'numeric',month:'long',year:'numeric'}),
       lessons=OFFICE[id].lessons.filter(l=>l.level===level).map(l=>`<li>${esc(l.title)}</li>`).join('');
 const html=`<!doctype html><html lang="id"><head><meta charset="utf-8"><title>Bukti Belajar Mandiri · ${esc(course)} ${esc(level)}</title><style>@page{size:A4 landscape;margin:14mm}*{box-sizing:border-box}body{margin:0;padding:30px;background:#f2f5fb;font-family:Segoe UI,Arial,sans-serif;color:#1b2a41}.page{max-width:1000px;min-height:620px;margin:auto;background:#fff;border:14px solid #172b46;outline:3px solid #b7c7da;outline-offset:-25px;padding:65px 70px;text-align:center}.mark{font-weight:900;letter-spacing:.15em;color:#13996c;font-size:15px}.sub{color:#71829a;letter-spacing:.2em;text-transform:uppercase;font-size:11px;margin:16px 0 26px}h1{font-size:38px;margin:0 0 20px;color:#1c2e47}.name{font-size:32px;font-weight:800;line-height:1.4;margin:20px 0 10px;border-bottom:1px solid #d3dce9;padding-bottom:13px}.details{font-size:20px;font-weight:700;color:#0c845d;margin:18px 0}.metrics{display:flex;justify-content:center;gap:25px;margin:28px auto;font-size:14px}.metrics span{padding:12px 20px;background:#eff8f3;border-radius:10px}.module-list{font-size:10px;display:grid;grid-template-columns:1fr 1fr;text-align:left;gap:7px 18px;list-style:circle;padding:0 45px;max-width:850px;margin:8px auto 28px;color:#4c5a6c}.footer{font-size:11px;color:#68798c;line-height:1.7}button{background:#13996c;color:#fff;border:0;border-radius:8px;padding:12px 18px;cursor:pointer;margin:20px auto;display:block}@media print{body{padding:0;background:#fff}.page{border:11px solid #172b46;outline:none;min-height:auto}button{display:none}}</style></head><body><div class="page"><div class="mark">OFFICELAB V2.1</div><div class="sub">Bukti penyelesaian belajar mandiri · tidak terakreditasi</div><h1>Apresiasi Pembelajaran</h1><p>Diberikan kepada</p><div class="name">${esc(name)}</div><p>yang telah menyelesaikan aktivitas pembelajaran mandiri</p><div class="details">Microsoft ${esc(course)} — Level ${esc(level)}</div><div class="metrics"><span>5 Materi Tuntas</span><span>5 Latihan Mikro Lulus</span><span>Skor Tes Terbaik ${st.best}%</span></div><ul class="module-list">${lessons}</ul><div class="footer">Tanggal cetak: ${esc(date)}<br>Bukti dibuat secara lokal oleh perangkat pengguna; tidak ada verifikasi independen atau penerbit sertifikat eksternal.</div></div><button onclick="window.print()">Cetak / Simpan sebagai PDF</button></body></html>`;
 download(`OfficeLab_v21_Bukti_Belajar_${id}_${level}.html`,html,'text/html;charset=utf-8');
}
function download(filename,content,mime){const blob=new Blob([content],{type:mime}),url=URL.createObjectURL(blob),a=document.createElement('a');a.href=url;a.download=filename;document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),5000);toast(`Berkas ${filename} disiapkan.`);}
function exportProgress(){const payload={app:'OfficeLab',version:2.1,exportedAt:new Date().toISOString(),data:state};download('OfficeLab_v21_Cadangan_Progres.json',JSON.stringify(payload,null,2),'application/json;charset=utf-8');}
async function importProgress(file){
 if(!file)return;if(file.size>1024000){toast('Berkas terlalu besar (maksimum 1 MB).');return;}
 try{
   const payload=JSON.parse(await file.text());
   if(payload.app!=='OfficeLab'||![1,2,2.1].includes(payload.version)||!payload.data||typeof payload.data!=='object')throw Error('Berkas cadangan bukan OfficeLab v1/v2/v2.1.');
   if(!payload.data.completedLessons||!payload.data.quizBest||!Array.isArray(payload.data.pptSlides))throw Error('Struktur cadangan belum lengkap.');
   if(!confirm('Impor cadangan akan mengganti progres saat ini. Lanjutkan?'))return;
   const imported=prepareState(payload.data,payload.version!==2.1);
   imported.wordHTML=cleanWordHTML(imported.wordHTML);
   state=imported;ui.quiz=null;ui.lessonChoice=null;ui.lessonFeedback=null;
   save();render();toast(`Cadangan OfficeLab v${payload.version} berhasil diimpor.`);
 }catch(err){toast('Impor gagal: '+(err.message||'Format berkas tidak valid.'));}
}
function buildSearchIndex(){const results=[];for(const [id,course] of Object.entries(OFFICE)){
  course.lessons.forEach((item,i)=>results.push({course:id,section:'lessons',index:i,type:'Materi',title:item.title,desc:item.lead,search:[item.title,item.lead,...item.points,...(item.steps||[]),item.exercise,item.output,(item.roles||[]).join(' ')].join(' ')}));
  course.shortcuts.forEach((s,i)=>results.push({course:id,section:'shortcuts',index:i,type:'Shortcut',title:s[0],desc:s[1],search:s.join(' ')}));
  if(course.formulas)course.formulas.forEach((s,i)=>results.push({course:id,section:'formulas',index:i,type:'Rumus',title:s.name,desc:s.desc,search:[s.name,s.desc,s.syntax,s.example].join(' ')}));
}return results;}
const searchIndex=buildSearchIndex();
function showSearch(){const modal=$('#search-modal');modal.classList.remove('hidden');const input=$('#global-search');input.value='';fillSearchResults('');input.focus();}
function hideSearch(){$('#search-modal').classList.add('hidden');$('#open-search').focus();}
function fillSearchResults(query){const q=(query||'').trim().toLowerCase();let list;if(!q){list=searchIndex.filter(x=>/Keuangan|Surat|PivotTable|Rapat|Anggaran|Ctrl \+ S/i.test(x.title+' '+x.desc)).slice(0,9);}else{list=searchIndex.map(x=>({item:x,score:(x.title.toLowerCase()===q?10:0)+(x.title.toLowerCase().includes(q)?5:0)+(x.desc.toLowerCase().includes(q)?2:0)+(x.search.toLowerCase().includes(q)?1:0)})).filter(x=>x.score).sort((a,b)=>b.score-a.score).slice(0,30).map(x=>x.item);}
$('#global-search-results').innerHTML=(list.length?list.map(x=>`<button class="search-result" data-action="search-pick" data-course="${x.course}" data-section="${x.section}" data-index="${x.index}"><span class="search-result-icon">${x.type==='Rumus'?'∑':x.type==='Shortcut'?'⌨':'▤'}</span><span><b>${esc(x.title)} <span class="course-pill">${esc(OFFICE[x.course].name)}</span></b><small>${esc(x.type)} · ${esc(x.desc.slice(0,130))}</small></span></button>`).join(''):'<p class="empty-state">Belum ditemukan. Coba istilah lain, misalnya SUM, Ctrl, Heading, atau slide.</p>');}
function onSearchPick(node){const course=node.dataset.course,section=node.dataset.section,index=Number(node.dataset.index);$('#search-modal').classList.add('hidden');if(section==='lessons')ui.lessonIndex[course]=index;else if(section==='shortcuts'){ui.shortcutCategory='Semua';ui.shortcutSearch=OFFICE[course].shortcuts[index][0];}else if(section==='formulas'){ui.formulaCategory='Semua';ui.formulaSearch=OFFICE[course].formulas[index].name;}go(course,section);}
// Tombol dan kontrol yang dibangkitkan dinamis menggunakan event delegation.
document.addEventListener('click',event=>{
  const t=event.target.closest('[data-action]');if(!t)return;const action=t.dataset.action;
  if(action==='case-select'){if(CASE_BY_ID[t.dataset.case]){state.role=t.dataset.case;save();go(state.course,'cases');}return;}
  if(action==='case-download-csv'){caseCsv(t.dataset.case||state.role);return;}
  if(action==='case-sheet'){caseSheet();return;}
  if(action==='case-reset'){if(confirm('Reset checklist studi kasus unit '+CASE_BY_ID[state.role].name+'?')){state.caseChecks[state.role]=[];save();render();}return;}
  if(action==='case-load-excel'){
    if(!confirm('Muat dataset '+CASE_BY_ID[state.role].name+' ke simulator Excel? Data worksheet latihan saat ini akan diganti.'))return;
    state.excelCaseId=state.role;state.excelCells=clone(EXCEL_CASES[state.role].cells);state.excelDone=[];ui.activeCell='E8';ui.activeTask=0;save();go('excel','practice');toast('Dataset latihan unit dimuat.');return;
  }
  if(action==='case-load-word'){
    if(!confirm('Muat template dokumen baru? Isi editor Word yang belum diekspor akan diganti.'))return;
    state.wordHTML=WORD_TEMPLATES[state.role].html;state.wordDone=[];ui.wordRange=null;save();go('word','practice');toast('Template Word unit dimuat.');return;
  }
  if(action==='case-load-ppt'){
    if(!confirm('Muat empat slide studi kasus? Deck PowerPoint latihan saat ini akan diganti.'))return;
    state.pptSlides=clone(PPT_TEMPLATES[state.role].slides);state.pptDone=[];state.pptPresented=false;ui.pptIndex=0;save();go('powerpoint','practice');toast('Template presentasi unit dimuat.');return;
  }
  if(action==='navigate'){go(t.dataset.course||state.course,t.dataset.section||'home');return;}
  if(action==='level-select'){chooseLevel(t.dataset.course||state.course,t.dataset.level);return;}
  if(action==='level-exam'){ui.quiz=null;state.selectedLevel[state.course]=t.dataset.level;save();go(state.course,'quiz');startQuiz('level',t.dataset.level);return;}
  if(action==='certificate-download'){exportCertificate(state.course,t.dataset.level);return;}
  if(action==='lesson-open'){ui.lessonIndex[state.course]=Number(t.dataset.index);state.selectedLevel[state.course]=OFFICE[state.course].lessons[ui.lessonIndex[state.course]].level;ui.lessonChoice=null;ui.lessonFeedback=null;go(state.course,'lessons');return;}
  if(action==='lesson-select'){const i=Number(t.dataset.index);if(i>=0&&i<OFFICE[state.course].lessons.length){ui.lessonIndex[state.course]=i;state.selectedLevel[state.course]=OFFICE[state.course].lessons[i].level;ui.lessonChoice=null;ui.lessonFeedback=null;save();render();}return;}
  if(action==='lesson-complete'){const i=Number(t.dataset.index),list=state.completedLessons[state.course],wasDone=list.includes(i);if(wasDone)list.splice(list.indexOf(i),1);else list.push(i);if(!wasDone&&i%5!==4){ui.lessonIndex[state.course]=i+1;ui.lessonChoice=null;ui.lessonFeedback=null;}save();render();toast(wasDone?'Tanda selesai dihapus.':'Materi ditandai selesai.');return;}
  if(action==='lesson-answer'){ui.lessonChoice=Number(t.dataset.index);ui.lessonFeedback=null;render();return;}
  if(action==='lesson-check'){const id=state.course,i=ui.lessonIndex[id],l=OFFICE[id].lessons[i];if(ui.lessonChoice==null)return;const correct=ui.lessonChoice===l.check.answer;ui.lessonFeedback={index:i,correct};if(correct&&!state.passedChecks[id].includes(i))state.passedChecks[id].push(i);save();render();toast(correct?'Jawaban benar! Latihan mikro dicatat.':'Jawaban belum tepat. Pelajari pembahasan dan coba lagi.');return;}
  if(action==='flash-flip'){ui.flashFlip=!ui.flashFlip;renderFlashcard();return;}
  if(action==='flash-next'||action==='flash-prev'){ui.flashIndex+=action==='flash-next'?1:-1;ui.flashFlip=false;renderFlashcard();return;}
  if(action==='try-formula'){const f=OFFICE.excel.formulas.find(x=>x.name===t.dataset.name);if(!f)return;go('excel','practice');const example=formulaForCase(f);ui.activeCell=example.target;selectExcelCell(example.target,false);$('#formula-bar').value=example.example;$('#formula-bar').focus();$('#formula-bar').select();toast(`Contoh ${f.name} dimasukkan ke bilah rumus ${example.target}. Tekan Enter/Terapkan untuk menghitung.`);return;}
  if(action==='excel-select-task'){ui.activeTask=Number(t.dataset.index);ui.activeCell=EXCEL_TASKS[ui.activeTask].cell;ui.practiceMessage=null;render();selectExcelCell(ui.activeCell,true);return;}
  if(action==='excel-hint'){ui.activeTask=Number(t.dataset.index);ui.activeCell=EXCEL_TASKS[ui.activeTask].cell;ui.practiceMessage={type:'',text:`Petunjuk ${EXCEL_TASKS[ui.activeTask].cell}: ${taskHint(EXCEL_TASKS[ui.activeTask])} — ketik rumus ini atau variasi yang setara.`};render();return;}
  if(action==='excel-commit'){excelCommit();return;}
  if(action==='excel-check'){checkTask(ui.activeTask);return;}
  if(action==='excel-check-all'){checkTask(0,true);return;}
  if(action==='excel-export'){exportExcelCSV();return;}
  if(action==='excel-reset'){if(!confirm('Reset seluruh data dan progres tantangan worksheet Excel?'))return;state.excelCells=clone(EXCEL_CASES[state.excelCaseId].cells);state.excelDone=[];ui.activeCell='E8';ui.activeTask=0;ui.practiceMessage=null;save();render();toast('Worksheet unit aktif direset.');return;}
  if(action==='word-command'){wordCommand(t.dataset.cmd);return;}
  if(action==='word-table'){insertWordTable();return;}
  if(action==='word-export'){exportWord();return;}
  if(action==='word-print'){printWord();return;}
  if(action==='word-reset'){if(!confirm('Reset dokumen dan progres praktik Word?'))return;state.wordHTML=WORD_INITIAL;state.wordDone=[];ui.wordRange=null;save();render();return;}
  if(action==='ppt-select'){ui.pptIndex=Number(t.dataset.index);render();return;}
  if(['ppt-add','ppt-duplicate','ppt-delete','ppt-up','ppt-down'].includes(action)){changePpt(action);return;}
  if(action==='ppt-theme'){state.pptTheme=t.dataset.theme;save();render();return;}
  if(action==='ppt-present'){startPresentation();return;}
  if(action==='ppt-export'){exportPpt();return;}
  if(action==='ppt-reset'){if(!confirm('Reset semua slide dan progres latihan PowerPoint?'))return;state.pptSlides=clone(PPT_INITIAL);state.pptTheme='blue';state.pptDone=[];state.pptPresented=false;ui.pptIndex=0;save();render();return;}
  if(action==='quiz-home'){ui.quiz=null;render();return;}
  if(action==='quiz-start'){startQuiz(t.dataset.mode||'level',t.dataset.level||state.selectedLevel[state.course]);return;}
  if(action==='quiz-select'){const q=quizCurrent();if(!q||q.checked)return;q.selected=Number(t.dataset.index);render();return;}
  if(action==='quiz-check'){const q=quizCurrent();if(!q||q.checked||q.selected==null)return;q.checked=true;if(q.selected===q.questions[q.index][2]){q.correct++;if(q.mode==='placement'){const lev=q.levels[q.index];q.correctByLevel[lev]=(q.correctByLevel[lev]||0)+1;}}render();return;}
  if(action==='quiz-next'){const q=quizCurrent();if(!q||!q.checked)return;if(q.index===q.questions.length-1)finishQuiz();else{q.index++;q.selected=null;q.checked=false;render();}return;}
  if(action==='toggle-theme'){state.theme=state.theme==='light'?'dark':'light';save();render();return;}
  if(action==='progress-export'){exportProgress();return;}
  if(action==='progress-import'){$('#import-progress')?.click();return;}
  if(action==='progress-reset'){if(!confirm('Hapus SELURUH progres dan latihan OfficeLab pada browser ini?'))return;state=initialState();ui.quiz=null;ui.lessonChoice=null;ui.lessonFeedback=null;ui.activeCell='E8';ui.activeTask=0;ui.pptIndex=0;save();go('excel','home');toast('Seluruh progres sudah direset.');return;}
  if(action==='search-pick'){onSearchPick(t);return;}
});
document.addEventListener('input',event=>{const t=event.target;
  if(t.id==='learner-name'){state.learnerName=t.value.slice(0,80);save();}
  else if(t.id==='shortcut-filter'){ui.shortcutSearch=t.value;updateShortcutRows();}
  else if(t.id==='formula-filter'){ui.formulaSearch=t.value;updateFormulaCards();}
  else if(t.id==='global-search')fillSearchResults(t.value);
  else if(t.id==='word-editor'){state.wordHTML=cleanWordHTML(t.innerHTML);save();refreshWordProgress();}
  else if(t.matches('.sheet-input')&&document.activeElement===t){const bar=$('#formula-bar');if(bar)bar.value=t.value;}
  else if(['ppt-title','ppt-body','ppt-notes'].includes(t.id)){const key=t.id.replace('ppt-','');state.pptSlides[ui.pptIndex][key]=t.value;save();pptLiveUpdate();}
});
document.addEventListener('change',event=>{const t=event.target;
  if(t.matches('input[data-action="case-task"]')){const id=t.dataset.case,key=t.dataset.task;if(CASE_BY_ID[id]?.tasks.some(x=>x.id===key)){const arr=state.caseChecks[id]||[];state.caseChecks[id]=t.checked?[...new Set([...arr,key])]:arr.filter(x=>x!==key);save();render();}return;}
  if(t.id==='excel-case-select'){const id=t.value;if(!EXCEL_CASES[id]||id===state.excelCaseId)return;if(!confirm('Ganti dataset simulator Excel? Isian dan hasil 14 tantangan saat ini akan direset.')){t.value=state.excelCaseId;return;}state.excelCaseId=id;state.excelCells=clone(EXCEL_CASES[id].cells);state.excelDone=[];state.role=id;ui.activeCell='E8';ui.activeTask=0;save();render();toast('Dataset '+CASE_BY_ID[id].name+' dimuat.');return;}
  if(t.id==='learner-role'){if(CASE_BY_ID[t.value]){state.role=t.value;save();render();}return;}
  if(t.id==='shortcut-category'){ui.shortcutCategory=t.value;updateShortcutRows();}
  else if(t.id==='formula-category'){ui.formulaCategory=t.value;updateFormulaCards();}
  else if(t.matches('.sheet-input')){setExcelCell(t.dataset.cell,t.value);}
  else if(t.id==='formula-bar'){setExcelCell(ui.activeCell,t.value);}
  else if(t.id==='word-style'){restoreWordSelection();try{document.execCommand('formatBlock',false,t.value);}catch(_){toast('Gaya teks tidak tersedia di browser ini.');}state.wordHTML=cleanWordHTML($('#word-editor').innerHTML);save();refreshWordProgress();}
  else if(t.id==='ppt-layout'){state.pptSlides[ui.pptIndex].layout=t.value;save();render();}
  else if(t.id==='import-progress'){importProgress(t.files?.[0]);t.value='';}
});
document.addEventListener('focusin',event=>{const t=event.target;if(t.matches('.sheet-input'))selectExcelCell(t.dataset.cell,false);});
document.addEventListener('selectionchange',()=>{const editor=$('#word-editor'),selection=window.getSelection();if(!editor||!selection?.rangeCount||!selection.anchorNode||!editor.contains(selection.anchorNode))return;try{ui.wordRange=selection.getRangeAt(0).cloneRange();}catch(_){}});
document.addEventListener('mousedown',event=>{const t=event.target.closest('[data-action="word-command"],[data-action="word-table"]');if(t)event.preventDefault();});
document.addEventListener('paste',event=>{const editor=event.target.closest?.('#word-editor');if(!editor)return;event.preventDefault();const text=event.clipboardData?.getData('text/plain')||'';document.execCommand('insertText',false,text);});
document.addEventListener('keydown',event=>{
  if((event.ctrlKey||event.metaKey)&&event.key.toLowerCase()==='k'){event.preventDefault();showSearch();return;}
  if(event.key==='Escape'){if(!$('#search-modal').classList.contains('hidden')){event.preventDefault();hideSearch();return;}if(!$('#presentation-overlay').classList.contains('hidden')){event.preventDefault();closePresentation();return;}}
  if(!$('#presentation-overlay').classList.contains('hidden')){if(['ArrowRight','PageDown',' ','Enter'].includes(event.key)){event.preventDefault();ui.presentationIndex++;refreshPresentation();}else if(['ArrowLeft','PageUp','Backspace'].includes(event.key)){event.preventDefault();ui.presentationIndex--;refreshPresentation();}return;}
  if(event.key==='Enter'&&event.target.id==='formula-bar'){event.preventDefault();excelCommit();}
  if(event.key==='Enter'&&event.target.matches('.sheet-input')){event.preventDefault();const addr=event.target.dataset.cell,part=addressParts(addr),next=`${colName(part.col)}${Math.min(9,part.row+1)}`;event.target.blur();setTimeout(()=>selectExcelCell(next,true),0);}
});
$('#open-about').addEventListener('click',()=>go(state.course,'about'));
$('#theme-toggle').addEventListener('click',()=>{state.theme=state.theme==='light'?'dark':'light';save();render();});
$('#open-search').addEventListener('click',showSearch);
$('#close-search').addEventListener('click',hideSearch);
$('#search-modal').addEventListener('click',event=>{if(event.target===$('#search-modal'))hideSearch();});
$('#presentation-prev').addEventListener('click',()=>{ui.presentationIndex--;refreshPresentation();});
$('#presentation-next').addEventListener('click',()=>{ui.presentationIndex++;refreshPresentation();});
$('#presentation-close').addEventListener('click',closePresentation);
// Sanitasi dokumen dari data browser sebelum pertama kali dirender.
state.wordHTML=cleanWordHTML(state.wordHTML);
render();
})();
