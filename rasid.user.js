// ==UserScript==
// @name         راصد العبارات الوصفية - البوابة التعليمية
// @namespace    mulawah-school
// @author       سيف جمال المعولي
// @updateURL    https://mawals.github.io/estebyan-grade/rasid.user.js
// @downloadURL  https://mawals.github.io/estebyan-grade/rasid.user.js
// @version      3.0
// @description  ينقل العبارات الوصفية من ملف الإكسل إلى صفحة «إدخال الدرجات» في البوابة التعليمية بمطابقة أسماء الطلاب تلقائيًا، مع تقرير بالطلاب غير المطابقين
// @match        *://*.moe.gov.om/*
// @include      /^https?:\/\/([^\/]*\.)?moe\.gov\.om(:\d+)?\/.*$/
// @require      https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js
// @grant        GM_setValue
// @grant        GM_getValue
// @grant        GM_deleteValue
// @grant        GM_notification
// @run-at       document-start
// ==/UserScript==

(function () {
  'use strict';
  if (window.top !== window.self) return;

  // أثناء الرصد التلقائي: نمنع رسائل البوابة (تم الحفظ...) من إيقاف العمل، وتُحفظ الرسالة للتقرير
  try {
    const a = typeof GM_getValue === 'function' ? GM_getValue('rz_auto', null) : JSON.parse(localStorage.getItem('rz_auto') || 'null');
    if (a) {
      const sc = document.createElement('script');
      sc.textContent = "if(!window.__rzA){window.__rzA=[window.alert,window.confirm];window.alert=function(m){document.documentElement.setAttribute('data-rz-msg',String(m))};window.confirm=function(m){document.documentElement.setAttribute('data-rz-msg',String(m));return true};}";
      const put = () => { const r = document.head || document.documentElement; if (!r) return false; r.appendChild(sc); sc.remove(); return true; };
      if (!put()) { const mo = new MutationObserver(() => { if (put()) mo.disconnect(); }); mo.observe(document, { childList: true, subtree: true }); }
    }
  } catch (e) {}
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', () => setTimeout(main, 300));
  else main();

  function main() {

  // يعمل فقط في صفحة إدخال الدرجات
  const isMarksPage = () => /MarksDataEntry/i.test(location.href) ||
    /إدخال الدرجات|وصف أداء الطالب/.test((document.body && document.body.innerText) || '');
  if (!isMarksPage() || document.getElementById('rz-panel')) return;

  // ================= التخزين =================
  const hasGM = typeof GM_getValue === 'function';
  const get = (k, d) => {
    try { if (hasGM) return GM_getValue(k, d); const v = localStorage.getItem(k); return v == null ? d : JSON.parse(v); } catch (e) { return d; }
  };
  const set = (k, v) => {
    try { if (hasGM) GM_setValue(k, v); else localStorage.setItem(k, JSON.stringify(v)); return true; } catch (e) { return false; }
  };
  const FILESK = 'rz_files', OPTK = 'rz_opt', THEMEK = 'rz_theme', LASTK = 'rz_last';
  let store = get(FILESK, { files: {}, current: '' }) || { files: {}, current: '' };
  let opt = Object.assign({ noOverwrite: false, autoSave: false, fuzzy: true }, get(OPTK, {}) || {});
  let wb = store.files[store.current] || null;
  const saveOpt = () => set(OPTK, opt);
  const saveStore = () => set(FILESK, store);

  const sleep = ms => new Promise(r => setTimeout(r, ms));
  const esc = v => String(v ?? '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const today = () => new Date().toISOString().slice(0, 10);
  const VER = (typeof GM_info !== 'undefined' && GM_info.script && GM_info.script.version) || '';

  // ================= المطابقة =================
  function norm(s) {
    return String(s || '')
      .replace(/[ً-ٰٟـ]/g, '')
      .replace(/[أإآٱ]/g, 'ا').replace(/ة/g, 'ه').replace(/ى/g, 'ي').replace(/ؤ/g, 'و').replace(/ئ/g, 'ي')
      .replace(/[^ء-يa-zA-Z0-9\s]/g, ' ')
      .replace(/(^|\s)عبد\s+/g, '$1عبد')
      .replace(/(^|\s)(بن|بنت|ابن)(?=\s|$)/g, ' ')
      .replace(/\s+/g, ' ').trim();
  }
  const tokens = s => norm(s).split(' ').filter(Boolean);
  function score(a, b) {
    const A = tokens(a), B = tokens(b);
    if (!A.length || !B.length || A[0] !== B[0]) return 0;
    const setB = new Set(B); let c = 0;
    A.forEach(t => { if (setB.has(t)) c++; });
    return c / Math.max(A.length, B.length);
  }
  const hasArabic = v => /[ء-ي]/.test(String(v || ''));

  // ================= قراءة الإكسل =================
  function colLetter(i) { let s = ''; i++; while (i > 0) { const m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = Math.floor((i - 1) / 26); } return s; }
  function analyseSheet(rows) {
    let headerRow = -1, nameCol = -1, phraseCol = -1;
    for (let r = 0; r < Math.min(rows.length, 20); r++) {
      (rows[r] || []).forEach((v, c) => {
        const t = String(v || '');
        if (nameCol < 0 && /student\s*name|اسم\s*الطالب|الاسم/i.test(t)) { nameCol = c; headerRow = r; }
        if (phraseCol < 0 && /عبار|وصف/.test(t)) phraseCol = c;
      });
    }
    if (nameCol < 0) {
      const cnt = {};
      rows.forEach(row => (row || []).forEach((v, c) => { if (hasArabic(v) && String(v).trim().split(/\s+/).length >= 3 && String(v).length < 60) cnt[c] = (cnt[c] || 0) + 1; }));
      nameCol = +Object.keys(cnt).sort((a, b) => cnt[b] - cnt[a])[0] || 0; headerRow = 0;
    }
    if (phraseCol < 0) {
      const cnt = {};
      rows.forEach(row => (row || []).forEach((v, c) => { if (c !== nameCol && hasArabic(v) && String(v).length > 15) cnt[c] = (cnt[c] || 0) + 1; }));
      const k = Object.keys(cnt).sort((a, b) => cnt[b] - cnt[a])[0];
      phraseCol = k === undefined ? -1 : +k;
    }
    return { headerRow, nameCol, phraseCol };
  }
  function students(sh, nameCol, phraseCol) {
    const out = [];
    sh.rows.forEach((row, r) => {
      if (r <= sh.info.headerRow) return;
      const name = String((row || [])[nameCol] || '').trim();
      if (!hasArabic(name) || tokens(name).length < 2) return;
      out.push({ name, phrase: phraseCol >= 0 ? String((row || [])[phraseCol] || '').trim() : '', row: r + 1 });
    });
    return out;
  }

  // ================= قراءة صفحة البوابة =================
  function pageStudents() {
    const re = /^\s*(.+?)\s*-\s*(\d{6,})\s*$/;
    const anchors = [];
    const w = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT);
    let n;
    while ((n = w.nextNode())) {
      if (!n.parentElement || n.parentElement.closest('#rz-panel,#rz-done')) continue;
      const m = (n.nodeValue || '').match(re);
      if (m && hasArabic(m[1])) anchors.push({ node: n.parentElement, name: m[1].trim(), id: m[2] });
    }
    const tas = [...document.querySelectorAll('textarea')].filter(t => !t.closest('#rz-panel'));
    const before = (a, b) => !!(a.compareDocumentPosition(b) & Node.DOCUMENT_POSITION_FOLLOWING);
    anchors.forEach((a, i) => {
      const next = anchors[i + 1];
      a.textarea = tas.find(t => before(a.node, t) && (!next || before(t, next.node)) && !t.disabled && !t.readOnly) || null;
    });
    return anchors;
  }
  function findSaveBtn() {
    return [...document.querySelectorAll('input[type=submit],input[type=button],button,a')]
      .find(el => !el.closest('#rz-panel,#rz-done') && /^\s*حفظ\s*$/.test(el.value || el.innerText || '')) || null;
  }
  function guessSheet(names) {
    const grades = { 'الخامس': 5, 'السادس': 6, 'السابع': 7, 'الثامن': 8, 'التاسع': 9, 'العاشر': 10, 'الحادي عشر': 11, 'الثاني عشر': 12 };
    let g = null, sh = null;
    document.querySelectorAll('select').forEach(s => {
      if (s.closest('#rz-panel')) return;
      const t = ((s.options[s.selectedIndex] || {}).text || '').trim();
      if (grades[t]) g = grades[t];
      const row = s.closest('tr');
      if (row && /الشعبة/.test(row.innerText) && /^\d+$/.test(t)) {
        const cells = [...(row.cells || [])];
        const li = cells.findIndex(c => /الشعبة/.test(c.innerText)), si = cells.findIndex(c => c.contains(s));
        if (li >= 0 && Math.abs(li - si) === 1) sh = t;
      }
    });
    if (g && sh) return names.find(n => n.replace(/\s/g, '') === g + '.' + sh) || null;
    return null;
  }

  // ================= السمات =================
  if (!document.getElementById('ax-cairo-font')) {
    const l = document.createElement('link'); l.id = 'ax-cairo-font'; l.rel = 'stylesheet';
    l.href = 'https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700;800;900&display=swap';
    document.head.appendChild(l);
  }
  const THEMES = {
    teal:    { name: 'فيروزي', dot: '#0e8a8a', v: '--bg:linear-gradient(160deg,#f0fdfc 0%,#c9eeec 100%);--solid:#f2fbfb;--box:#fbffff;--line:#9fd6d3;--ink:#063b3b;--muted:#3f7a78;--a1:#14b8a6;--a2:#0e7c7b;--aink:#ffffff;--hi:#0e7c7b;--glow:rgba(20,184,166,.3);--soft:rgba(20,184,166,.12);--btn:#dcf3f1;--stop:#be123c' },
    gold:    { name: 'ذهبي',   dot: '#d4a017', v: '--bg:linear-gradient(160deg,#fffbeb 0%,#fbecc0 100%);--solid:#fff8e1;--box:#fffdf6;--line:#ead08a;--ink:#3d2c05;--muted:#8b6b1f;--a1:#e8b923;--a2:#b8860b;--aink:#ffffff;--hi:#a16207;--glow:rgba(212,160,23,.35);--soft:rgba(232,185,35,.15);--btn:#fbefc8;--stop:#be123c' },
    royal:   { name: 'ملكي',   dot: '#1e40af', v: '--bg:linear-gradient(160deg,#1e3a8a 0%,#172554 100%);--solid:#1b2f6e;--box:rgba(255,255,255,.07);--line:rgba(255,255,255,.16);--ink:#f1f5ff;--muted:#b4c2ee;--a1:#fcd34d;--a2:#f59e0b;--aink:#1e1b4b;--hi:#fcd34d;--glow:rgba(252,211,77,.3);--soft:rgba(252,211,77,.12);--btn:rgba(255,255,255,.1);--stop:#fda4af' },
    violet:  { name: 'بنفسجي', dot: '#8b5cf6', v: '--bg:linear-gradient(160deg,#faf5ff 0%,#e9ddfd 100%);--solid:#fbf8ff;--box:#fdfcff;--line:#d6c4f6;--ink:#33105e;--muted:#7a5aa6;--a1:#a855f7;--a2:#7c3aed;--aink:#ffffff;--hi:#7c3aed;--glow:rgba(139,92,246,.3);--soft:rgba(168,85,247,.12);--btn:#efe6fd;--stop:#be123c' },
    dark:    { name: 'داكن',   dot: '#1c1917', v: '--bg:#1c1917;--solid:#1c1917;--box:#0c0a09;--line:#292524;--ink:#f5f5f4;--muted:#a8a29e;--a1:#2dd4bf;--a2:#0d9488;--aink:#0c0a09;--hi:#5eead4;--glow:rgba(45,212,191,.16);--soft:rgba(45,212,191,.1);--btn:#292524;--stop:#fda4af' },
  };
  const curTheme = () => (THEMES[get(THEMEK, 'teal')] ? get(THEMEK, 'teal') : 'teal');
  const thSt = document.createElement('style');
  thSt.textContent = Object.entries(THEMES).map(([k, t]) => `[data-rzt="${k}"]{${t.v}}`).join('\n') + `
    .rz-ok{outline:3px solid #22c55e !important;background:#effaf1 !important}
    .rz-fuzzy{outline:3px solid #f59e0b !important;background:#fff8e6 !important}
    .rz-miss{outline:3px solid #ef4444 !important}
    .rz-save-hi{outline:4px solid #f59e0b !important;box-shadow:0 0 0 8px rgba(245,158,11,.35) !important;animation:rzPulse 1s ease-in-out 4}
    @keyframes rzPulse{50%{box-shadow:0 0 0 14px rgba(245,158,11,.1)}}`;
  document.head.appendChild(thSt);

  // ================= اللوحة =================
  const panel = document.createElement('div');
  panel.id = 'rz-panel'; panel.dir = 'rtl';
  panel.innerHTML = `
    <style>
      #rz-panel{position:fixed;bottom:18px;left:18px;z-index:2147483646;background:var(--bg);color:var(--ink);border:1px solid var(--line);border-radius:22px;padding:15px;width:340px;max-height:calc(100vh - 36px);overflow-y:auto;font-family:'Cairo',Tahoma,Arial,sans-serif;font-size:12.5px;line-height:1.5;box-shadow:0 22px 45px -12px rgba(0,0,0,.35),0 0 0 1px var(--glow);box-sizing:border-box;transition:background .3s,color .3s}
      #rz-panel *{box-sizing:border-box;font-family:'Cairo',Tahoma,Arial,sans-serif}
      #rz-head{display:flex;align-items:center;justify-content:space-between;padding-bottom:11px;margin-bottom:11px;border-bottom:1px solid var(--line)}
      .rz-brand{display:flex;align-items:center;gap:9px}
      .rz-badge{position:relative;width:34px;height:34px;border-radius:11px;background:linear-gradient(135deg,var(--a1),var(--a2));color:var(--aink);font-weight:900;font-size:10px;display:flex;align-items:center;justify-content:center;box-shadow:0 4px 12px var(--glow)}
      .rz-badge::after{content:"";position:absolute;bottom:-2px;left:-2px;width:9px;height:9px;border-radius:50%;background:#10b981;border:2px solid var(--solid)}
      .rz-title{font-weight:900;font-size:13.5px;line-height:1.25}
      .rz-sub{font-size:10.5px;color:var(--muted);font-weight:600}
      .rz-tools{display:flex;align-items:center;gap:7px}
      .rz-dot{width:12px;height:12px;border-radius:50%;cursor:pointer;border:2px solid var(--solid);box-shadow:0 0 0 1px var(--line);display:inline-block;margin:0 1px;transition:transform .15s}
      .rz-dot:hover{transform:scale(1.2)} .rz-dot.on{box-shadow:0 0 0 2px var(--hi)}
      #rz-min{width:28px;height:28px;border-radius:9px;background:var(--btn);color:var(--muted);border:1px solid var(--line);display:flex;align-items:center;justify-content:center;cursor:pointer;font-weight:900;user-select:none}
      .rz-box{background:var(--box);border:1px solid var(--line);border-radius:14px;padding:9px 11px;margin-bottom:8px}
      .rz-lbl{font-weight:800;font-size:11.5px;color:var(--muted);margin:2px 0 4px;display:flex;justify-content:space-between}
      #rz-panel select{width:100%;padding:5px 8px;border:1px solid var(--line);border-radius:9px;background:var(--solid);color:var(--ink);font-size:12px;font-weight:700}
      .rz-row{display:flex;gap:6px;align-items:center}.rz-row>*{flex:1;min-width:0}
      #rz-panel button{display:flex;align-items:center;justify-content:center;gap:6px;width:100%;margin:5px 0;padding:9px 12px;border:1px solid var(--line);border-radius:12px;cursor:pointer;font-weight:800;font-size:12.5px;background:var(--btn);color:var(--ink);transition:all .2s}
      #rz-panel button:hover{border-color:var(--a1);color:var(--hi)}
      #rz-panel button:active{transform:scale(.98)}
      #rz-panel button:disabled{opacity:.5;cursor:not-allowed}
      #rz-panel button[data-a=go]{background:linear-gradient(135deg,var(--a1),var(--a2));color:var(--aink);border-color:transparent;font-size:13.5px;font-weight:900;box-shadow:0 8px 20px -4px var(--glow)}
      #rz-panel button[data-a=go]:hover{filter:brightness(1.08);transform:translateY(-1px)}
      #rz-panel button[data-a=del]{width:auto;flex:0 0 auto;margin:0;padding:5px 10px;background:rgba(225,29,72,.1);color:var(--stop);border-color:rgba(225,29,72,.3);font-size:11.5px}
      #rz-panel button[data-a=new]{margin:6px 0 0;padding:6px;font-size:11.5px;border-style:dashed}
      #rz-panel button.rz-all{background:var(--soft);border:1.5px solid var(--a1);color:var(--hi);font-weight:900}
      #rz-panel button.rz-stop{background:rgba(225,29,72,.1);color:var(--stop);border-color:rgba(225,29,72,.3)}
      .rz-grid{display:grid;grid-template-columns:1fr 1fr;gap:6px;margin:5px 0}
      .rz-grid button{margin:0 !important;padding:7.5px 8px !important;font-size:11.5px !important}
      #rz-panel button[data-a=diag]{background:transparent;color:var(--muted);border-style:dashed;padding:6px;font-size:11.5px}
      .rz-chk{display:flex;align-items:center;gap:7px;margin:3px 0;font-size:11.5px;font-weight:600;cursor:pointer}
      .rz-chk input{accent-color:var(--a2);width:14px;height:14px;margin:0}
      .rz-info{font-size:11px;color:var(--muted);margin-top:4px;font-weight:600}
      #rz-status{margin-top:8px;padding:9px 11px;background:var(--box);border:1px solid var(--line);border-right:3px solid var(--a1);border-radius:11px;white-space:pre-line;font-size:11.5px;font-weight:700;line-height:1.6}
      #rz-prog{display:none;height:6px;background:var(--btn);border-radius:99px;margin-top:8px;overflow:hidden}
      #rz-bar{height:100%;width:0;background:linear-gradient(90deg,var(--a1),var(--a2));transition:width .2s}
      #rz-list{margin-top:6px;max-height:190px;overflow:auto;font-size:11px}
      #rz-list div{padding:3px 6px;border-bottom:1px dashed var(--line);display:flex;gap:6px}
      #rz-list .k{flex:0 0 16px;text-align:center}
      .c-ok{color:#16a34a}.c-w{color:#d97706}.c-b{color:#dc2626}
      .rz-footer{margin-top:10px;padding-top:8px;border-top:1px solid var(--line);display:flex;align-items:center;gap:7px}
      .rz-ic{width:26px;height:26px;border-radius:8px;background:var(--soft);border:1px solid var(--line);color:var(--hi);display:flex;align-items:center;justify-content:center;font-size:14px}
      .rz-credit{font-size:10.5px;color:var(--muted);line-height:1.35}.rz-credit b{color:var(--ink);font-weight:800;display:block}
    </style>
    <div id="rz-head">
      <div class="rz-brand">
        <div class="rz-badge">راصد</div>
        <div><div class="rz-title">راصد العبارات الوصفية</div><div class="rz-sub">من الإكسل إلى البوابة • الإصدار ${esc(VER)}</div></div>
      </div>
      <div class="rz-tools">
        <span>${Object.entries(THEMES).map(([k, t]) => `<span class="rz-dot" data-t="${k}" title="${t.name}" style="background:${t.dot}"></span>`).join('')}</span>
        <span id="rz-min" title="تصغير / تكبير">—</span>
      </div>
    </div>
    <div id="rz-body">
      <div class="rz-box">
        <div class="rz-lbl"><span>📄 ملف الإكسل</span><span id="rz-fcount"></span></div>
        <div class="rz-row"><select id="rz-files"></select><button data-a="del" title="حذف الملف المختار">🗑 حذف</button></div>
        <button data-a="new">➕ رفع ملف إكسل جديد</button>
        <input type="file" id="rz-file" accept=".xlsx,.xls,.xlsm,.csv" style="display:none">
      </div>
      <div class="rz-box">
        <div class="rz-lbl"><span>📑 الورقة (الشعبة)</span><span id="rz-auto"></span></div>
        <select id="rz-sheet"></select>
        <div class="rz-row" style="margin-top:6px">
          <div><div class="rz-lbl">عمود الاسم</div><select id="rz-name"></select></div>
          <div><div class="rz-lbl">عمود العبارة</div><select id="rz-phrase"></select></div>
        </div>
        <div class="rz-info" id="rz-sinfo"></div>
      </div>
      <div class="rz-box">
        <label class="rz-chk"><input type="checkbox" id="rz-noover"> لا تكتب فوق عبارة موجودة في البوابة</label>
        <label class="rz-chk"><input type="checkbox" id="rz-fuzzy"> مطابقة تقريبية للأسماء المختلفة قليلًا</label>
        <label class="rz-chk"><input type="checkbox" id="rz-autosave"> الضغط على «حفظ» تلقائيًا بعد الرصد</label>
      </div>
      <button data-a="go">▶️ رصد العبارات (هذه الشعبة)</button>
      <button data-a="all" class="rz-all">🔁 رصد وحفظ كل الصفوف والشعب</button>
      <button data-a="stopall" class="rz-stop" style="display:none">⏹ إيقاف الرصد التلقائي</button>
      <div class="rz-grid">
        <button data-a="preview">👁️ معاينة المطابقة</button>
        <button data-a="report">📥 تقرير Excel</button>
        <button data-a="tosave">⬇️ الذهاب لزر الحفظ</button>
        <button data-a="clear">🧹 إزالة التظليل</button>
      </div>
      <button data-a="diag">🔍 فحص عناصر الصفحة</button>
      <div id="rz-status">اختر الصف والمادة والشعبة في البوابة ثم اضغط «عرض».</div>
      <div id="rz-prog"><div id="rz-bar"></div></div>
      <div id="rz-list"></div>
      <div class="rz-footer">
        <div class="rz-ic">✍️</div>
        <div class="rz-credit"><b>سيف جمال المعولي</b><span>معولة بن شمس (5-12)</span></div>
      </div>
    </div>`;
  document.body.appendChild(panel);
  const $ = id => panel.querySelector('#' + id);
  const status = t => ($('rz-status').textContent = t);

  function applyTheme(k) {
    panel.setAttribute('data-rzt', k);
    const d = document.getElementById('rz-done'); if (d) d.setAttribute('data-rzt', k);
    panel.querySelectorAll('.rz-dot').forEach(d => d.classList.toggle('on', d.dataset.t === k));
  }
  applyTheme(curTheme());
  panel.querySelectorAll('.rz-dot').forEach(d => d.onclick = () => { set(THEMEK, d.dataset.t); applyTheme(d.dataset.t); });
  $('rz-min').onclick = () => { const b = $('rz-body'), h = b.style.display === 'none'; b.style.display = h ? '' : 'none'; $('rz-min').textContent = h ? '—' : '+'; };
  [['rz-noover', 'noOverwrite'], ['rz-fuzzy', 'fuzzy'], ['rz-autosave', 'autoSave']].forEach(([id, k]) => {
    $(id).checked = !!opt[k]; $(id).onchange = () => { opt[k] = $(id).checked; saveOpt(); };
  });

  // ================= الملفات والأوراق =================
  function fillFiles() {
    const s = $('rz-files'); s.innerHTML = '';
    const names = Object.keys(store.files);
    if (!names.length) s.add(new Option('— لا يوجد ملف —', ''));
    names.forEach(n => s.add(new Option(n, n)));
    s.value = store.current || '';
    panel.querySelector('[data-a=del]').disabled = !names.length;
    $('rz-fcount').textContent = names.length ? names.length + ' محفوظ' : '';
  }
  function fillCols(sheetName) {
    const sh = wb.sheets[sheetName];
    const maxC = Math.max(1, ...sh.rows.slice(0, 60).map(r => (r || []).length));
    const label = c => {
      let h = '';
      for (let r = 0; r <= Math.min(sh.info.headerRow + 4, sh.rows.length - 1); r++) {
        const v = String((sh.rows[r] || [])[c] || '').trim();
        if (v && (hasArabic(v) || /name/i.test(v))) { h = v; break; }
      }
      return colLetter(c) + (h ? ' — ' + h.slice(0, 20) : '');
    };
    ['rz-name', 'rz-phrase'].forEach(id => {
      const s = $(id); s.innerHTML = '';
      if (id === 'rz-phrase') s.add(new Option('— اختر —', -1));
      for (let c = 0; c < maxC; c++) s.add(new Option(label(c), c));
    });
    $('rz-name').value = sh.info.nameCol;
    $('rz-phrase').value = sh.info.phraseCol;
    updateInfo();
  }
  function updateInfo() {
    if (!wb) { $('rz-sinfo').textContent = ''; return; }
    const sh = wb.sheets[$('rz-sheet').value]; if (!sh) return;
    const list = students(sh, +$('rz-name').value, +$('rz-phrase').value);
    const withP = list.filter(x => x.phrase).length;
    $('rz-sinfo').innerHTML = `👥 ${list.length} طالب — ✍️ <b>${withP}</b> لديهم عبارة${+$('rz-phrase').value < 0 ? ' — <span class="c-b">اختر عمود العبارة</span>' : ''}`;
  }
  function fillSheets() {
    fillFiles();
    const s = $('rz-sheet'); s.innerHTML = '';
    if (!wb) { $('rz-name').innerHTML = ''; $('rz-phrase').innerHTML = ''; $('rz-sinfo').textContent = ''; $('rz-auto').textContent = ''; return; }
    const names = Object.keys(wb.sheets);
    names.forEach(n => s.add(new Option(n, n)));
    const g = guessSheet(names);
    $('rz-auto').textContent = g ? '✨ اختيرت تلقائيًا' : '';
    s.value = g || (names.includes(wb.lastSheet) ? wb.lastSheet : names[0]);
    fillCols(s.value);
  }
  panel.querySelector('[data-a=new]').onclick = () => $('rz-file').click();
  $('rz-file').onchange = e => {
    const f = e.target.files[0]; if (!f) return;
    const rd = new FileReader();
    rd.onload = ev => {
      try {
        const book = XLSX.read(new Uint8Array(ev.target.result), { type: 'array' });
        wb = { fileName: f.name, added: today(), sheets: {} };
        book.SheetNames.forEach(n => {
          const rows = XLSX.utils.sheet_to_json(book.Sheets[n], { header: 1, defval: '', raw: false });
          wb.sheets[n] = { rows, info: analyseSheet(rows) };
        });
        store.files[f.name] = wb; store.current = f.name;
        if (!saveStore()) status('⚠️ الملف يعمل الآن لكن لم يُحفظ — احذف ملفات قديمة.');
        else status(`✅ تم رفع «${f.name}» (${book.SheetNames.length} ورقة).`);
        fillSheets();
      } catch (err) { status('❌ تعذّرت قراءة الملف: ' + err.message); }
      e.target.value = '';
    };
    rd.readAsArrayBuffer(f);
  };
  $('rz-files').onchange = () => { store.current = $('rz-files').value; wb = store.files[store.current] || null; saveStore(); fillSheets(); };
  panel.querySelector('[data-a=del]').onclick = () => {
    const n = $('rz-files').value;
    if (!n || !store.files[n]) return;
    if (!confirm('حذف الملف «' + n + '» من الأداة؟\n(لن يُحذف من جهازك)')) return;
    delete store.files[n];
    store.current = Object.keys(store.files)[0] || '';
    wb = store.files[store.current] || null;
    saveStore(); clearMarks(); fillSheets();
    status(wb ? '🗑 تم حذف الملف.' : '🗑 تم حذف الملف. ارفع ملفًا جديدًا.');
  };
  $('rz-sheet').onchange = () => { if (wb) { wb.lastSheet = $('rz-sheet').value; saveStore(); } $('rz-auto').textContent = ''; fillCols($('rz-sheet').value); };
  $('rz-name').onchange = updateInfo;
  $('rz-phrase').onchange = updateInfo;

  // ================= المطابقة والرصد =================
  function clearMarks() { document.querySelectorAll('.rz-ok,.rz-fuzzy,.rz-miss,.rz-save-hi').forEach(el => el.classList.remove('rz-ok', 'rz-fuzzy', 'rz-miss', 'rz-save-hi')); }
  function setVal(ta, v) {
    const setter = Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value').set;
    setter.call(ta, v);
    ['input', 'change', 'keyup', 'blur'].forEach(t => ta.dispatchEvent(new Event(t, { bubbles: true })));
  }
  function plan() {
    if (!wb) { status('⚠️ ارفع ملف الإكسل أولًا.'); return null; }
    const sheetName = $('rz-sheet').value, nameCol = +$('rz-name').value, phraseCol = +$('rz-phrase').value;
    if (phraseCol < 0) { status('⚠️ اختر عمود العبارة الوصفية.'); return null; }
    const xl = students(wb.sheets[sheetName], nameCol, phraseCol);
    const pg = pageStudents();
    if (!pg.length) { status('⚠️ لم أجد أسماء طلاب في الصفحة — اضغط «عرض» في البوابة أولًا.'); return null; }
    return planCore(xl, pg, sheetName);
  }
  function planCore(xl, pg, sheetName) {
    const used = new Set();
    const items = pg.map(s => {
      const nn = norm(s.name);
      let idx = xl.findIndex((x, i) => !used.has(i) && norm(x.name) === nn), kind = 'exact', sc = 1;
      if (idx < 0 && opt.fuzzy) {
        let best = -1, bs = 0, second = 0;
        xl.forEach((x, i) => { if (used.has(i)) return; const v = score(s.name, x.name); if (v > bs) { second = bs; bs = v; best = i; } else if (v > second) second = v; });
        if (best >= 0 && bs >= 0.75 && bs > second) { idx = best; kind = 'fuzzy'; sc = bs; }
      }
      if (idx >= 0) used.add(idx);
      const x = idx >= 0 ? xl[idx] : null;
      let st;
      if (!x) st = 'nomatch';
      else if (!x.phrase) st = 'nophrase';
      else if (!s.textarea) st = 'nobox';
      else if (opt.noOverwrite && s.textarea.value.trim() && s.textarea.value.trim() !== x.phrase) st = 'skip';
      else if (s.textarea.value.trim() === x.phrase) st = 'same';
      else st = kind === 'exact' ? 'fill' : 'fillfuzzy';
      const long = x && x.phrase && s.textarea && s.textarea.maxLength > 0 && x.phrase.length > s.textarea.maxLength;
      return { s, x, kind, sc, st, long };
    });
    const extra = xl.filter((x, i) => !used.has(i) && x.phrase);
    return { sheetName, items, extra, total: pg.length };
  }
  const LABEL = {
    fill: ['✔', 'c-ok', 'رُصدت'], fillfuzzy: ['≈', 'c-w', 'مطابقة تقريبية'], same: ['✔', 'c-ok', 'موجودة مسبقًا (نفس العبارة)'],
    skip: ['⏭', 'c-w', 'تُركت — توجد عبارة أخرى'], nophrase: ['○', 'c-w', 'لا توجد عبارة في الإكسل'],
    nomatch: ['✖', 'c-b', 'غير موجود في الإكسل'], nobox: ['✖', 'c-b', 'لم أجد مربع الوصف'],
  };
  function renderList(p, done) {
    const rows = p.items.filter(i => !['fill', 'same'].includes(i.st) || i.long).map(i => {
      const [k, c, t] = LABEL[i.st];
      const extra = i.st === 'fillfuzzy' ? ` ← ${esc(i.x.name)}` : '';
      return `<div><span class="k ${c}">${k}</span><span>${esc(i.s.name)} — <span class="${c}">${done || !i.st.startsWith('fill') ? t : 'ستُرصد' + (i.st === 'fillfuzzy' ? ' (تقريبي)' : '')}${extra}</span>${i.long ? ' <span class="c-b">(العبارة أطول من المسموح)</span>' : ''}</span></div>`;
    });
    p.extra.forEach(x => rows.push(`<div><span class="k c-b">!</span><span>${esc(x.name)} — <span class="c-b">في الإكسل فقط</span></span></div>`));
    $('rz-list').innerHTML = rows.join('');
  }
  function counts(p) {
    const c = {}; p.items.forEach(i => c[i.st] = (c[i.st] || 0) + 1);
    return { ok: (c.fill || 0) + (c.same || 0), fuzzy: c.fillfuzzy || 0, nophrase: c.nophrase || 0, skip: c.skip || 0, miss: (c.nomatch || 0) + (c.nobox || 0), extra: p.extra.length };
  }
  function summary(p, done) {
    const c = counts(p);
    return `📑 الورقة ${p.sheetName} — 👥 طلاب الصفحة: ${p.total}
${done ? '✔ رُصدت' : '✔ مطابقة تامة'}: ${c.ok}${c.fuzzy ? `\n≈ تقريبية (راجعها): ${c.fuzzy}` : ''}${c.nophrase ? `\n○ بدون عبارة في الإكسل: ${c.nophrase}` : ''}${c.skip ? `\n⏭ تُركت: ${c.skip}` : ''}${c.miss ? `\n✖ لم تُطابق: ${c.miss}` : ''}${c.extra ? `\n! في الإكسل وليست في الصفحة: ${c.extra}` : ''}`;
  }

  let busy = false, last = null;
  async function run() {
    if (busy) return;
    const p = plan(); if (!p) return;
    busy = true; clearMarks();
    const todo = p.items.filter(i => i.st === 'fill' || i.st === 'fillfuzzy');
    $('rz-prog').style.display = 'block';
    for (let n = 0; n < todo.length; n++) {
      const i = todo[n];
      setVal(i.s.textarea, i.x.phrase);
      $('rz-bar').style.width = ((n + 1) / todo.length * 100) + '%';
      status(`✍️ رصد ${n + 1} / ${todo.length}: ${i.s.name}`);
      if (n % 3 === 0) await sleep(20);
    }
    p.items.forEach(i => {
      const ta = i.s.textarea; if (!ta) return;
      if (i.st === 'fill' || i.st === 'same') ta.classList.add('rz-ok');
      else if (i.st === 'fillfuzzy') ta.classList.add('rz-fuzzy');
      else if (i.st === 'nomatch') ta.classList.add('rz-miss');
    });
    setTimeout(() => ($('rz-prog').style.display = 'none'), 800);
    status(summary(p, true)); renderList(p, true);
    last = { date: today(), sheet: p.sheetName, file: wb.fileName, rows: p.items.map(i => [i.s.name, i.s.id, i.x ? i.x.name : '', LABEL[i.st][2], i.x ? i.x.phrase : '']), extra: p.extra.map(x => x.name) };
    set(LASTK, last);
    busy = false;
    const c = counts(p);
    if (opt.autoSave && !c.fuzzy && !c.miss && todo.length) {
      const b = findSaveBtn();
      if (b) { status(summary(p, true) + '\n💾 جاري الحفظ...'); await sleep(600); b.click(); return; }
    }
    showDone(p, todo.length);
  }

  // ================= فقاعة الانتهاء =================
  function chime() {
    try {
      const ac = new (window.AudioContext || window.webkitAudioContext)();
      [523, 659, 784, 1047].forEach((f, i) => {
        const o = ac.createOscillator(), g = ac.createGain();
        o.type = 'sine'; o.frequency.value = f;
        g.gain.setValueAtTime(0.0001, ac.currentTime + i * 0.15);
        g.gain.exponentialRampToValueAtTime(0.25, ac.currentTime + i * 0.15 + 0.02);
        g.gain.exponentialRampToValueAtTime(0.0001, ac.currentTime + i * 0.15 + 0.5);
        o.connect(g).connect(ac.destination); o.start(ac.currentTime + i * 0.15); o.stop(ac.currentTime + i * 0.15 + 0.55);
      });
    } catch (e) {}
  }
  function gotoSave() {
    const b = findSaveBtn();
    if (!b) { status('⚠️ لم أجد زر «حفظ» في الصفحة.'); return; }
    b.scrollIntoView({ behavior: 'smooth', block: 'center' });
    b.classList.add('rz-save-hi');
    setTimeout(() => b.classList.remove('rz-save-hi'), 6000);
  }
  function showDone(p, filled) {
    const old = document.getElementById('rz-done'); if (old) old.remove();
    const c = counts(p);
    const clean = !c.fuzzy && !c.miss && !c.extra;
    const accent = clean ? '#10b981' : '#f59e0b';
    const box = document.createElement('div');
    box.id = 'rz-done'; box.dir = 'rtl'; box.setAttribute('data-rzt', curTheme());
    box.innerHTML = `
      <style>
        #rz-done{position:fixed;inset:0;z-index:2147483647;background:rgba(12,10,9,.5);backdrop-filter:blur(5px);display:flex;align-items:center;justify-content:center;font-family:'Cairo',Tahoma,Arial,sans-serif;animation:rzF .3s ease}
        #rz-done *{box-sizing:border-box;font-family:'Cairo',Tahoma,Arial,sans-serif}
        #rz-done .card{position:relative;background:var(--bg);color:var(--ink);border-radius:26px;padding:32px 26px 20px;width:min(420px,92vw);text-align:center;box-shadow:0 25px 70px rgba(0,0,0,.45);border:1px solid var(--line);border-top:6px solid ${accent};animation:rzP .45s cubic-bezier(.2,1.4,.4,1)}
        #rz-done .ic{width:84px;height:84px;margin:-72px auto 12px;background:var(--solid);border:3px solid ${accent};border-radius:24px;display:grid;place-items:center;font-size:42px;box-shadow:0 12px 28px rgba(0,0,0,.25);animation:rzB 1.8s ease-in-out infinite}
        #rz-done h2{margin:4px 0;font-size:19px;font-weight:900}
        #rz-done p.s{margin:2px 0 10px;color:var(--hi);font-size:13px;font-weight:700}
        #rz-done .stats{display:flex;gap:8px;margin:12px 0 8px}
        #rz-done .st{flex:1;background:var(--box);border:1px solid var(--line);border-radius:14px;padding:9px 4px;font-size:11.5px;color:var(--muted);font-weight:700}
        #rz-done .st b{display:block;font-size:20px;font-weight:900;color:var(--ink)}
        #rz-done .st.w b{color:#d97706}#rz-done .st.r b{color:#dc2626}
        #rz-done .note{font-size:12px;padding:8px 12px;border-radius:10px;background:var(--soft);margin:10px 0;font-weight:700}
        #rz-done .btns{display:flex;gap:8px;margin-top:14px}
        #rz-done button{flex:1;border:1px solid var(--line);border-radius:12px;padding:10px 6px;font-weight:800;font-size:12.5px;cursor:pointer;background:var(--btn);color:var(--ink)}
        #rz-done button.main{background:linear-gradient(135deg,var(--a1),var(--a2));color:var(--aink);border-color:transparent;font-weight:900}
        #rz-done .x{position:absolute;top:12px;left:12px;width:30px;height:30px;border-radius:10px;background:var(--btn);border:1px solid var(--line);display:grid;place-items:center;cursor:pointer;font-weight:900;color:var(--muted)}
        #rz-done .credit{margin-top:12px;padding-top:9px;border-top:1px solid var(--line);font-size:11px;color:var(--muted);font-weight:600}
        @keyframes rzF{from{opacity:0}}@keyframes rzP{from{transform:scale(.7);opacity:0}}@keyframes rzB{50%{transform:translateY(-5px) rotate(-4deg)}}
      </style>
      <div class="card">
        <span class="x" data-x="close">✕</span>
        <div class="ic">${clean ? '✅' : '📝'}</div>
        <h2>${clean ? 'اكتمل رصد العبارات بنجاح' : 'اكتمل الرصد — مع ملاحظات'}</h2>
        <p class="s">الورقة ${esc(p.sheetName)} • ${p.total} طالب في الصفحة</p>
        <div class="stats">
          <div class="st"><b>${filled}</b>رُصدت الآن</div>
          <div class="st ${c.fuzzy ? 'w' : ''}"><b>${c.fuzzy}</b>تقريبية</div>
          <div class="st ${c.miss + c.extra ? 'r' : ''}"><b>${c.miss + c.extra}</b>لم تُطابق</div>
        </div>
        <div class="note">${clean ? '💾 لا تنسَ الضغط على «حفظ» أسفل صفحة البوابة.' : '⚠️ راجع المربعات البرتقالية والحمراء، ثم اضغط «حفظ».'}</div>
        <div class="btns">
          <button class="main" data-x="save">⬇️ الذهاب لزر الحفظ</button>
          <button data-x="report">📥 تقرير Excel</button>
          <button data-x="close">إغلاق</button>
        </div>
        <div class="credit">تصميم وتطوير: سيف جمال المعولي • معولة بن شمس (5-12)</div>
      </div>`;
    const close = () => box.remove();
    box.addEventListener('click', e => {
      const x = (e.target.closest('[data-x]') || {}).dataset?.x;
      if (e.target === box || x === 'close') close();
      if (x === 'save') { close(); gotoSave(); }
      if (x === 'report') exportReport();
    });
    document.body.appendChild(box);
    chime();
    try { GM_notification({ title: '📝 راصد العبارات الوصفية', text: `رُصدت ${filled} عبارة — الورقة ${p.sheetName}`, timeout: 10000 }); } catch (e) {}
  }

  // ================= تقرير Excel =================
  function exportReport() {
    const d = last || get(LASTK, null);
    if (!d) { status('⚠️ لا يوجد رصد سابق — اضغط «رصد العبارات» أولًا.'); return; }
    if (d.auto) return exportAutoReport(d);
    const aoa = [
      ['تقرير رصد العبارات الوصفية'], ['الملف', d.file], ['الورقة', d.sheet], ['التاريخ', d.date], [],
      ['م', 'اسم الطالب في البوابة', 'الرقم', 'الاسم في الإكسل', 'الحالة', 'العبارة الوصفية'],
      ...d.rows.map((r, i) => [i + 1, ...r]),
    ];
    if (d.extra.length) aoa.push([], ['طلاب في الإكسل لديهم عبارة وليسوا في صفحة البوابة'], ...d.extra.map((n, i) => [i + 1, n]));
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws['!cols'] = [{ wch: 5 }, { wch: 36 }, { wch: 15 }, { wch: 36 }, { wch: 26 }, { wch: 70 }];
    const book = XLSX.utils.book_new(); book.Workbook = { Views: [{ RTL: true }] };
    XLSX.utils.book_append_sheet(book, ws, 'تقرير الرصد');
    XLSX.writeFile(book, `تقرير_رصد_العبارات_${String(d.sheet).replace(/[\\\/?*\[\]:]/g, '-')}_${d.date}.xlsx`);
  }

  function diagnose() {
    const pg = pageStudents();
    const withBox = pg.filter(s => s.textarea).length;
    status([
      '🔍 فحص الصفحة:',
      'أسماء الطلاب: ' + pg.length,
      'مربعات «وصف أداء الطالب»: ' + withBox,
      'زر الحفظ: ' + (findSaveBtn() ? 'موجود' : 'غير موجود'),
      'الشعبة المقترحة: ' + (wb ? (guessSheet(Object.keys(wb.sheets)) || 'لم تُحدد') : '—'),
      'مكتبة الإكسل: ' + (typeof XLSX !== 'undefined' ? 'تعمل' : 'غير محمّلة ❌'),
    ].join('\n'));
  }

  // ================= الرصد التلقائي لكل الصفوف والشعب =================
  const AUTOK = 'rz_auto';
  const AR = s => String(s || '').replace(/[ً-ٟـ]/g, '').replace(/[أإآ]/g, 'ا').replace(/ة/g, 'ه').replace(/ى/g, 'ي').replace(/[:：]/g, '').replace(/\s+/g, ' ').trim();
  const GRADES = { 'الخامس': 5, 'السادس': 6, 'السابع': 7, 'الثامن': 8, 'التاسع': 9, 'العاشر': 10, 'الحادي عشر': 11, 'الثاني عشر': 12 };
  let leaving = false, stopAuto = false, autoBusy = false;
  window.addEventListener('beforeunload', () => { leaving = true; });
  window.addEventListener('pagehide', () => { leaving = true; });
  const isVis = el => !!(el && el.getClientRects().length && getComputedStyle(el).visibility !== 'hidden');
  const SKIP_LABELS = ['الصف', 'الشعبه', 'اسم الطالب', 'الرقم المدرسي', 'وصف اداء الطالب', 'المجموع'];

  function selLabel(sel) {
    const td = sel.closest('td'); if (!td) return '';
    let p = td.previousElementSibling;
    while (p && !AR(p.innerText)) p = p.previousElementSibling;
    return p ? AR(p.innerText) : '';
  }
  const pageSelects = () => [...document.querySelectorAll('select')]
    .filter(s => !s.closest('#rz-panel,#rz-done') && isVis(s)).map(el => ({ el, label: selLabel(el) }));
  const findSel = lab => (pageSelects().find(x => x.label === AR(lab)) || {}).el || null;
  const realOpts = sel => [...sel.options].filter(o => o.value && !['0', '-1'].includes(o.value) && !/اختر|^-+|الكل|select/i.test(o.text));
  const selText = sel => ((sel.options[sel.selectedIndex] || {}).text || '').trim();
  function findShowBtn() {
    return [...document.querySelectorAll('input[type=submit],input[type=button],button,a')]
      .find(el => !el.closest('#rz-panel,#rz-done') && isVis(el) && /^\s*عرض\s*$/.test(el.value || el.innerText || '')) || null;
  }
  const studSig = () => pageStudents().map(s => s.id).join(',');
  async function changeSel(sel, value) {
    sel.value = value;
    sel.dispatchEvent(new Event('change', { bubbles: true }));
    await sleep(2500);
  }
  async function waitUntil(fn, ms) {
    const end = Date.now() + ms;
    while (Date.now() < end && !stopAuto) { const v = fn(); if (v) return v; await sleep(400); }
    return null;
  }
  // رسائل البوابة (alert / confirm) لا توقف الرصد التلقائي؛ تُحفظ الرسالة في التقرير
  function muteDialogs(on) {
    const sc = document.createElement('script');
    sc.textContent = on
      ? "if(!window.__rzA){window.__rzA=[window.alert,window.confirm];window.alert=function(m){document.documentElement.setAttribute('data-rz-msg',String(m))};window.confirm=function(m){document.documentElement.setAttribute('data-rz-msg',String(m));return true};}"
      : "if(window.__rzA){window.alert=window.__rzA[0];window.confirm=window.__rzA[1];delete window.__rzA;}";
    (document.head || document.documentElement).appendChild(sc); sc.remove();
  }
  const portalMsg = () => { const m = document.documentElement.getAttribute('data-rz-msg') || ''; document.documentElement.removeAttribute('data-rz-msg'); return m; };

  function gradeNum(t) { const a = AR(t); for (const k in GRADES) if (a === AR(k)) return GRADES[k]; const m = a.match(/\d+/); return m ? +m[0] : null; }
  function findSheetFor(gradeText, secText) {
    const g = gradeNum(gradeText), sc = (String(secText).match(/\d+/) || [])[0];
    if (!g || !sc) return null;
    const want = g + '.' + sc;
    const order = [store.current, ...Object.keys(store.files).filter(n => n !== store.current)];
    for (const fn of order) {
      const f = store.files[fn]; if (!f) continue;
      const sn = Object.keys(f.sheets).find(n => n.replace(/\s/g, '') === want);
      if (sn) return { file: fn, sheetName: sn, sh: f.sheets[sn] };
    }
    return null;
  }
  function bumpA(s, key) {
    if (s.rk === key) s.rc++; else { s.rk = key; s.rc = 0; }
    set(AUTOK, s); return s.rc > 3;
  }
  function nextSec(s) { s.si++; s.phase = 'sel'; s.rk = ''; s.rc = 0; s.pending = null; set(AUTOK, s); }
  function nextGrade(s) { s.gi++; s.si = 0; s.phase = 'sel'; s.rk = ''; s.rc = 0; s.pending = null; set(AUTOK, s); }
  function logA(s, e) { s.log.push(Object.assign({ grade: '', sec: '', sheet: '', filled: 0, fuzzy: 0, miss: 0, nophrase: 0, status: '', note: '' }, e)); set(AUTOK, s); }
  function autoStatus(s, msg) {
    const g = s.cur ? `الصف ${s.gi + 1}/${s.nG || '?'}: ${s.cur.grade}${s.cur.sec ? ` — الشعبة ${s.si + 1}/${s.nS || '?'}: ${s.cur.sec}` : ''}` : '';
    const saved = s.log.filter(x => x.status.startsWith('✅')).length;
    status(`🔁 الرصد التلقائي\n${g}\n${msg}\n💾 شعب محفوظة: ${saved}`);
    $('rz-prog').style.display = 'block';
    if (s.nG) $('rz-bar').style.width = Math.min(100, ((s.gi + (s.nS ? s.si / s.nS : 0)) / s.nG) * 100) + '%';
  }
  function showRunning(on) {
    panel.querySelector('[data-a=all]').style.display = on ? 'none' : '';
    panel.querySelector('[data-a=stopall]').style.display = on ? '' : 'none';
    panel.querySelector('[data-a=go]').disabled = on;
  }

  function startAuto() {
    if (!Object.keys(store.files).length) { status('⚠️ ارفع ملف الإكسل أولًا.'); return; }
    const gSel = findSel('الصف'), sSel = findSel('الشعبة');
    if (!gSel || !sSel) { status('⚠️ لم أجد قائمتي «الصف» و«الشعبة» في الصفحة.'); return; }
    const nG = realOpts(gSel).length;
    // نحفظ اختيارات المعلم الأخرى (المادة، الاختبار، الفصل...) لنعيدها بعد تغيير الصف
    const fixed = {};
    pageSelects().forEach(({ el, label }) => {
      if (!label || SKIP_LABELS.includes(label) || /وصف|المجموع/.test(label)) return;
      const t = selText(el); if (t && !/اختر|^-+/.test(t)) fixed[label] = t;
    });
    const fx = Object.entries(fixed).map(([k, v]) => `• ${k}: ${v}`).join('\n');
    if (!confirm(`سيمر الراصد على كل الصفوف (${nG}) وكل الشعب، ويرصد العبارات ويضغط «حفظ» لكل شعبة.\n\nسيُثبَّت لكل الصفوف:\n${fx}\n\nكل شعبة تحتاج ورقة باسمها في الإكسل (مثل 8.2).\nمتابعة؟`)) return;
    stopAuto = false;
    set(AUTOK, { gi: 0, si: 0, phase: 'sel', fixed, log: [], details: [], rk: '', rc: 0, started: today() });
    showRunning(true); clearMarks();
    autoStep();
  }

  async function autoStep() {
    if (autoBusy) return;
    let s = get(AUTOK, null); if (!s) return;
    autoBusy = true; showRunning(true); muteDialogs(true);
    try {
      for (let guard = 0; guard < 60 && !leaving && !stopAuto; guard++) {
        s = get(AUTOK, null); if (!s) return;

        // ---- الصف ----
        const gSel = findSel('الصف');
        if (!gSel) { if (bumpA(s, 'nog')) { autoFinish(s, true, 'لم تظهر قائمة الصف'); return; } await sleep(1500); continue; }
        const grades = realOpts(gSel);
        s.nG = grades.length;
        if (s.gi >= grades.length) { autoFinish(s, false); return; }
        const g = grades[s.gi];
        s.cur = { grade: g.text.trim(), sec: '' }; set(AUTOK, s);
        if (gSel.value !== g.value) {
          if (bumpA(s, 'g' + s.gi)) { logA(s, { grade: g.text, status: '❌ تخطي', note: 'تعذر اختيار الصف' }); nextGrade(s); continue; }
          s.phase = 'sel'; set(AUTOK, s);
          autoStatus(s, 'اختيار الصف...');
          await changeSel(gSel, g.value); continue;
        }

        // ---- إعادة المادة والاختبار وغيرها كما اختارها المعلم ----
        let changed = false, missing = null;
        for (const [lab, txt] of Object.entries(s.fixed || {})) {
          const x = findSel(lab); if (!x) continue;
          if (AR(selText(x)) === AR(txt)) continue;
          const o = [...x.options].find(o => AR(o.text) === AR(txt));
          if (!o) { missing = [lab, txt]; break; }
          if (bumpA(s, 'f' + lab + s.gi)) { missing = [lab, txt]; break; }
          autoStatus(s, `اختيار ${lab}: ${txt}`);
          await changeSel(x, o.value); changed = true; break;
        }
        if (leaving) return;
        if (missing) {
          if (!bumpA(s, 'miss' + s.gi + missing[0])) { await sleep(1500); continue; } // ربما القائمة لم تكتمل بعد
          logA(s, { grade: g.text, status: '⚠️ تخطي الصف', note: `قائمة «${missing[0]}» لا تحتوي «${missing[1]}» لهذا الصف` });
          nextGrade(s); continue;
        }
        if (changed) continue;

        // ---- الشعبة ----
        const sSel = findSel('الشعبة');
        const secs = sSel ? realOpts(sSel) : [];
        if (!secs.length) {
          if (bumpA(s, 'nos' + s.gi)) { logA(s, { grade: g.text, status: '⚠️ تخطي الصف', note: 'لا توجد شعب لهذا الصف' }); nextGrade(s); continue; }
          await sleep(1500); continue;
        }
        s.nS = secs.length;
        if (s.si >= secs.length) { nextGrade(s); continue; }
        const sec = secs[s.si];
        s.cur.sec = sec.text.trim(); set(AUTOK, s);
        const key = s.gi + '.' + s.si;
        if (sSel.value !== sec.value) {
          if (bumpA(s, 's' + key)) { logA(s, { grade: g.text, sec: sec.text, status: '❌ تخطي', note: 'تعذر اختيار الشعبة' }); nextSec(s); continue; }
          s.phase = 'sel'; set(AUTOK, s);
          autoStatus(s, 'اختيار الشعبة...');
          await changeSel(sSel, sec.value); continue;
        }

        const found = findSheetFor(g.text, sec.text);
        if (!found) {
          logA(s, { grade: g.text, sec: sec.text, status: '⚠️ لا توجد ورقة', note: `لا توجد ورقة باسم ${gradeNum(g.text)}.${(sec.text.match(/\d+/) || [''])[0]} في ملفات الإكسل` });
          nextSec(s); continue;
        }

        // ---- عرض ----
        if (s.phase === 'sel') {
          const btn = findShowBtn();
          if (!btn) { logA(s, { grade: g.text, sec: sec.text, status: '❌ تخطي', note: 'لم أجد زر «عرض»' }); nextSec(s); continue; }
          s.phase = 'shown'; s.sig = studSig(); set(AUTOK, s);
          autoStatus(s, '👥 عرض الطلاب...');
          btn.click();
          await waitUntil(() => leaving || (studSig() && studSig() !== s.sig), 20000);
          if (leaving) return;
          await sleep(800); continue;
        }

        // ---- رصد وحفظ ----
        if (s.phase === 'shown') {
          const pg = pageStudents();
          if (!pg.length) {
            if (bumpA(s, 'show' + key)) { logA(s, { grade: g.text, sec: sec.text, sheet: found.sheetName, status: '❌ تخطي', note: 'لم يظهر الطلاب بعد «عرض»' }); nextSec(s); continue; }
            s.phase = 'sel'; set(AUTOK, s); continue;
          }
          const sh = found.sh;
          const p = planCore(students(sh, sh.info.nameCol, sh.info.phraseCol), pg, found.sheetName);
          const todo = p.items.filter(i => i.st === 'fill' || i.st === 'fillfuzzy');
          clearMarks();
          todo.forEach(i => setVal(i.s.textarea, i.x.phrase));
          p.items.forEach(i => { const ta = i.s.textarea; if (!ta) return; if (i.st === 'fill' || i.st === 'same') ta.classList.add('rz-ok'); else if (i.st === 'fillfuzzy') ta.classList.add('rz-fuzzy'); else if (i.st === 'nomatch') ta.classList.add('rz-miss'); });
          const c = counts(p);
          const tot = +((document.body.innerText.match(/عدد الطلبة\s*(\d+)/) || [])[1] || 0);
          const notes = [];
          if (sh.info.phraseCol < 0) notes.push('لم يُعثر على عمود العبارات في الورقة');
          if (tot > pg.length) notes.push(`عدد الطلبة ${tot} والظاهر ${pg.length} — بقية الطلاب في صفحة أخرى`);
          if (c.extra) notes.push(`${c.extra} في الإكسل وليسوا في البوابة`);
          p.items.forEach(i => s.details.push([g.text.trim(), sec.text.trim(), i.s.name, i.s.id, i.x ? i.x.name : '', LABEL[i.st][2], i.x ? i.x.phrase : '']));
          const entry = { grade: g.text.trim(), sec: sec.text.trim(), sheet: found.sheetName, filled: todo.length, fuzzy: c.fuzzy, miss: c.miss, nophrase: c.nophrase, note: notes.join(' • ') };
          if (!todo.length) { logA(s, Object.assign(entry, { status: '○ لا جديد', note: (entry.note ? entry.note + ' • ' : '') + 'لا توجد عبارات جديدة لرصدها' })); nextSec(s); continue; }
          const sv = findSaveBtn();
          if (!sv) { logA(s, Object.assign(entry, { status: '❌ لم يُحفظ', note: 'لم أجد زر «حفظ»' })); nextSec(s); continue; }
          s.pending = entry; s.phase = 'saved'; set(AUTOK, s);
          autoStatus(s, `✍️ رُصدت ${todo.length} — 💾 جاري الحفظ...`);
          await sleep(700);
          sv.click();
          await waitUntil(() => leaving, 12000);
          if (leaving) return;
          continue; // الحفظ تم بدون إعادة تحميل
        }

        if (s.phase === 'saved') {
          const m = portalMsg();
          const e = Object.assign({}, s.pending || { grade: g.text, sec: sec.text }, { status: '✅ تم الحفظ' });
          if (m) e.note = (e.note ? e.note + ' • ' : '') + 'رسالة البوابة: ' + m;
          logA(s, e); nextSec(s);
          autoStatus(s, `✅ حُفظت الشعبة ${e.sec}`);
          await sleep(600); continue;
        }
        s.phase = 'sel'; set(AUTOK, s);
      }
    } finally {
      autoBusy = false;
      if (stopAuto) { const s2 = get(AUTOK, null); if (s2) autoFinish(s2, true, 'أوقفه المستخدم'); }
    }
  }

  function autoFinish(s, partial, why) {
    set(AUTOK, null); try { if (hasGM) GM_deleteValue(AUTOK); } catch (e) {}
    muteDialogs(false); showRunning(false);
    $('rz-prog').style.display = 'none';
    last = { auto: true, date: today(), log: s.log, details: s.details, partial: !!partial, why: why || '' };
    set(LASTK, last);
    const ok = s.log.filter(x => x.status.startsWith('✅')).length;
    const filled = s.log.reduce((a, x) => a + (x.status.startsWith('✅') ? x.filled : 0), 0);
    const bad = s.log.filter(x => !x.status.startsWith('✅') && !x.status.startsWith('○')).length;
    status(`${partial ? '⏹ توقف الرصد التلقائي' + (why ? ' — ' + why : '') : '✅ انتهى الرصد التلقائي'}\n💾 شعب محفوظة: ${ok}\n✍️ عبارات رُصدت: ${filled}${bad ? `\n⚠️ شعب بها ملاحظة: ${bad} (انظر التقرير)` : ''}`);
    $('rz-list').innerHTML = s.log.map(x => `<div><span class="k ${x.status.startsWith('✅') ? 'c-ok' : x.status.startsWith('○') ? 'c-w' : 'c-b'}">${esc(x.status.split(' ')[0])}</span><span>${esc(x.grade)} ${x.sec ? '/ ' + esc(x.sec) : ''} — ${esc(x.status.slice(2))}${x.filled ? ` (${x.filled})` : ''}${x.note ? ` <span class="c-w">${esc(x.note)}</span>` : ''}</span></div>`).join('');
    showAutoDone(s, ok, filled, bad, partial);
  }

  function showAutoDone(s, ok, filled, bad, partial) {
    const fake = { sheetName: 'كل الصفوف', total: s.log.length, items: [], extra: [] };
    showDone(fake, filled);
    const card = document.querySelector('#rz-done .card'); if (!card) return;
    card.querySelector('h2').textContent = partial ? 'توقف الرصد التلقائي قبل الاكتمال' : bad ? 'اكتمل رصد كل الشعب — مع ملاحظات' : 'اكتمل رصد وحفظ كل الشعب بنجاح';
    card.querySelector('p.s').textContent = `${s.log.length} شعبة مرّ عليها الراصد`;
    card.querySelector('.stats').innerHTML = `<div class="st"><b>${ok}</b>شعبة حُفظت</div><div class="st"><b>${filled}</b>عبارة رُصدت</div><div class="st ${bad ? 'w' : ''}"><b>${bad}</b>شعبة بها ملاحظة</div>`;
    card.querySelector('.note').textContent = bad || partial ? '⚠️ نزّل «تقرير Excel» لمعرفة الشعب والطلاب الذين يحتاجون مراجعة.' : '✅ تم الحفظ في البوابة لكل الشعب.';
    const sb = card.querySelector('[data-x=save]'); if (sb) sb.remove();
  }

  function exportAutoReport(d) {
    const sum = [
      ['تقرير الرصد التلقائي للعبارات الوصفية'], ['التاريخ', d.date], d.partial ? ['⏹ تنبيه', 'توقف قبل الاكتمال' + (d.why ? ' — ' + d.why : '')] : [], [],
      ['م', 'الصف', 'الشعبة', 'الورقة', 'الحالة', 'عبارات رُصدت', 'تقريبية', 'لم تُطابق', 'بدون عبارة', 'ملاحظة'],
      ...d.log.map((x, i) => [i + 1, x.grade, x.sec, x.sheet, x.status, x.filled, x.fuzzy, x.miss, x.nophrase, x.note]),
    ];
    const ws1 = XLSX.utils.aoa_to_sheet(sum);
    ws1['!cols'] = [{ wch: 5 }, { wch: 14 }, { wch: 8 }, { wch: 10 }, { wch: 16 }, { wch: 12 }, { wch: 9 }, { wch: 10 }, { wch: 11 }, { wch: 60 }];
    const ws2 = XLSX.utils.aoa_to_sheet([['الصف', 'الشعبة', 'اسم الطالب في البوابة', 'الرقم', 'الاسم في الإكسل', 'الحالة', 'العبارة الوصفية'], ...d.details]);
    ws2['!cols'] = [{ wch: 12 }, { wch: 8 }, { wch: 36 }, { wch: 15 }, { wch: 36 }, { wch: 26 }, { wch: 70 }];
    ws2['!autofilter'] = { ref: ws2['!ref'] };
    const book = XLSX.utils.book_new(); book.Workbook = { Views: [{ RTL: true }] };
    XLSX.utils.book_append_sheet(book, ws1, 'ملخص الشعب');
    XLSX.utils.book_append_sheet(book, ws2, 'تفاصيل الطلاب');
    XLSX.writeFile(book, `تقرير_الرصد_التلقائي_${d.date}.xlsx`);
  }

  panel.addEventListener('click', e => {
    const a = (e.target.closest('button[data-a]') || {}).dataset?.a;
    if (a === 'all') startAuto();
    if (a === 'stopall') { stopAuto = true; const s = get(AUTOK, null); if (s && !autoBusy) autoFinish(s, true, 'أوقفه المستخدم'); }
    if (a === 'go') run();
    if (a === 'preview') { const p = plan(); if (p) { clearMarks(); status('👁️ معاينة (لم يُكتب شيء)\n' + summary(p, false)); renderList(p, false); } }
    if (a === 'report') exportReport();
    if (a === 'tosave') gotoSave();
    if (a === 'clear') { clearMarks(); $('rz-list').innerHTML = ''; status('🧹 تمت إزالة التظليل.'); }
    if (a === 'diag') diagnose();
  });

  fillSheets();
  if (get(AUTOK, null)) {
    // استئناف تلقائي بعد كل تحميل للصفحة
    showRunning(true); status('🔁 استئناف الرصد التلقائي...');
    setTimeout(autoStep, 1500);
  } else if (wb) {
    const n = pageStudents().length;
    if (n) status(`👥 ${n} طالب في الصفحة — اضغط «رصد العبارات».`);
  }
  }
})();
