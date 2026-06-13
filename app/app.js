/* =====================================================================
   Trening Analytics — logika aplikacji
   ===================================================================== */
'use strict';

// ---------------------------------------------------------------- stan
const State = {
  wb: null,          // parsowany workbook (xlsx-engine)
  model: null,       // model treningowy (training-model)
  fileHandle: null,  // uchwyt pliku (File System Access) – do zapisu w miejscu
  fileName: null,
  view: 'dashboard',
  charts: [],        // aktywne instancje Chart (do niszczenia)
  selectedExercise: 0,
  entryWeek: null
};

// ---------------------------------------------------------------- utils
const $ = (sel, root = document) => root.querySelector(sel);
const $$ = (sel, root = document) => Array.from(root.querySelectorAll(sel));
const el = (tag, attrs = {}, html) => {
  const n = document.createElement(tag);
  for (const k in attrs) {
    if (k === 'class') n.className = attrs[k];
    else if (k === 'style') n.style.cssText = attrs[k];
    else if (k.startsWith('on')) n.addEventListener(k.slice(2), attrs[k]);
    else if (attrs[k] != null) n.setAttribute(k, attrs[k]);
  }
  if (html != null) n.innerHTML = html;
  return n;
};
const fmt = (n, dec = 0) => (n == null || isNaN(n)) ? '—'
  : Number(n).toLocaleString('pl-PL', { minimumFractionDigits: dec, maximumFractionDigits: dec });
const fmtKg = (n) => n == null ? '—' : fmt(n, n % 1 ? 1 : 0) + ' kg';
const esc = (s) => String(s == null ? '' : s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
const weekLabels = (n) => Array.from({ length: n }, (_, i) => 'T' + (i + 1));

function toast(msg, type = '') {
  const t = el('div', { class: 'toast ' + type }, msg);
  $('#toastWrap').appendChild(t);
  setTimeout(() => { t.style.opacity = '0'; t.style.transform = 'translateX(30px)'; t.style.transition = 'all .3s'; }, 3200);
  setTimeout(() => t.remove(), 3600);
}

// --------------------------------------------------------- IndexedDB (uchwyt pliku)
const IDB = {
  db: null,
  open() {
    return new Promise((res) => {
      const r = indexedDB.open('trening-analytics', 1);
      r.onupgradeneeded = () => r.result.createObjectStore('kv');
      r.onsuccess = () => { this.db = r.result; res(this.db); };
      r.onerror = () => res(null);
    });
  },
  async set(k, v) {
    if (!this.db) await this.open();
    if (!this.db) return;
    return new Promise((res) => {
      const tx = this.db.transaction('kv', 'readwrite');
      tx.objectStore('kv').put(v, k); tx.oncomplete = () => res(); tx.onerror = () => res();
    });
  },
  async get(k) {
    if (!this.db) await this.open();
    if (!this.db) return null;
    return new Promise((res) => {
      const tx = this.db.transaction('kv', 'readonly');
      const rq = tx.objectStore('kv').get(k);
      rq.onsuccess = () => res(rq.result); rq.onerror = () => res(null);
    });
  }
};

// =====================================================================
//  WCZYTYWANIE PLIKU
// =====================================================================
async function openViaPicker() {
  if (!window.showOpenFilePicker) { $('#fileInput').click(); return; }
  try {
    const [handle] = await window.showOpenFilePicker({
      types: [{ description: 'Plik Excel', accept: { 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'] } }]
    });
    State.fileHandle = handle;
    await IDB.set('fileHandle', handle);
    const file = await handle.getFile();
    await loadFile(file, true);
  } catch (e) {
    if (e && e.name !== 'AbortError') { console.error(e); $('#fileInput').click(); }
  }
}

async function loadFile(file, canSave) {
  try {
    const buf = await file.arrayBuffer();
    const wb = await XLSXEngine.load(buf);
    const model = TrainingModel.build(wb);
    State.wb = wb;
    State.model = model;
    State.fileName = file.name;
    State.canSave = !!canSave;
    $('#sideFileName').textContent = file.name;
    // zapamiętaj snapshot modelu (do szybkiego podglądu bez pliku)
    try { localStorage.setItem('ta_model', JSON.stringify(model)); localStorage.setItem('ta_fname', file.name); } catch (e) {}
    toast('Wczytano: ' + file.name + ' · ' + model.training.weekCount + ' tygodni', '');
    State.selectedExercise = 0;
    State.entryWeek = model.training.weekCount + 1;
    render();
  } catch (e) {
    console.error(e);
    toast('Błąd wczytywania pliku: ' + e.message, 'err');
  }
}

// próba przywrócenia ostatniego pliku przy starcie
async function tryRestore() {
  // 1) uchwyt pliku (pełne możliwości – z zapisem)
  try {
    const handle = await IDB.get('fileHandle');
    if (handle && handle.queryPermission) {
      const perm = await handle.queryPermission({ mode: 'readwrite' });
      if (perm === 'granted') {
        State.fileHandle = handle;
        const file = await handle.getFile();
        await loadFile(file, true);
        return true;
      }
    }
  } catch (e) {}
  // 2) snapshot modelu (tylko podgląd)
  try {
    const snap = localStorage.getItem('ta_model');
    if (snap) {
      State.model = JSON.parse(snap);
      State.fileName = localStorage.getItem('ta_fname') || 'snapshot';
      State.canSave = false;
      State.selectedExercise = 0;
      State.entryWeek = State.model.training.weekCount + 1;
      $('#sideFileName').textContent = State.fileName + ' (podgląd)';
      render();
      return true;
    }
  } catch (e) {}
  return false;
}

// =====================================================================
//  ZAPIS DO PLIKU
// =====================================================================
async function saveFile() {
  if (!State.wb) { toast('Brak wczytanego pliku do zapisu. Otwórz raport przyciskiem 📂.', 'err'); return; }
  try {
    if (State.fileHandle && State.fileHandle.createWritable) {
      const perm = await State.fileHandle.requestPermission({ mode: 'readwrite' });
      if (perm === 'granted') {
        const blob = await XLSXEngine.toBlob(State.wb, 'blob');
        const w = await State.fileHandle.createWritable();
        await w.write(blob); await w.close();
        // odśwież snapshot
        State.model = TrainingModel.build(State.wb);
        try { localStorage.setItem('ta_model', JSON.stringify(State.model)); } catch (e) {}
        toast('💾 Zapisano w pliku: ' + State.fileName, '');
        return;
      }
    }
    // fallback: pobierz nowy plik
    const u8 = await XLSXEngine.toBlob(State.wb, 'uint8array');
    const blob = new Blob([u8], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = (State.fileName || 'trening').replace(/\.xlsx$/i, '') + ' (edytowany).xlsx';
    a.click();
    URL.revokeObjectURL(a.href);
    State.model = TrainingModel.build(State.wb);
    try { localStorage.setItem('ta_model', JSON.stringify(State.model)); } catch (e) {}
    toast('⬇ Pobrano zapisany plik (wyślij go trenerowi).', 'info');
  } catch (e) {
    console.error(e);
    toast('Błąd zapisu: ' + e.message, 'err');
  }
}

// eksport kopii pliku (ta sama forma + nowe dane) — do wysłania trenerowi
async function exportFile() {
  if (!State.wb) { toast('Najpierw otwórz raport (📂), aby go wyeksportować.', 'err'); return; }
  try {
    const u8 = await XLSXEngine.toBlob(State.wb, 'uint8array');
    const blob = new Blob([u8], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = nextWeekFileName(State.fileName);
    a.click();
    URL.revokeObjectURL(a.href);
    toast('⬇ Wyeksportowano plik dla trenera.', 'info');
  } catch (e) { console.error(e); toast('Błąd eksportu: ' + e.message, 'err'); }
}
// zaproponuj nazwę pliku z numerem aktualnego tygodnia
function nextWeekFileName(name) {
  const base = (name || 'Trening').replace(/\.xlsx$/i, '');
  const wk = State.model ? State.model.training.weekCount : null;
  if (wk && /tydzie/i.test(base)) return base.replace(/tydzie[nń]\s*\d+/i, 'tydzień ' + wk) + '.xlsx';
  return base + (wk ? ' tydzień ' + wk : '') + '.xlsx';
}

// =====================================================================
//  WYKRESY — wspólny motyw
// =====================================================================
const COL = {
  acc: '#c6ff3a', mint: '#7af0c8', red: '#ff5d6c', orange: '#ffb648',
  blue: '#5db4ff', violet: '#b78bff', dim: '#9aa6b8', line: '#262d3b', faint: '#616e83'
};
function setupCharts() {
  if (!window.Chart) return;
  Chart.defaults.color = COL.dim;
  Chart.defaults.font.family = "'Inter', sans-serif";
  Chart.defaults.font.size = 11.5;
  Chart.defaults.plugins.legend.labels.usePointStyle = true;
  Chart.defaults.plugins.legend.labels.boxWidth = 8;
  Chart.defaults.plugins.legend.labels.padding = 14;
}
function destroyCharts() { State.charts.forEach(c => { try { c.destroy(); } catch (e) {} }); State.charts = []; }
function mkChart(canvas, cfg) {
  const c = new Chart(canvas, cfg);
  State.charts.push(c);
  return c;
}
const gridCfg = { grid: { color: COL.line, drawBorder: false }, ticks: { color: COL.faint } };
function baseOpts(extra) {
  return Object.assign({
    responsive: true, maintainAspectRatio: false,
    interaction: { mode: 'index', intersect: false },
    plugins: {
      legend: { display: true, position: 'top', align: 'end' },
      tooltip: {
        backgroundColor: '#11151d', borderColor: COL.line, borderWidth: 1,
        titleColor: '#e7ecf3', bodyColor: '#cfd7e3', padding: 11, cornerRadius: 9, displayColors: true
      }
    },
    scales: { x: gridCfg, y: gridCfg }
  }, extra || {});
}
function gradient(ctx, color, a1 = .35, a2 = 0) {
  const g = ctx.createLinearGradient(0, 0, 0, 300);
  g.addColorStop(0, hexA(color, a1)); g.addColorStop(1, hexA(color, a2));
  return g;
}
function hexA(hex, a) {
  const n = parseInt(hex.slice(1), 16);
  return `rgba(${(n >> 16) & 255},${(n >> 8) & 255},${n & 255},${a})`;
}

// =====================================================================
//  ANALITYKA POMOCNICZA
// =====================================================================
function weeklyTonnage(model) {
  const n = model.training.weekCount;
  const arr = new Array(n).fill(0);
  const has = new Array(n).fill(false);
  model.training.exercises.forEach(ex => ex.weeks.forEach((w, i) => {
    if (w.tonnage != null) { arr[i] += w.tonnage; has[i] = true; }
  }));
  return arr.map((v, i) => has[i] ? v : null);
}
function weeklySets(model) {
  const n = model.training.weekCount;
  const arr = new Array(n).fill(0);
  const has = new Array(n).fill(false);
  model.training.exercises.forEach(ex => ex.weeks.forEach((w, i) => {
    if (w.numSets) { arr[i] += w.numSets; has[i] = true; }
  }));
  return arr.map((v, i) => has[i] ? v : null);
}
// najlepszy e1RM ćwiczenia + PR-flaga per tydzień
function withPR(ex) {
  let best = -Infinity;
  return ex.weeks.map(w => {
    let pr = false;
    if (w.e1rm != null && w.e1rm > best + 0.01) { pr = true; best = w.e1rm; }
    return Object.assign({ pr }, w);
  });
}
function parseRepTarget(s) {
  const m = /(\d+)\s*[-–]\s*(\d+)/.exec(String(s || ''));
  if (m) return { min: +m[1], max: +m[2] };
  const one = /(\d+)/.exec(String(s || ''));
  return one ? { min: +one[1], max: +one[1] } : null;
}
function lastNonNull(arr) { for (let i = arr.length - 1; i >= 0; i--) if (arr[i] != null) return arr[i]; return null; }
function firstNonNull(arr) { for (let i = 0; i < arr.length; i++) if (arr[i] != null) return arr[i]; return null; }

// główne boje (do porównania siły)
function mainLifts(model) {
  const want = [
    { re: /wyciskanie sztang/i, name: 'Ławka (sztanga)' },
    { re: /hack squat/i, name: 'Hack squat' },
    { re: /rumu.ski|martwy ci/i, name: 'RDL' },
    { re: /leg press/i, name: 'Leg press' },
    { re: /t-bar|wios.owanie t/i, name: 'Wiosło T-bar' }
  ];
  const out = [];
  want.forEach(w => {
    const ex = model.training.exercises.find(e => w.re.test(e.name));
    if (ex) out.push({ name: w.name, ex });
  });
  return out;
}

// =====================================================================
//  RENDER — router
// =====================================================================
const VIEW_META = {
  dashboard: ['Pulpit', 'Przegląd całej współpracy — siła, sylwetka, regeneracja.'],
  strength: ['Siła & progresja', 'Progresja obciążeń, objętości i szacowanego 1RM dla każdego ćwiczenia.'],
  body: ['Sylwetka & pomiary', 'Waga, obwody, kalorie, białko, sen i kroki w czasie.'],
  reports: ['Raporty tygodniowe', 'Twoje cotygodniowe odpowiedzi i samopoczucie.'],
  entry: ['Wpis tygodnia', 'Uzupełnij wykonane serie i zapisz je z powrotem do pliku dla trenera.'],
  rpe: ['Skala RPE', 'Punkt odniesienia dla zapasu powtórzeń (RIR).']
};

function render() {
  destroyCharts();
  const [title, sub] = VIEW_META[State.view] || ['', ''];
  $('#viewTitle').textContent = title;
  $('#viewSub').textContent = sub;
  $$('.nav-item').forEach(n => n.classList.toggle('active', n.dataset.view === State.view));
  const root = $('#viewRoot');
  root.innerHTML = '';

  if (!State.model) { renderWelcome(root); return; }

  switch (State.view) {
    case 'dashboard': renderDashboard(root); break;
    case 'strength': renderStrength(root); break;
    case 'body': renderBody(root); break;
    case 'reports': renderReports(root); break;
    case 'entry': renderEntry(root); break;
    case 'rpe': renderRPE(root); break;
  }
}

// ---------------------------------------------------------- ekran startowy
function renderWelcome(root) {
  const dz = el('div', { class: 'dropzone', id: 'dropzone' });
  dz.innerHTML = `
    <div class="big">📂</div>
    <h2>Wczytaj swój raport treningowy</h2>
    <p>Przeciągnij tu plik <b>.xlsx</b> od trenera lub kliknij „Otwórz raport”.<br>
    Aplikacja czyta wszystkie tygodnie — także te w kolumnach ukrytych przez trenera.</p>
    <div style="margin-top:22px"><button class="btn primary" id="dzBtn"><span class="ico">📂</span> Otwórz raport</button></div>
    <p style="margin-top:26px;font-size:12px;color:var(--txt-faint)">
      Wskazówka: użyj przycisku „Otwórz raport” (Chrome/Edge), aby aplikacja mogła też <b>zapisywać</b> wpisy z powrotem do tego samego pliku.
    </p>`;
  root.appendChild(dz);
  $('#dzBtn').addEventListener('click', openViaPicker);
  setupDropzone(dz);
}
function setupDropzone(dz) {
  ['dragenter', 'dragover'].forEach(ev => dz.addEventListener(ev, e => { e.preventDefault(); dz.classList.add('drag'); }));
  ['dragleave', 'drop'].forEach(ev => dz.addEventListener(ev, e => { e.preventDefault(); dz.classList.remove('drag'); }));
  dz.addEventListener('drop', e => {
    const f = e.dataTransfer.files[0];
    if (f) loadFile(f, false);
  });
}

// =====================================================================
//  WIDOK: PULPIT
// =====================================================================
function renderDashboard(root) {
  const M = State.model, T = M.training, MET = M.metrics;
  const wl = weekLabels(T.weekCount);

  // --- KPI ---
  const waga = MET ? MET.wagaAvg : [];
  const wagaStart = MET ? firstNonNull(waga) : null;
  const wagaNow = MET ? lastNonNull(waga) : null;
  const wagaDelta = (wagaStart != null && wagaNow != null) ? (wagaNow - wagaStart) : null;
  const bench = T.exercises.find(e => /wyciskanie sztang/i.test(e.name));
  const benchBest = bench ? Math.max(...bench.weeks.map(w => w.e1rm || 0)) : 0;
  const ton = weeklyTonnage(M);
  const tonTotal = ton.reduce((a, b) => a + (b || 0), 0);
  const krokiAvg = MET ? avg(MET.steps) : null;
  const senAvg = MET ? avg(MET.senAvg) : null;
  const kcalAvg = MET ? avg(MET.kcalAvg) : null;

  const kpi = el('div', { class: 'kpis' });
  kpi.innerHTML = `
    ${kpiCard('Aktualny tydzień', 'T' + T.weekCount, T.days.length + ' treningi/tydz.', 'neutral')}
    ${kpiCard('Waga teraz', wagaNow != null ? fmt(wagaNow, 1) : '—', 'kg',
      wagaDelta != null ? (wagaDelta < 0 ? 'down-good' : 'up') : 'neutral',
      wagaDelta != null ? (wagaDelta > 0 ? '+' : '') + fmt(wagaDelta, 1) + ' kg od startu' : '')}
    ${kpiCard('Ławka — szac. 1RM', benchBest ? fmt(benchBest, 0) : '—', 'kg', 'up')}
    ${kpiCard('Łączny tonaż', fmt(Math.round(tonTotal / 1000), 0), 't przerzucone', 'up')}
    ${kpiCard('Śr. kroki / dzień', krokiAvg ? fmt(Math.round(krokiAvg), 0) : '—', '', 'neutral')}
    ${kpiCard('Śr. sen', senAvg ? fmt(senAvg, 1) : '—', 'h / noc', senAvg && senAvg < 6.5 ? 'down' : 'neutral')}
  `;
  root.appendChild(kpi);

  // --- waga + pas (redukcja) ---
  const g1 = el('div', { class: 'grid', style: 'grid-template-columns: 1.6fr 1fr; margin-top:16px' });
  g1.appendChild(cardWithChart('Redukcja w czasie', 'Waga (średnia tyg.) i obwód pasa', 'h-lg', 'chWaga'));
  g1.appendChild(cardWithChart('Tonaż treningowy', 'Suma kg × powtórzenia na tydzień', 'h-lg', 'chTon'));
  root.appendChild(g1);

  // --- main lifts e1rm ---
  const g2 = el('div', { class: 'grid', style: 'grid-template-columns: 1fr; margin-top:16px' });
  g2.appendChild(cardWithChart('Progresja siły — główne boje', 'Szacowany 1RM (Epley) głównych ćwiczeń', 'h-lg', 'chLifts'));
  root.appendChild(g2);

  // --- małe metryki ---
  if (MET) {
    const g3 = el('div', { class: 'grid', style: 'grid-template-columns: repeat(auto-fit,minmax(280px,1fr)); margin-top:16px' });
    g3.appendChild(cardWithChart('Kroki / dzień', 'Średnia tygodniowa', 'h-md', 'chKroki'));
    g3.appendChild(cardWithChart('Kalorie & białko', 'Średnia dzienna', 'h-md', 'chKcal'));
    g3.appendChild(cardWithChart('Sen', 'Średnia godzin / noc', 'h-md', 'chSen'));
    root.appendChild(g3);
  }

  // --- render charts ---
  if (MET) {
    const ctx = $('#chWaga').getContext('2d');
    mkChart($('#chWaga'), {
      type: 'line',
      data: {
        labels: weekLabels(MET.weeksFilled),
        datasets: [
          { label: 'Waga (kg)', data: MET.wagaAvg.slice(0, MET.weeksFilled), borderColor: COL.acc, backgroundColor: gradient(ctx, COL.acc, .25), fill: true, tension: .35, borderWidth: 2.5, pointRadius: 2, pointHoverRadius: 5, yAxisID: 'y' },
          { label: 'Pas (cm)', data: MET.pas.slice(0, MET.weeksFilled), borderColor: COL.orange, backgroundColor: 'transparent', fill: false, tension: .35, borderWidth: 2, pointRadius: 2, spanGaps: true, yAxisID: 'y1', borderDash: [4, 3] }
        ]
      },
      options: baseOpts({
        scales: {
          x: gridCfg,
          y: Object.assign({ position: 'left', title: { display: true, text: 'kg' } }, gridCfg),
          y1: Object.assign({ position: 'right', grid: { drawOnChartArea: false }, ticks: { color: COL.orange }, title: { display: true, text: 'cm' } })
        }
      })
    });
  }

  const ctxT = $('#chTon').getContext('2d');
  mkChart($('#chTon'), {
    type: 'bar',
    data: { labels: wl, datasets: [{ label: 'Tonaż (kg)', data: ton, backgroundColor: hexA(COL.blue, .55), borderColor: COL.blue, borderWidth: 1, borderRadius: 5 }] },
    options: baseOpts({ plugins: { legend: { display: false } } })
  });

  // main lifts
  const lifts = mainLifts(M);
  const liftColors = [COL.acc, COL.violet, COL.orange, COL.blue, COL.mint];
  mkChart($('#chLifts'), {
    type: 'line',
    data: {
      labels: wl,
      datasets: lifts.map((l, i) => ({
        label: l.name,
        data: l.ex.weeks.map(w => w.e1rm != null ? +w.e1rm.toFixed(1) : null),
        borderColor: liftColors[i % liftColors.length],
        backgroundColor: 'transparent', tension: .3, borderWidth: 2.4, pointRadius: 2, pointHoverRadius: 5, spanGaps: true
      }))
    },
    options: baseOpts({ scales: { x: gridCfg, y: Object.assign({ title: { display: true, text: 'szac. 1RM (kg)' } }, gridCfg) } })
  });

  if (MET) {
    const goalSteps = 10000;
    mkChart($('#chKroki'), {
      type: 'bar',
      data: { labels: weekLabels(MET.weeksFilled), datasets: [{ label: 'Kroki', data: MET.steps.slice(0, MET.weeksFilled), backgroundColor: MET.steps.slice(0, MET.weeksFilled).map(v => v >= goalSteps ? hexA(COL.acc, .6) : hexA(COL.dim, .4)), borderRadius: 4 }] },
      options: baseOpts({
        plugins: { legend: { display: false }, tooltip: Chart.defaults.plugins.tooltip,
          annotation: undefined },
        scales: { x: gridCfg, y: Object.assign({ suggestedMin: 7000 }, gridCfg) }
      })
    });
    const ctxK = $('#chKcal').getContext('2d');
    mkChart($('#chKcal'), {
      type: 'bar',
      data: {
        labels: weekLabels(MET.weeksFilled),
        datasets: [
          { type: 'bar', label: 'Kalorie', data: MET.kcalAvg.slice(0, MET.weeksFilled), backgroundColor: hexA(COL.orange, .45), borderRadius: 4, yAxisID: 'y' },
          { type: 'line', label: 'Białko (g)', data: MET.bialkoAvg.slice(0, MET.weeksFilled), borderColor: COL.mint, backgroundColor: 'transparent', tension: .3, borderWidth: 2, pointRadius: 2, yAxisID: 'y1' }
        ]
      },
      options: baseOpts({
        scales: {
          x: gridCfg,
          y: Object.assign({ position: 'left', ticks: { color: COL.orange } }, gridCfg),
          y1: Object.assign({ position: 'right', grid: { drawOnChartArea: false }, ticks: { color: COL.mint }, suggestedMin: 150, suggestedMax: 260 })
        }
      })
    });
    const ctxS = $('#chSen').getContext('2d');
    mkChart($('#chSen'), {
      type: 'line',
      data: { labels: weekLabels(MET.weeksFilled), datasets: [{ label: 'Sen (h)', data: MET.senAvg.slice(0, MET.weeksFilled), borderColor: COL.violet, backgroundColor: gradient(ctxS, COL.violet, .25), fill: true, tension: .35, borderWidth: 2.2, pointRadius: 2 }] },
      options: baseOpts({ plugins: { legend: { display: false } }, scales: { x: gridCfg, y: Object.assign({ suggestedMin: 4, suggestedMax: 9 }, gridCfg) } })
    });
  }
}
function avg(arr) { const v = (arr || []).filter(x => x != null); return v.length ? v.reduce((a, b) => a + b, 0) / v.length : null; }
function kpiCard(label, val, unit, deltaType, deltaText) {
  let cls = 'neutral', txt = deltaText || '';
  if (deltaType === 'down-good') cls = 'up';   // spadek wagi = dobrze (zielony)
  else if (deltaType === 'up') cls = 'up';
  else if (deltaType === 'down') cls = 'down';
  return `<div class="kpi">
    <div class="k-label">${esc(label)}</div>
    <div class="k-val">${esc(val)} ${unit ? '<small>' + esc(unit) + '</small>' : ''}</div>
    ${txt ? `<div class="k-delta ${cls}">${esc(txt)}</div>` : ''}
  </div>`;
}
function cardWithChart(title, sub, h, canvasId) {
  const c = el('div', { class: 'card' });
  c.innerHTML = `<div class="card-head"><div><h3>${esc(title)}</h3>${sub ? `<div class="card-sub" style="margin:0">${esc(sub)}</div>` : ''}</div></div>
    <div class="chart-wrap ${h}"><canvas id="${canvasId}"></canvas></div>`;
  return c;
}

// =====================================================================
//  WIDOK: SIŁA
// =====================================================================
function renderStrength(root) {
  const M = State.model, T = M.training;
  // chipsy ćwiczeń pogrupowane wg dnia
  const chips = el('div', { class: 'card', style: 'margin-bottom:16px' });
  let chipsHtml = '';
  T.days.forEach(d => {
    chipsHtml += `<div class="section-title" style="margin:8px 0 8px">${esc(d.label)}</div><div class="chips">`;
    d.exercises.forEach(ex => {
      const idx = T.exercises.indexOf(ex);
      chipsHtml += `<div class="chip ${idx === State.selectedExercise ? 'active' : ''}" data-ex="${idx}">${esc(ex.name)}</div>`;
    });
    chipsHtml += '</div>';
  });
  chips.innerHTML = chipsHtml;
  root.appendChild(chips);
  $$('.chip', chips).forEach(c => c.addEventListener('click', () => { State.selectedExercise = +c.dataset.ex; render(); }));

  const ex = T.exercises[State.selectedExercise];
  if (!ex) return;
  const weeks = withPR(ex);
  const wl = weekLabels(T.weekCount);

  // rekordy
  const e1 = weeks.map(w => w.e1rm).filter(x => x != null);
  const loads = weeks.map(w => w.eff).filter(x => x != null);
  const tons = weeks.map(w => w.tonnage).filter(x => x != null);
  const bestE1 = e1.length ? Math.max(...e1) : null;
  const maxLoad = loads.length ? Math.max(...loads) : null;
  const maxTon = tons.length ? Math.max(...tons) : null;
  const firstLoad = firstNonNull(weeks.map(w => w.eff));
  const lastLoad = lastNonNull(weeks.map(w => w.eff));
  const loadGain = (firstLoad != null && lastLoad != null) ? lastLoad - firstLoad : null;

  const recs = el('div', { class: 'kpis', style: 'margin-bottom:16px' });
  recs.innerHTML = `
    ${kpiCard('Rekord szac. 1RM', bestE1 ? fmt(bestE1, 1) : '—', 'kg', 'up')}
    ${kpiCard('Max ciężar roboczy', maxLoad ? fmt(maxLoad, maxLoad % 1 ? 1 : 0) : '—', ex.weeks.some(w => w.weight && w.weight.perHand) ? 'kg (łącznie)' : 'kg', 'up')}
    ${kpiCard('Najlepszy tonaż', maxTon ? fmt(Math.round(maxTon), 0) : '—', 'kg / trening', 'up')}
    ${kpiCard('Przyrost ciężaru', loadGain != null ? (loadGain >= 0 ? '+' : '') + fmt(loadGain, loadGain % 1 ? 1 : 0) : '—', 'kg', loadGain >= 0 ? 'up' : 'down')}
  `;
  root.appendChild(recs);

  // plan info
  const info = el('div', { class: 'card', style: 'margin-bottom:16px' });
  info.innerHTML = `<div class="row" style="gap:22px">
    <div><div class="card-sub" style="margin:0">Partie</div><b>${esc(ex.muscles || '—')}</b></div>
    <div><div class="card-sub" style="margin:0">Cel powtórzeń</div><b>${esc(ex.repsTarget || '—')}</b></div>
    <div><div class="card-sub" style="margin:0">RIR (zapas)</div><b>${esc(ex.rir || '—')}</b></div>
    <div><div class="card-sub" style="margin:0">Tempo</div><b>${esc(ex.tempo || '—')}</b></div>
    <div><div class="card-sub" style="margin:0">Przerwy</div><b>${esc(ex.rest || '—')}</b></div>
    <div><div class="card-sub" style="margin:0">Plan ciężaru</div><b>${esc(ex.planWeight || '—')}</b></div>
  </div>`;
  root.appendChild(info);

  // ostrzeżenie o możliwej zmianie ćwiczenia (zmiana typu obciążenia w trakcie)
  if (ex.swapHints && ex.swapHints.length) {
    const kindPl = { bb: 'sztanga', db: 'hantle', machine: 'maszyna' };
    const wks = ex.swapHints.map(h => `T${h.week} (${kindPl[h.from] || h.from}→${kindPl[h.to] || h.to})`).join(', ');
    const note = el('div', { class: 'card', style: 'margin-bottom:16px;border-color:var(--orange)' });
    note.innerHTML = `<div class="row" style="gap:10px">
      <span class="pill orange">⚠ Uwaga</span>
      <div style="font-size:12.5px;color:var(--txt-dim)">
        Wykryto zmianę typu obciążenia: <b>${esc(wks)}</b>. Jeśli trener podmienił tu ćwiczenie,
        wcześniejsze tygodnie w tym wierszu mogą dotyczyć poprzedniego ćwiczenia — wykres pokazuje to miejsce w planie, a nazwa jest aktualna.
      </div></div>`;
    root.appendChild(note);
  }

  // główny wykres
  const mainCard = cardWithChart('Progresja: ' + ex.name, 'Ciężar efektywny, szacowany 1RM i tonaż w kolejnych tygodniach', 'h-xl', 'chEx');
  root.appendChild(mainCard);

  const ctx = $('#chEx').getContext('2d');
  mkChart($('#chEx'), {
    data: {
      labels: wl,
      datasets: [
        { type: 'bar', label: 'Tonaż (kg)', data: weeks.map(w => w.tonnage != null ? Math.round(w.tonnage) : null), backgroundColor: hexA(COL.blue, .30), borderColor: hexA(COL.blue, .5), borderWidth: 1, borderRadius: 4, yAxisID: 'y1', order: 3 },
        { type: 'line', label: 'Ciężar efekt. (kg)', data: weeks.map(w => w.eff), borderColor: COL.acc, backgroundColor: gradient(ctx, COL.acc, .18), fill: true, tension: .3, borderWidth: 2.6, pointRadius: 3, pointHoverRadius: 6, spanGaps: true, yAxisID: 'y', order: 1 },
        { type: 'line', label: 'Szac. 1RM (kg)', data: weeks.map(w => w.e1rm != null ? +w.e1rm.toFixed(1) : null), borderColor: COL.violet, backgroundColor: 'transparent', borderDash: [5, 4], tension: .3, borderWidth: 2, pointRadius: weeks.map(w => w.pr ? 5 : 0), pointBackgroundColor: COL.violet, spanGaps: true, yAxisID: 'y', order: 2 }
      ]
    },
    options: baseOpts({
      scales: {
        x: gridCfg,
        y: Object.assign({ position: 'left', title: { display: true, text: 'kg' } }, gridCfg),
        y1: Object.assign({ position: 'right', grid: { drawOnChartArea: false }, ticks: { color: COL.blue }, title: { display: true, text: 'tonaż' } })
      }
    })
  });

  // tabela serii — kolory wg Twoich oznaczeń RIR w pliku
  const tcard = el('div', { class: 'card', style: 'margin-top:16px' });
  let rows = '';
  weeks.forEach(w => {
    if (!w.hasData) return;
    const setCells = [0, 1, 2, 3].map(i => {
      const s = w.sets[i];
      if (!s) return '<td class="muted">·</td>';
      const color = colorForCat(s.color);
      const style = color ? `style="color:${color};font-weight:700"` : '';
      const tip = s.top ? ' title="Top-set (cięższa seria)"' : '';
      return `<td ${style}${tip}>${esc(s.raw)}${s.top ? ' ▲' : ''}</td>`;
    }).join('');
    rows += `<tr>
      <td class="l"><b>T${w.week}</b> ${w.pr ? '<span class="pr-star" title="Rekord 1RM">★</span>' : ''}</td>
      <td>${w.weight ? esc(w.weight.raw) : '—'}</td>
      ${setCells}
      <td>${w.totalReps || '—'}</td>
      <td>${w.tonnage != null ? fmt(Math.round(w.tonnage)) : '—'}</td>
      <td><b>${w.e1rm != null ? fmt(w.e1rm, 1) : '—'}</b></td>
    </tr>`;
  });
  tcard.innerHTML = `<div class="card-head"><h3>Szczegóły tygodni</h3>
    <div class="legend-row">
      <span class="lg"><span class="sw" style="background:#92D050"></span>zapas 2–4</span>
      <span class="lg"><span class="sw" style="background:#FFC000"></span>zapas 1–2</span>
      <span class="lg"><span class="sw" style="background:#FF0000"></span>nieudana</span>
      <span class="lg">▲ top-set</span></div></div>
    <div class="tbl-scroll"><table class="tbl">
      <thead><tr><th class="l">Tydzień</th><th>Ciężar</th><th>S1</th><th>S2</th><th>S3</th><th>S4</th><th>Σ powt.</th><th>Tonaż</th><th>1RM</th></tr></thead>
      <tbody>${rows || '<tr><td colspan="9" class="muted l">Brak danych</td></tr>'}</tbody>
    </table></div>`;
  root.appendChild(tcard);
}
function colorForCat(c) { return c === 'green' ? '#92D050' : c === 'orange' ? '#FFC000' : c === 'red' ? '#FF5d6c' : null; }

// =====================================================================
//  WIDOK: SYLWETKA
// =====================================================================
function renderBody(root) {
  const M = State.model, MET = M.metrics;
  if (!MET) { root.appendChild(el('div', { class: 'empty-state' }, '<div class="big">📐</div>Brak arkusza pomiarów w pliku.')); return; }
  const n = MET.weeksFilled;
  const wl = weekLabels(n);

  // KPI sylwetki
  const dWaga = delta(MET.wagaAvg), dPas = delta(MET.pas), dKlatka = delta(MET.klatka), dBiceps = delta(MET.biceps);
  const kp = el('div', { class: 'kpis', style: 'margin-bottom:16px' });
  kp.innerHTML = `
    ${kpiCard('Waga', fmt(lastNonNull(MET.wagaAvg), 1), 'kg', dWaga < 0 ? 'up' : 'down', deltaTxt(dWaga, 'kg'))}
    ${kpiCard('Pas', fmt(lastNonNull(MET.pas), 0), 'cm', dPas < 0 ? 'up' : 'down', deltaTxt(dPas, 'cm'))}
    ${kpiCard('Klatka', fmt(lastNonNull(MET.klatka), 0), 'cm', 'neutral', deltaTxt(dKlatka, 'cm'))}
    ${kpiCard('Biceps', fmt(lastNonNull(MET.biceps), 1), 'cm', 'neutral', deltaTxt(dBiceps, 'cm'))}
  `;
  root.appendChild(kp);

  const g1 = el('div', { class: 'grid', style: 'grid-template-columns:1fr' });
  g1.appendChild(cardWithChart('Waga — pełny obraz', 'Średnia tygodniowa', 'h-lg', 'chBodyWaga'));
  root.appendChild(g1);

  const g2 = el('div', { class: 'grid', style: 'grid-template-columns:1fr; margin-top:16px' });
  g2.appendChild(cardWithChart('Obwody ciała', 'Udo, pas, biodra, klatka, biceps (cm)', 'h-lg', 'chObwody'));
  root.appendChild(g2);

  const g3 = el('div', { class: 'grid', style: 'grid-template-columns:repeat(auto-fit,minmax(300px,1fr)); margin-top:16px' });
  g3.appendChild(cardWithChart('Kalorie & białko', 'Średnia dzienna', 'h-md', 'chBodyKcal'));
  g3.appendChild(cardWithChart('Sen', 'Średnia h / noc', 'h-md', 'chBodySen'));
  g3.appendChild(cardWithChart('Kroki', 'Średnia / dzień', 'h-md', 'chBodyKroki'));
  root.appendChild(g3);

  // waga
  const ctx = $('#chBodyWaga').getContext('2d');
  mkChart($('#chBodyWaga'), {
    type: 'line',
    data: { labels: wl, datasets: [{ label: 'Waga (kg)', data: MET.wagaAvg.slice(0, n), borderColor: COL.acc, backgroundColor: gradient(ctx, COL.acc, .28), fill: true, tension: .35, borderWidth: 2.6, pointRadius: 3, pointHoverRadius: 6 }] },
    options: baseOpts({ plugins: { legend: { display: false } } })
  });

  // obwody
  const obw = [
    { label: 'Udo', data: MET.udo, c: COL.acc },
    { label: 'Pas', data: MET.pas, c: COL.orange },
    { label: 'Biodra', data: MET.biodra, c: COL.blue },
    { label: 'Klatka', data: MET.klatka, c: COL.violet },
    { label: 'Biceps', data: MET.biceps, c: COL.mint }
  ];
  mkChart($('#chObwody'), {
    type: 'line',
    data: { labels: wl, datasets: obw.map(o => ({ label: o.label, data: o.data.slice(0, n), borderColor: o.c, backgroundColor: 'transparent', tension: .3, borderWidth: 2, pointRadius: 2, spanGaps: true })) },
    options: baseOpts({})
  });

  // kcal & białko
  mkChart($('#chBodyKcal'), {
    data: {
      labels: wl,
      datasets: [
        { type: 'bar', label: 'Kalorie', data: MET.kcalAvg.slice(0, n), backgroundColor: hexA(COL.orange, .45), borderRadius: 4, yAxisID: 'y' },
        { type: 'line', label: 'Białko (g)', data: MET.bialkoAvg.slice(0, n), borderColor: COL.mint, tension: .3, borderWidth: 2, pointRadius: 2, yAxisID: 'y1' }
      ]
    },
    options: baseOpts({ scales: { x: gridCfg, y: Object.assign({ ticks: { color: COL.orange } }, gridCfg), y1: Object.assign({ position: 'right', grid: { drawOnChartArea: false }, ticks: { color: COL.mint }, suggestedMin: 150, suggestedMax: 280 }) } })
  });
  const ctxS = $('#chBodySen').getContext('2d');
  mkChart($('#chBodySen'), {
    type: 'line', data: { labels: wl, datasets: [{ label: 'Sen', data: MET.senAvg.slice(0, n), borderColor: COL.violet, backgroundColor: gradient(ctxS, COL.violet, .22), fill: true, tension: .35, borderWidth: 2.2, pointRadius: 2 }] },
    options: baseOpts({ plugins: { legend: { display: false } }, scales: { x: gridCfg, y: Object.assign({ suggestedMin: 4, suggestedMax: 9 }, gridCfg) } })
  });
  mkChart($('#chBodyKroki'), {
    type: 'bar', data: { labels: wl, datasets: [{ label: 'Kroki', data: MET.steps.slice(0, n), backgroundColor: MET.steps.slice(0, n).map(v => v >= 10000 ? hexA(COL.acc, .6) : hexA(COL.dim, .4)), borderRadius: 4 }] },
    options: baseOpts({ plugins: { legend: { display: false } }, scales: { x: gridCfg, y: Object.assign({ suggestedMin: 7000 }, gridCfg) } })
  });
}
function delta(arr) { const f = firstNonNull(arr), l = lastNonNull(arr); return (f != null && l != null) ? l - f : null; }
function deltaTxt(d, unit) { return d == null ? '' : (d > 0 ? '+' : '') + fmt(d, Math.abs(d) % 1 ? 1 : 0) + ' ' + unit + ' od startu'; }

// =====================================================================
//  WIDOK: RAPORTY
// =====================================================================
function renderReports(root) {
  const M = State.model;
  const qr = M.questions.raport, qt = M.questions.treningowe;

  // mini-wykresy: ocena planu (treningowe D=indeks 2) i głód (raport G=ostatni)
  const ratingIdx = qt ? qt.headers.findIndex(h => /oceniasz.*plan/i.test(h.text)) : -1;
  const hungerIdx = qr ? qr.headers.findIndex(h => /g.odem|g.łod/i.test(h.text)) : (qr ? qr.headers.length - 1 : -1);

  const charts = el('div', { class: 'grid', style: 'grid-template-columns:1fr 1fr; margin-bottom:16px' });
  charts.appendChild(cardWithChart('Ocena planu treningowego', 'Skala 1–10 wg Twoich raportów', 'h-md', 'chRating'));
  charts.appendChild(cardWithChart('Poziom głodu', 'Skala 0–6 (0 = brak, 6 = ciągły)', 'h-md', 'chHunger'));
  root.appendChild(charts);

  // accordion raportów
  const wrap = el('div');
  const allWeeks = new Set();
  if (qr) qr.rows.forEach(r => allWeeks.add(r.week));
  if (qt) qt.rows.forEach(r => allWeeks.add(r.week));
  const sorted = Array.from(allWeeks).sort((a, b) => b - a);

  sorted.forEach((wk, i) => {
    const rRow = qr ? qr.rows.find(r => r.week === wk) : null;
    const tRow = qt ? qt.rows.find(r => r.week === wk) : null;
    const box = el('div', { class: 'report-week' + (i === 0 ? ' open' : '') });
    let body = '';
    if (tRow) {
      qt.headers.forEach((h, j) => { if (tRow.answers[j]) body += qaHtml(h.text, tRow.answers[j]); });
    }
    if (rRow) {
      qr.headers.forEach((h, j) => { if (rRow.answers[j]) body += qaHtml(h.text, rRow.answers[j]); });
    }
    box.innerHTML = `<div class="rw-head"><span class="wk">Tydzień ${wk}</span>
      <span class="pill gray">${(tRow ? 1 : 0) + (rRow ? 1 : 0)} sekcje</span></div>
      <div class="rw-body">${body || '<div class="qa"><div class="a empty">Brak odpowiedzi</div></div>'}</div>`;
    box.querySelector('.rw-head').addEventListener('click', () => box.classList.toggle('open'));
    wrap.appendChild(box);
  });
  root.appendChild(wrap);

  // wykresy
  if (qt && ratingIdx >= 0) {
    const labels = qt.rows.map(r => 'T' + r.week);
    const data = qt.rows.map(r => { const v = parseFloat(String(r.answers[ratingIdx]).replace(',', '.')); return isNaN(v) ? null : v; });
    mkChart($('#chRating'), {
      type: 'line', data: { labels, datasets: [{ label: 'Ocena', data, borderColor: COL.acc, backgroundColor: 'transparent', tension: .3, borderWidth: 2.4, pointRadius: 3, spanGaps: true }] },
      options: baseOpts({ plugins: { legend: { display: false } }, scales: { x: gridCfg, y: Object.assign({ suggestedMin: 6, suggestedMax: 10 }, gridCfg) } })
    });
  }
  if (qr && hungerIdx >= 0) {
    const labels = qr.rows.map(r => 'T' + r.week);
    const data = qr.rows.map(r => { const v = parseFloat(String(r.answers[hungerIdx]).replace(',', '.')); return isNaN(v) ? null : v; });
    mkChart($('#chHunger'), {
      type: 'bar', data: { labels, datasets: [{ label: 'Głód', data, backgroundColor: data.map(v => v >= 4 ? hexA(COL.red, .6) : v >= 3 ? hexA(COL.orange, .6) : hexA(COL.mint, .55)), borderRadius: 4 }] },
      options: baseOpts({ plugins: { legend: { display: false } }, scales: { x: gridCfg, y: Object.assign({ suggestedMin: 0, suggestedMax: 6 }, gridCfg) } })
    });
  }
}
function qaHtml(q, a) {
  return `<div class="qa"><div class="q">${esc(q)}</div><div class="a">${esc(a)}</div></div>`;
}

// =====================================================================
//  WIDOK: WPIS TYGODNIA
// =====================================================================
function renderEntry(root) {
  const M = State.model, T = M.training;
  const canSave = State.wb != null;

  // pasek: wybór tygodnia
  const bar = el('div', { class: 'card', style: 'margin-bottom:16px' });
  const maxWeek = T.weekCount + 1;
  let opts = '';
  for (let w = 1; w <= maxWeek; w++) {
    const isNew = w > T.weekCount;
    opts += `<option value="${w}" ${w === State.entryWeek ? 'selected' : ''}>Tydzień ${w}${isNew ? ' (nowy)' : ''}</option>`;
  }
  bar.innerHTML = `<div class="row">
    <label class="fld" style="min-width:200px">Wpisywany tydzień
      <select id="entryWeekSel">${opts}</select></label>
    <div class="spacer"></div>
    ${canSave ? '' : '<span class="pill orange">Tryb podglądu — otwórz plik przyciskiem 📂, aby zapisywać</span>'}
    <button class="btn primary" id="btnSaveEntry" ${canSave ? '' : 'disabled'}><span class="ico">💾</span> Zapisz do pliku</button>
  </div>
  <div class="card-sub" style="margin-top:10px;margin-bottom:0">
    Wpisz ciężar i wykonane powtórzenia. Pod każdą serią wybierz kolor wysiłku (jak w pliku):
    <span class="lg" style="display:inline-flex;gap:5px;align-items:center;margin:0 4px"><span class="sw" style="background:#92D050"></span>zapas 2–4</span>
    <span class="lg" style="display:inline-flex;gap:5px;align-items:center;margin:0 4px"><span class="sw" style="background:#FFC000"></span>zapas 1–2</span>
    <span class="lg" style="display:inline-flex;gap:5px;align-items:center;margin:0 4px"><span class="sw" style="background:#FF0000"></span>nieudana</span>.
    Puste pola pomijane. Zapis trafia do tych samych komórek z zachowaniem formatowania.
  </div>`;
  root.appendChild(bar);
  $('#entryWeekSel').addEventListener('change', e => { State.entryWeek = +e.target.value; render(); });

  const w = State.entryWeek;
  const wIdx = w - 1;

  // formularze per dzień
  T.days.forEach(day => {
    const dayBox = el('div', { class: 'entry-day' });
    dayBox.innerHTML = `<div class="ed-head"><span class="dot"></span><h3>${esc(day.label)}</h3></div>`;
    day.exercises.forEach(ex => {
      const cur = ex.weeks[wIdx]; // może nie istnieć dla nowego tygodnia
      const prev = wIdx > 0 ? ex.weeks[wIdx - 1] : null;
      const wc = TrainingModel.WEIGHT_COL0 + (w - 1) * TrainingModel.WEEK_STRIDE;
      const weightRef = TrainingModel.numToCol(wc) + ex.row;
      const setRefs = [1, 2, 3, 4].map(s => TrainingModel.numToCol(wc + s) + ex.row);

      const presc = ex.prescription.filter(x => x != null).join(' / ');
      const box = el('div', { class: 'entry-ex' });
      box.innerHTML = `
        <div class="ee-top">
          <div><span class="ee-name">${esc(ex.name)}</span> <span class="ee-meta">· ${esc(ex.muscles)}</span></div>
          <div class="ee-plan">
            <span>Cel: <b>${esc(ex.repsTarget || '—')}</b></span>
            <span>RIR: <b>${esc(ex.rir || '—')}</b></span>
            ${presc ? `<span>Rozpiska: <b>${esc(presc)}</b></span>` : ''}
            ${ex.planWeight ? `<span>Plan: <b>${esc(ex.planWeight)}</b></span>` : ''}
          </div>
        </div>
        <div class="entry-inputs">
          <div><div class="mini-label">Ciężar</div><input type="text" data-ref="${weightRef}" value="${cur && cur.weight ? esc(cur.weight.raw) : ''}" placeholder="${prev && prev.weight ? esc(prev.weight.raw) : 'np. 90kg'}"></div>
          ${[0, 1, 2, 3].map(s => {
            const sv = cur && cur.sets[s] ? cur.sets[s].raw : '';
            const pv = prev && prev.sets[s] ? prev.sets[s].raw : '';
            const col = cur && cur.sets[s] ? cur.sets[s].color : null;
            return setCellHtml(setRefs[s], s + 1, sv, pv, col);
          }).join('')}
          <div class="set-cell"><div class="mini-label">Notatka / odczucia</div><input type="text" data-ref="${TrainingModel.numToCol(5) + ex.row}" placeholder="opcjonalnie"></div>
        </div>
        ${prev && prev.hasData ? `<div class="prev-hint">Poprzedni tydzień (T${w - 1}): ${prev.weight ? esc(prev.weight.raw) : '—'} × [${prev.sets.map(s => s.raw).join(', ')}]</div>` : ''}
      `;
      dayBox.appendChild(box);
    });
    root.appendChild(dayBox);
  });

  bindColorDots(root);
  if (canSave) $('#btnSaveEntry').addEventListener('click', () => commitEntry(root));
}

// pojedyncza komórka serii z wyborem koloru (RIR)
function tintClass(c) { return c === 'green' ? 'tint-green' : c === 'orange' ? 'tint-orange' : c === 'red' ? 'tint-red' : ''; }
function setCellHtml(ref, n, value, placeholder, color) {
  return `<div class="set-cell">
    <div class="mini-label">Seria ${n}</div>
    <input type="text" data-ref="${ref}" data-color="${color || ''}" class="${tintClass(color)}" value="${esc(value)}" placeholder="${placeholder !== '' ? esc(placeholder) : '–'}">
    <div class="cdots" title="Oznacz wysiłek serii (kolor jak w pliku)">
      <span class="cdot g ${color === 'green' ? 'active' : ''}" data-c="green" title="Zielony — zapas 2–4 powt."></span>
      <span class="cdot o ${color === 'orange' ? 'active' : ''}" data-c="orange" title="Pomarańczowy — zapas 1–2 powt."></span>
      <span class="cdot r ${color === 'red' ? 'active' : ''}" data-c="red" title="Czerwony — seria nieudana"></span>
      <span class="cdot x ${!color ? 'active' : ''}" data-c="" title="Bez koloru"></span>
    </div>
  </div>`;
}
function bindColorDots(root) {
  $$('.cdots', root).forEach(group => {
    const input = group.parentElement.querySelector('input[data-ref]');
    $$('.cdot', group).forEach(dot => {
      dot.addEventListener('click', () => {
        $$('.cdot', group).forEach(d => d.classList.remove('active'));
        dot.classList.add('active');
        const c = dot.dataset.c;
        input.dataset.color = c;
        input.classList.remove('tint-green', 'tint-orange', 'tint-red');
        if (tintClass(c)) input.classList.add(tintClass(c));
      });
    });
  });
}

function commitEntry(root) {
  const sheet = State.model.training.sheetName;
  const styleMap = State.model.training.effortStyles || {};
  let count = 0;
  $$('input[data-ref]', root).forEach(inp => {
    const ref = inp.dataset.ref;
    const val = inp.value.trim();
    if (val === '') return;
    // liczba czy tekst?
    const numVal = val.replace(',', '.');
    const asNum = /^-?\d+(\.\d+)?$/.test(numVal) ? parseFloat(numVal) : null;
    // styl koloru tylko dla serii (pola z atrybutem data-color)
    let style; // undefined = dziedzicz format kolumny
    if (inp.hasAttribute('data-color')) {
      const c = inp.dataset.color;
      if (c && styleMap[c] != null) style = styleMap[c];
      else if (styleMap.neutral != null) style = styleMap.neutral; // wybrane „bez koloru”
    }
    XLSXEngine.setCell(State.wb, sheet, ref, asNum != null ? asNum : val, style);
    count++;
  });
  if (count === 0) { toast('Nie wpisano żadnych wartości.', 'info'); return; }
  saveFile().then(() => { /* model odświeżony w saveFile */ });
}

// =====================================================================
//  WIDOK: RPE
// =====================================================================
function renderRPE(root) {
  const rows = [
    ['10', 'Nie dało się wykonać więcej powtórzeń ani zwiększyć ciężaru', '94–100%', 'red'],
    ['9.5', 'Nie dało się więcej powtórzeń, można minimalnie ↑ ciężar', '89–94%', 'red'],
    ['9', 'Można było wykonać 1 dodatkowe powtórzenie', '82–89%', 'orange'],
    ['8.5', 'Na pewno 1 dodatkowe, szansa na 2', '75–82%', 'orange'],
    ['8', 'Można było wykonać 2 dodatkowe powtórzenia', '75–82%', 'mint'],
    ['7.5', 'Na pewno 2 dodatkowe, szansa na 3', '75–82%', 'mint'],
    ['7', 'Można było wykonać 3 dodatkowe powtórzenia', '65–75%', 'green'],
    ['5–6', '4 do 6 dodatkowych powtórzeń', '65–75%', 'green'],
    ['1–4', 'Bardzo niski poziom wysiłku', '50–60%', 'gray']
  ];
  const card = el('div', { class: 'card pad-lg' });
  card.innerHTML = `
    <h3>Skala RPE (Reps in Reserve)</h3>
    <div class="card-sub">RPE to wskazówka. Większość serii rób blisko upadku, ale ZAWSZE zostaw zapas ~1 powtórzenia. Strefa najefektywniejsza dla Ciebie: <b>RPE 7–9</b>.</div>
    <div class="tbl-scroll" style="max-height:none">
    <table class="tbl"><thead><tr><th class="l">RPE</th><th class="l">Co znaczy</th><th>% ciężaru max</th></tr></thead><tbody>
    ${rows.map(r => `<tr><td class="l"><span class="pill ${r[3]}">${r[0]}</span></td><td class="l">${esc(r[1])}</td><td>${esc(r[2])}</td></tr>`).join('')}
    </tbody></table></div>
    <div class="section-title">Legenda kolorów w raporcie</div>
    <div class="legend-row" style="font-size:13px">
      <span class="lg"><span class="sw" style="background:var(--acc)"></span>Zielony — zapas 2–4 powt. (lekka seria)</span>
      <span class="lg"><span class="sw" style="background:var(--orange)"></span>Pomarańczowy — zapas 1–2 powt. (czujesz opór)</span>
      <span class="lg"><span class="sw" style="background:var(--red)"></span>Czerwony — seria nieudana</span>
    </div>`;
  root.appendChild(card);
}

// =====================================================================
//  INIT
// =====================================================================
function init() {
  setupCharts();
  $$('.nav-item').forEach(n => n.addEventListener('click', () => { State.view = n.dataset.view; render(); }));
  $('#btnOpen').addEventListener('click', openViaPicker);
  $('#btnReload').addEventListener('click', () => $('#fileInput').click());
  $('#btnExport').addEventListener('click', exportFile);
  $('#fileInput').addEventListener('change', e => { const f = e.target.files[0]; if (f) loadFile(f, false); e.target.value = ''; });

  // drop na całe okno
  ['dragover', 'drop'].forEach(ev => window.addEventListener(ev, e => e.preventDefault()));
  window.addEventListener('drop', e => { const f = e.dataTransfer && e.dataTransfer.files[0]; if (f && /\.xlsx$/i.test(f.name)) loadFile(f, false); });

  render();
  tryRestore();
}
document.addEventListener('DOMContentLoaded', init);
