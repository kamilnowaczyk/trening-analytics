/* =====================================================================
   training-model.js — budowa modelu treningowego z parsowanego xlsx
   ---------------------------------------------------------------------
   Wejście: obiekt workbook z xlsx-engine.load()
   Wyjście: ustrukturyzowany model { meta, weeks, days, exercises,
            metrics, questions, rpe }.

   Tygodnie w arkuszu treningowym są wykrywane PO POZYCJI kolumn,
   a nie po nagłówkach (które bywają błędne). Tydzień N:
       kolumna ciężaru   = 17 + (N-1)*5      (Q dla tygodnia 1)
       serie (reps)      = kolejne 4 kolumny
   Dzięki temu ukryte kolumny nie mają znaczenia — czytamy po adresie.
   ===================================================================== */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) {
    module.exports = factory(require('./xlsx-engine.js'));
  } else {
    root.TrainingModel = factory(root.XLSXEngine);
  }
})(typeof self !== 'undefined' ? self : this, function (XE) {
  'use strict';

  const numToCol = XE.numToCol;
  const WEIGHT_COL0 = 17;      // kolumna Q = pierwszy tydzień
  const WEEK_STRIDE = 5;       // co 5 kolumn kolejny tydzień
  const SETS_PER_WEEK = 4;     // do 4 serii na tydzień

  const EMPTY = new Set(['', 'x', 'X', '-', '–', '—', null, undefined]);
  const isEmpty = (v) => v == null || EMPTY.has(String(v).trim());

  // ---- parsowanie ciężaru: "82,5kg", "2x40kg", "27 nr10", "nr11 54" ---
  function parseWeight(raw) {
    if (isEmpty(raw)) return null;
    const s = String(raw).trim().replace(/,/g, '.');
    const out = { raw: String(raw).trim(), num: null, perHand: false, pin: null };
    const pin = /nr\s*\.?\s*(\d+)/i.exec(s);
    if (pin) out.pin = parseInt(pin[1], 10);
    // "2x40", "2 x 37.5", "2X12,5"
    const dbl = /(\d+)\s*[xX]\s*([\d.]+)/.exec(s);
    if (dbl) {
      out.perHand = true;
      out.num = parseFloat(dbl[2]);
      return out;
    }
    // pierwsza liczba która NIE jest numerem sztabki
    let m, re = /([\d]+(?:\.\d+)?)/g;
    while ((m = re.exec(s)) !== null) {
      const idx = m.index;
      const before = s.slice(Math.max(0, idx - 4), idx).toLowerCase();
      if (/nr\s*\.?\s*$/.test(before)) continue; // to numer sztabki, pomiń
      out.num = parseFloat(m[1]);
      break;
    }
    if (out.num == null && out.pin != null) out.num = null; // tylko maszyna
    return out;
  }

  // efektywny ciężar (kg realnie obciążające) — hantle liczone x2
  function effectiveLoad(w) {
    if (!w || w.num == null) return null;
    return w.perHand ? w.num * 2 : w.num;
  }

  // ---- parsowanie powtórzeń ------------------------------------------
  // Obsługuje: 8 · "9." · myoreps "11+7" · top-sety (odchył od reguły, gdzie
  // jedna seria jest cięższa, a reszta na bazowym obciążeniu):
  //   "110(6)" "110/6" "115kg(7)" "130kg/6" "6(100kg)" "2x12,5/12"
  //   "nr13 28,4kg/14" "nr11/10"
  // Zwraca { reps, raw, extra(myorep), w(efektywny ciężar tej serii|null), pin, perHand, top }
  function parseReps(raw) {
    if (isEmpty(raw)) return null;
    if (typeof raw === 'number') return { reps: raw, raw: raw, extra: null, w: null, pin: null, perHand: false, top: false };
    const orig = String(raw).trim();
    const out = { reps: null, raw: orig, extra: null, w: null, pin: null, perHand: false, top: false };
    if (/^max$/i.test(orig)) { out.extra = 'max'; return out; }
    let t = orig.replace(/,/g, '.');

    // numer sztabki (maszyna/wyciąg)
    const pin = /nr\s*\.?\s*(\d+(?:\.\d+)?)/i.exec(t);
    if (pin) { out.pin = parseFloat(pin[1]); t = t.replace(pin[0], ' '); }

    // top-set hantlami "2x12.5/12"
    const db = /(\d+)\s*[x]\s*(\d+(?:\.\d+)?)/i.exec(t);
    if (db) { out.perHand = true; out.w = parseFloat(db[2]) * 2; out.top = true; t = t.replace(db[0], ' '); }

    // top-set z kg "130kg" / "28.4kg"
    if (out.w == null) {
      const kg = /(\d+(?:\.\d+)?)\s*kg/i.exec(t);
      if (kg) { out.w = parseFloat(kg[1]); out.top = true; t = t.replace(kg[0], ' '); }
    }

    // myoreps "13+10"
    const myo = /(\d+)\s*\+\s*(\d+)/.exec(t);
    if (myo) { out.reps = parseInt(myo[1], 10); out.extra = parseInt(myo[2], 10); return out; }

    // powtórzenia w nawiasie "(7)" lub po ukośniku "/6"
    const rm = /[\/(]\s*(\d+(?:\.\d+)?)\s*\)?/.exec(t);
    if (rm) {
      out.reps = parseFloat(rm[1]);
      // wiodąca liczba przed / lub ( to ciężar top-setu, jeśli duża i nie złapana wcześniej
      const lead = /^\s*(\d+(?:\.\d+)?)\s*[\/(]/.exec(t);
      if (out.w == null && lead && parseFloat(lead[1]) > 30) { out.w = parseFloat(lead[1]); out.top = true; }
      return out;
    }

    // pozostała pierwsza liczba = powtórzenia
    const n = /(\d+(?:\.\d+)?)/.exec(t);
    if (n) out.reps = parseFloat(n[1]);
    return out;
  }

  // tolerancyjny parser liczby (pomiary/waga): "97 ,4" -> 97.4
  function num(raw) {
    if (isEmpty(raw)) return null;
    if (typeof raw === 'number') return raw;
    const s = String(raw).replace(/\s/g, '').replace(',', '.');
    const m = /(-?\d+(?:\.\d+)?)/.exec(s);
    return m ? parseFloat(m[1]) : null;
  }

  // ---- znajdź arkusz po fragmencie nazwy ------------------------------
  function findSheet(wb, ...frags) {
    for (const name of wb.order) {
      const low = name.toLowerCase();
      if (frags.some((f) => low.includes(f.toLowerCase()))) return name;
    }
    return null;
  }

  // ===================================================================
  //  ARKUSZ TRENINGOWY
  // ===================================================================
  function buildTraining(wb) {
    const sheetName = wb.order[0]; // pierwszy arkusz = plan/log
    const sh = wb.sheets[sheetName];
    const get = (col, row) => sh.cells.get(numToCol(col) + row);
    const effortColorAt = (ref) => (XE.effortCategory ? XE.effortCategory(wb, sheetName, ref) : null);
    // litery -> numery kolumn planu
    const C = { A:1,B:2,C:3,D:4,E:5,F:6,G:7,H:8,I:9,J:10,K:11,L:12,M:13,N:14,O:15,P:16 };

    // maksymalna liczba bloków-tygodni jaka mieści się w arkuszu
    const maxPossible = Math.ceil((sh.maxCol - WEIGHT_COL0 + 1) / WEEK_STRIDE) + 2;
    // weekCount policzymy PO zebraniu ćwiczeń — tylko z realnych danych,
    // żeby nie łapać nagłówków "Ciężar" powtarzanych w każdym bloku.
    let weekCount = 0;

    const exercises = [];
    const days = [];
    let curDay = null;
    let dayIdx = 0;

    for (let r = 1; r <= sh.maxRow; r++) {
      const a = get(C.A, r);
      const c = get(C.C, r);
      const d = get(C.D, r);

      // wiersz nagłówka dnia: A == "Kolejność"
      if (typeof a === 'string' && a.trim().toLowerCase().startsWith('kolejno')) {
        dayIdx++;
        curDay = { idx: dayIdx, label: c ? String(c).trim() : ('Dzień ' + dayIdx), headerRow: r, exercises: [] };
        days.push(curDay);
        continue;
      }
      // wiersz "PRZERWA ..." — pomijamy
      if (typeof c === 'string' && /przerwa/i.test(c)) continue;

      // wiersz ćwiczenia: A jest liczbą (kolejność) i jest nazwa ćwiczenia
      const orderNum = typeof a === 'number' ? a : (typeof a === 'string' && /^\d+$/.test(a.trim()) ? parseInt(a, 10) : null);
      if (orderNum == null || isEmpty(d)) continue;

      const ex = {
        row: r,
        order: orderNum,
        dayIdx: curDay ? curDay.idx : 0,
        dayLabel: curDay ? curDay.label : '',
        superset: get(C.B, r) || '',
        muscles: c ? String(c).trim() : '',
        name: String(d).trim(),
        tempo: get(C.F, r) || '',
        rest: get(C.G, r) || '',
        setsPlanned: get(C.H, r),
        rir: get(C.I, r),
        repsTarget: get(C.J, r) || '',
        warmup: get(C.K, r) || '',
        planWeight: get(C.L, r),
        prescription: [get(C.M, r), get(C.N, r), get(C.O, r), get(C.P, r)]
          .map((v) => (isEmpty(v) ? null : v)),
        weeks: [] // per-tydzień
      };

      for (let w = 1; w <= maxPossible; w++) {
        const wc = WEIGHT_COL0 + (w - 1) * WEEK_STRIDE;
        const weight = parseWeight(get(wc, r));
        const eff = effectiveLoad(weight);
        const sets = [];
        for (let s = 1; s <= SETS_PER_WEEK; s++) {
          const ref = numToCol(wc + s) + r;
          const rp = parseReps(get(wc + s, r));
          if (rp) {
            // efektywny ciężar tej serii: top-set ma własny, reszta = bazowy
            rp.setEff = rp.w != null ? rp.w : eff;
            rp.color = effortColorAt ? effortColorAt(ref) : null;
            rp.ref = ref;
            sets.push(rp);
          }
        }
        const repsArr = sets.map((s) => s.reps).filter((x) => x != null);
        const topReps = repsArr.length ? Math.max(...repsArr) : null;
        const totalReps = repsArr.reduce((a, b) => a + b, 0);
        // tonaż i 1RM liczone PER SERIA (uwzględnia cięższe top-sety)
        let tonnage = null, e1rm = null;
        sets.forEach((s) => {
          if (s.reps == null || s.setEff == null) return;
          tonnage = (tonnage || 0) + s.setEff * s.reps;
          const est = s.setEff * (1 + s.reps / 30);
          if (s.reps > 0 && (e1rm == null || est > e1rm)) e1rm = est;
        });
        // typ obciążenia (do wykrycia zmiany ćwiczenia)
        const kind = weight ? (weight.perHand ? 'db' : (weight.pin != null ? 'machine' : 'bb')) : null;
        ex.weeks.push({
          week: w, weightCell: numToCol(wc) + r, weight, eff, kind,
          sets, repsArr, topReps, totalReps, numSets: sets.length,
          hasTopSet: sets.some((s) => s.top),
          tonnage, e1rm,
          hasData: !!(weight || sets.length)
        });
      }
      exercises.push(ex);
      if (curDay) curDay.exercises.push(ex);
    }

    // weekCount = ostatni blok, w którym JAKIEKOLWIEK ćwiczenie ma realne dane
    for (const ex of exercises) {
      ex.weeks.forEach((wk) => {
        if (wk.hasData && wk.week > weekCount) weekCount = wk.week;
      });
    }
    if (weekCount === 0) weekCount = 1;
    // przytnij tygodnie do realnej liczby
    for (const ex of exercises) ex.weeks = ex.weeks.slice(0, weekCount);

    // wykryj możliwą zmianę ćwiczenia: zmiana typu obciążenia (bb/db/machine)
    // w trakcie serii danych — może oznaczać podmianę ćwiczenia w tym slocie.
    for (const ex of exercises) {
      ex.swapHints = [];
      let prevKind = null;
      ex.weeks.forEach((wk) => {
        if (!wk.kind) return;
        if (prevKind && wk.kind !== prevKind) ex.swapHints.push({ week: wk.week, from: prevKind, to: wk.kind });
        prevKind = wk.kind;
      });
    }

    return { sheetName, weekCount, exercises, days, effortStyles: wb.effortStyles || null };
  }

  // ===================================================================
  //  ARKUSZ POMIARÓW / WAGI / KROKÓW
  // ===================================================================
  function buildMetrics(wb) {
    const name = findSheet(wb, 'kroki', 'pomiar', 'waga');
    if (!name) return null;
    const sh = wb.sheets[name];
    const cell = (col, row) => sh.cells.get(numToCol(col) + row);
    // kolumna tygodnia i: B=2 -> tydzień1
    const wcol = (i) => 2 + (i - 1);

    // ile tygodni w nagłówku (wiersz 2: B2='Tydzień 1'...)
    let weeks = 0;
    for (let i = 1; i <= 60; i++) {
      const h = cell(wcol(i), 2);
      if (h && /tydzie/i.test(String(h))) weeks = i; else if (i > 1) break;
    }

    function rowSeries(row, parser) {
      const arr = [];
      for (let i = 1; i <= weeks; i++) {
        const v = cell(wcol(i), row);
        arr.push(isEmpty(v) ? null : (parser ? parser(v) : v));
      }
      return arr;
    }
    function findRow(colA_text, minRow) {
      for (let r = (minRow || 1); r <= sh.maxRow; r++) {
        const v = sh.cells.get('A' + r);
        if (v && String(v).toLowerCase().includes(colA_text.toLowerCase())) return r;
      }
      return null;
    }
    // kotwica: nagłówek tygodniowej tabeli pomiarów (A='Pomiary' + B='Tydzień 1')
    let measAnchor = 1;
    for (let r = 1; r <= sh.maxRow; r++) {
      const a = sh.cells.get('A' + r), b = sh.cells.get('B' + r);
      if (a && /^pomiary$/i.test(String(a).trim()) && b && /tydzie/i.test(String(b))) { measAnchor = r; break; }
    }

    const steps   = rowSeries(findRow('wykonanych krok') || 8, num);
    // obwody i waga-średnia czytamy DOPIERO od kotwicy (pomijamy pomiary wstępne)
    const wagaAvg = rowSeries(findRow('Waga (', measAnchor) || 19, num);
    const udo     = rowSeries(findRow('Udo', measAnchor) || 14, num);
    const pas     = rowSeries(findRow('Pas', measAnchor) || 15, num);
    const biodra  = rowSeries(findRow('Biodra', measAnchor) || 16, num);
    const klatka  = rowSeries(findRow('Klatka', measAnchor) || 17, num);
    const biceps  = rowSeries(findRow('Biceps (', measAnchor) || 18, num);
    const kcalAvg = rowSeries(findRow('rednia kalorii') || 44, num);
    // sen — średnia z 7 dni (wiersze po nagłówku "Sen")
    const senHead = findRow('Sen (') || 28;
    const senRows = [];
    for (let r = senHead + 1; r <= senHead + 7; r++) {
      const lbl = sh.cells.get('A' + r);
      if (lbl && /czwartek|pi.tek|sobota|niedziela|poniedzia|wtorek|.roda/i.test(String(lbl))) senRows.push(r);
    }
    const senAvg = [];
    for (let i = 1; i <= weeks; i++) {
      const vals = senRows.map((r) => num(cell(wcol(i), r))).filter((x) => x != null);
      senAvg.push(vals.length ? vals.reduce((a, b) => a + b, 0) / vals.length : null);
    }
    // białko — średnia z dni (wiersze po "Białko")
    const bHead = findRow('ko (ilo') || findRow('Bia') || 45;
    const bAvg = [];
    if (bHead) {
      const bRows = [];
      for (let r = bHead + 1; r <= bHead + 7; r++) {
        const lbl = sh.cells.get('A' + r);
        if (lbl && /czwartek|pi.tek|sobota|niedziela|poniedzia|wtorek|.roda/i.test(String(lbl))) bRows.push(r);
      }
      for (let i = 1; i <= weeks; i++) {
        const vals = bRows.map((r) => num(cell(wcol(i), r))).filter((x) => x != null);
        bAvg.push(vals.length ? vals.reduce((a, b) => a + b, 0) / vals.length : null);
      }
    }

    // ostatni tydzień z jakimikolwiek danymi (przytnij puste ogony)
    let weeksFilled = 0;
    [steps, wagaAvg, udo, pas, biodra, klatka, biceps, kcalAvg, senAvg, bAvg].forEach((arr) => {
      for (let i = arr.length - 1; i >= 0; i--) {
        if (arr[i] != null) { if (i + 1 > weeksFilled) weeksFilled = i + 1; break; }
      }
    });
    if (weeksFilled === 0) weeksFilled = weeks;

    return { sheetName: name, weeks, weeksFilled, steps, wagaAvg, udo, pas, biodra, klatka, biceps, kcalAvg, senAvg, bialkoAvg: bAvg,
      rowMap: { udo: findRow('Udo'), pas: findRow('Pas'), biodra: findRow('Biodra'),
                klatka: findRow('Klatka'), biceps: findRow('Biceps ('), waga: findRow('Waga ('),
                steps: findRow('wykonanych krok') } };
  }

  // ===================================================================
  //  ARKUSZE PYTAŃ
  // ===================================================================
  function buildQuestions(wb, frag) {
    const name = findSheet(wb, frag);
    if (!name) return null;
    const sh = wb.sheets[name];
    // nagłówki pytań w wierszu 3 (B..)
    const headers = [];
    for (let c = 2; c <= 10; c++) {
      const h = sh.cells.get(numToCol(c) + 3);
      if (h) headers.push({ col: c, text: String(h).trim() });
    }
    const rows = [];
    for (let r = 4; r <= sh.maxRow; r++) {
      const a = sh.cells.get('A' + r);
      if (!a || !/tydzie/i.test(String(a))) continue;
      const wk = parseInt((/(\d+)/.exec(String(a)) || [])[1], 10);
      const answers = headers.map((h) => {
        const v = sh.cells.get(numToCol(h.col) + r);
        return isEmpty(v) ? null : String(v).trim();
      });
      if (answers.some((x) => x != null)) rows.push({ week: wk, answers });
    }
    return { sheetName: name, headers, rows };
  }

  // ===================================================================
  function build(wb) {
    const training = buildTraining(wb);
    const metrics = buildMetrics(wb);
    const qTren = buildQuestions(wb, 'pytania treningowe');
    const qRap = buildQuestions(wb, 'pytania do raportu');
    const meta = {
      sheets: wb.order,
      weekCount: training.weekCount,
      generatedAt: new Date().toISOString()
    };
    return { meta, training, metrics, questions: { treningowe: qTren, raport: qRap } };
  }

  return { build, parseWeight, parseReps, num, effectiveLoad,
    WEIGHT_COL0, WEEK_STRIDE, SETS_PER_WEEK, numToCol, isEmpty };
});
