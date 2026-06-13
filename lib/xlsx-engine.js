/* =====================================================================
   xlsx-engine.js  —  lekki silnik do czytania i zapisu .xlsx
   ---------------------------------------------------------------------
   Działa identycznie w przeglądarce i w Node (parsing oparty na regex,
   bez zależności od DOMParser). Czyta KAŻDĄ komórkę po adresie (np. "DI5"),
   niezależnie od tego, czy kolumna jest ukryta w Excelu — bo ukrycie to
   tylko atrybut wyglądu, a dane w XML są zawsze obecne.

   Eksportuje globalny obiekt  window.XLSXEngine  (przeglądarka)
   lub module.exports                              (Node).
   Wymaga wcześniej załadowanego JSZip.
   ===================================================================== */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) {
    module.exports = factory(require('./jszip.min.js'));
  } else {
    root.XLSXEngine = factory(root.JSZip);
  }
})(typeof self !== 'undefined' ? self : this, function (JSZip) {
  'use strict';

  // ---- adresy kolumn: "A"->1, "Z"->26, "AA"->27 ... i odwrotnie -------
  function colToNum(letters) {
    let n = 0;
    for (let i = 0; i < letters.length; i++) {
      n = n * 26 + (letters.charCodeAt(i) - 64);
    }
    return n;
  }
  function numToCol(n) {
    let s = '';
    while (n > 0) {
      const m = (n - 1) % 26;
      s = String.fromCharCode(65 + m) + s;
      n = Math.floor((n - 1) / 26);
    }
    return s;
  }
  function splitRef(ref) {
    const m = /^([A-Z]+)(\d+)$/.exec(ref);
    return { col: colToNum(m[1]), row: parseInt(m[2], 10), colLetter: m[1] };
  }

  // ---- dekodowanie encji XML ------------------------------------------
  function xmlDecode(s) {
    if (s == null) return s;
    return s
      .replace(/&lt;/g, '<')
      .replace(/&gt;/g, '>')
      .replace(/&quot;/g, '"')
      .replace(/&apos;/g, "'")
      .replace(/&#x([0-9a-fA-F]+);/g, (_, h) => String.fromCodePoint(parseInt(h, 16)))
      .replace(/&#(\d+);/g, (_, d) => String.fromCodePoint(parseInt(d, 10)))
      .replace(/&amp;/g, '&');
  }
  function xmlEncode(s) {
    return String(s)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  // ---- sharedStrings ---------------------------------------------------
  function parseSharedStrings(xml) {
    const out = [];
    if (!xml) return out;
    // każdy <si> ... </si>  — może mieć wiele <t> (rich text)
    const siRe = /<si\b[^>]*>([\s\S]*?)<\/si>/g;
    let m;
    while ((m = siRe.exec(xml)) !== null) {
      const inner = m[1];
      let text = '';
      const tRe = /<t\b[^>]*>([\s\S]*?)<\/t>/g;
      let tm;
      let any = false;
      while ((tm = tRe.exec(inner)) !== null) {
        any = true;
        text += xmlDecode(tm[1]);
      }
      if (!any) {
        // <t/> pusty albo brak — sprawdź self-closing
        text = '';
      }
      out.push(text);
    }
    return out;
  }

  // ---- mapowanie nazw arkuszy -> pliki --------------------------------
  function parseWorkbookSheets(workbookXml, relsXml) {
    // workbook.xml: <sheet name="..." sheetId=".." r:id="rIdN"/>
    const rels = {}; // rId -> target
    if (relsXml) {
      const rRe = /<Relationship\b[^>]*?Id="([^"]+)"[^>]*?Target="([^"]+)"[^>]*?\/?>/g;
      let rm;
      while ((rm = rRe.exec(relsXml)) !== null) {
        rels[rm[1]] = rm[2];
      }
    }
    const sheets = [];
    const sRe = /<sheet\b[^>]*?\/?>/g;
    let sm;
    while ((sm = sRe.exec(workbookXml)) !== null) {
      const tag = sm[0];
      const name = (/name="([^"]*)"/.exec(tag) || [])[1];
      const rid = (/r:id="([^"]*)"/.exec(tag) || [])[1];
      let target = rels[rid] || '';
      target = target.replace(/^\/?xl\//, '').replace(/^\//, '');
      if (!target.startsWith('worksheets/')) {
        // czasem target to po prostu worksheets/sheetN.xml
        if (!/worksheets\//.test(target)) target = 'worksheets/' + target.replace(/^.*\//, '');
      }
      sheets.push({ name: xmlDecode(name), path: 'xl/' + target });
    }
    return sheets;
  }

  // ---- parsowanie styles.xml: indeks stylu (xf) -> kolor wypełnienia --
  function parseStyles(xml) {
    if (!xml) return { xfColor: [] };
    const fills = [];
    const fb = /<fills[^>]*>([\s\S]*?)<\/fills>/.exec(xml);
    if (fb) {
      const fillRe = /<fill>([\s\S]*?)<\/fill>/g; let fm;
      while ((fm = fillRe.exec(fb[1])) !== null) {
        const fg = /<fgColor[^>]*rgb="([0-9A-Fa-f]{6,8})"/.exec(fm[1]);
        fills.push(fg ? fg[1].slice(-6).toUpperCase() : null);
      }
    }
    const xfColor = [];
    const xb = /<cellXfs[^>]*>([\s\S]*?)<\/cellXfs>/.exec(xml);
    if (xb) {
      const xfRe = /<xf\b[^>]*?(?:\/>|>[\s\S]*?<\/xf>)/g; let xm;
      while ((xm = xfRe.exec(xb[1])) !== null) {
        const fid = /fillId="(\d+)"/.exec(xm[0]);
        xfColor.push(fid ? (fills[+fid[1]] || null) : null);
      }
    }
    return { xfColor };
  }
  // kategoria koloru z rgb
  function colorCategory(rgb) {
    if (!rgb) return null;
    switch (rgb) {
      case '92D050': case '00FF00': case '00B050': return 'green';
      case 'FFC000': case 'FF9900': return 'orange';
      case 'FF0000': return 'red';
      default: return null;
    }
  }

  // ---- parsowanie jednego arkusza w mapę  ref -> value ----------------
  function parseSheet(xml, shared) {
    const cells = new Map(); // "A1" -> wartość js
    const styles = new Map(); // "A1" -> indeks stylu (s)
    let maxRow = 0, maxCol = 0;
    const hiddenCols = [];
    // ukryte kolumny (tylko informacyjnie) <col min=".." max=".." hidden="1"/>
    const colRe = /<col\b[^>]*?\/?>/g;
    let cmt;
    while ((cmt = colRe.exec(xml)) !== null) {
      if (/hidden="1"|hidden="true"/.test(cmt[0])) {
        const mn = parseInt((/min="(\d+)"/.exec(cmt[0]) || [])[1], 10);
        const mx = parseInt((/max="(\d+)"/.exec(cmt[0]) || [])[1], 10);
        for (let c = mn; c <= mx; c++) hiddenCols.push(numToCol(c));
      }
    }
    // komórki: <c r="A1" s="3" t="s"><v>5</v></c>  lub  <c .../>
    const cRe = /<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g;
    let m;
    while ((m = cRe.exec(xml)) !== null) {
      const attrs = m[1];
      const body = m[2];
      const ref = (/r="([A-Z]+\d+)"/.exec(attrs) || [])[1];
      if (!ref) continue;
      const sAttr = (/\bs="(\d+)"/.exec(attrs) || [])[1];
      if (sAttr != null) styles.set(ref, parseInt(sAttr, 10));
      const t = (/t="([^"]+)"/.exec(attrs) || [])[1] || 'n';
      let value = null;
      if (body != null) {
        if (t === 'inlineStr') {
          let text = '';
          const tRe = /<t\b[^>]*>([\s\S]*?)<\/t>/g;
          let tm;
          while ((tm = tRe.exec(body)) !== null) text += xmlDecode(tm[1]);
          value = text;
        } else {
          const vm = /<v\b[^>]*>([\s\S]*?)<\/v>/.exec(body);
          if (vm) {
            const raw = vm[1];
            if (t === 's') {
              value = shared[parseInt(raw, 10)] != null ? shared[parseInt(raw, 10)] : '';
            } else if (t === 'b') {
              value = raw === '1';
            } else if (t === 'str' || t === 'e') {
              value = xmlDecode(raw);
            } else {
              // numeryczny
              const num = parseFloat(raw);
              value = isNaN(num) ? xmlDecode(raw) : num;
            }
          }
        }
      }
      if (value !== null && value !== '') {
        const p = splitRef(ref);
        if (p.row > maxRow) maxRow = p.row;
        if (p.col > maxCol) maxCol = p.col;
        cells.set(ref, value);
      }
    }
    return { cells, styles, maxRow, maxCol, hiddenCols };
  }

  // ---- API: wczytanie całego workbooka --------------------------------
  async function load(arrayBufferOrUint8) {
    const zip = await JSZip.loadAsync(arrayBufferOrUint8);
    const sharedXml = zip.file('xl/sharedStrings.xml')
      ? await zip.file('xl/sharedStrings.xml').async('string') : '';
    const shared = parseSharedStrings(sharedXml);
    const workbookXml = await zip.file('xl/workbook.xml').async('string');
    const relsFile = zip.file('xl/_rels/workbook.xml.rels');
    const relsXml = relsFile ? await relsFile.async('string') : '';
    const sheetDefs = parseWorkbookSheets(workbookXml, relsXml);
    const stylesXml = zip.file('xl/styles.xml') ? await zip.file('xl/styles.xml').async('string') : '';
    const styleInfo = parseStyles(stylesXml);

    const sheets = {};
    const order = [];
    for (const def of sheetDefs) {
      const f = zip.file(def.path);
      if (!f) continue;
      const xml = await f.async('string');
      sheets[def.name] = Object.assign({ path: def.path, xml: xml }, parseSheet(xml, shared));
      order.push(def.name);
    }
    const wb = { zip, shared, sheets, order, pending: {}, styleInfo };
    wb.effortStyles = detectEffortStyles(wb, order[0]);
    return wb;
  }

  // kategoria koloru wypełnienia komórki (green/orange/red/null)
  function effortCategory(wb, sheetName, ref) {
    const sh = wb.sheets[sheetName];
    if (!sh || !sh.styles) return null;
    const s = sh.styles.get(ref);
    if (s == null) return null;
    return colorCategory(wb.styleInfo.xfColor[s]);
  }

  // wykryj reprezentatywne indeksy stylów dla każdego koloru — najlepiej
  // ze stylów już użytych na komórkach serii (kolumny >= 17), żeby font
  // i obramowanie pasowały do reszty tabeli. Fallback: legenda / dowolny.
  function detectEffortStyles(wb, sheetName) {
    const sh = wb.sheets[sheetName];
    const xfColor = wb.styleInfo.xfColor;
    const res = { green: null, orange: null, red: null, neutral: null };
    if (!sh || !sh.styles) return res;
    const tally = { green: {}, orange: {}, red: {} };
    const neutralTally = {};
    sh.styles.forEach((s, ref) => {
      const p = splitRef(ref);
      if (p.col < 17) return;            // tylko obszar serii
      const cat = colorCategory(xfColor[s]);
      if (cat) tally[cat][s] = (tally[cat][s] || 0) + 1;
      else if (xfColor[s] == null) neutralTally[s] = (neutralTally[s] || 0) + 1;
    });
    const top = (obj) => {
      let best = null, n = -1;
      for (const k in obj) if (obj[k] > n) { n = obj[k]; best = parseInt(k, 10); }
      return best;
    };
    res.green = top(tally.green);
    res.orange = top(tally.orange);
    res.red = top(tally.red);
    res.neutral = top(neutralTally);
    // fallback: znajdź DOWOLNY xf danego koloru
    const anyOf = (cat) => { for (let i = 0; i < xfColor.length; i++) if (colorCategory(xfColor[i]) === cat) return i; return null; };
    if (res.green == null) res.green = anyOf('green');
    if (res.orange == null) res.orange = anyOf('orange');
    if (res.red == null) res.red = anyOf('red');
    return res;
  }

  // pomocnik: pobierz wartość komórki danego arkusza
  function get(wb, sheetName, ref) {
    const sh = wb.sheets[sheetName];
    if (!sh) return null;
    const v = sh.cells.get(ref);
    return v == null ? null : v;
  }

  // ===================================================================
  //  ZAPIS — chirurgiczna edycja XML (zachowuje formatowanie i style)
  // ===================================================================

  // zaplanuj wpis komórki: value = liczba | string | null(=wyczyść)
  // style = indeks stylu (kolor) | undefined(=dziedzicz) | null(=bez zmian koloru)
  function setCell(wb, sheetName, ref, value, style) {
    if (!wb.pending) wb.pending = {};
    if (!wb.pending[sheetName]) wb.pending[sheetName] = {};
    wb.pending[sheetName][ref] = { value: value, style: style };
    // zaktualizuj też model w pamięci
    const sh = wb.sheets[sheetName];
    if (sh) {
      if (value == null || value === '') sh.cells.delete(ref);
      else sh.cells.set(ref, value);
      if (style != null && sh.styles) sh.styles.set(ref, style);
    }
  }

  function hasPending(wb) {
    return wb.pending && Object.keys(wb.pending).some((s) => Object.keys(wb.pending[s]).length);
  }

  // zbuduj XML komórki, dziedzicząc styl (atrybut s="..") jeśli podany
  function buildCellXml(ref, value, styleAttr) {
    const s = styleAttr ? ' s="' + styleAttr + '"' : '';
    if (value == null || value === '') {
      return '<c r="' + ref + '"' + s + '/>';
    }
    if (typeof value === 'number' && isFinite(value)) {
      return '<c r="' + ref + '"' + s + '><v>' + value + '</v></c>';
    }
    // tekst -> inline string (nie ruszamy sharedStrings)
    return '<c r="' + ref + '"' + s + ' t="inlineStr"><is><t xml:space="preserve">' +
      xmlEncode(value) + '</t></is></c>';
  }

  // wyłuskaj styl s="N" z istniejącego znacznika <c ...>
  function styleOf(cTag) {
    const m = /\ss="(\d+)"/.exec(cTag);
    return m ? m[1] : null;
  }

  // zastosuj wpisy do surowego XML jednego arkusza (zwraca nowy XML)
  function applyToSheetXml(xml, writes) {
    // writes: { ref: value }
    // rozbij na wiersze sheetData
    const sdMatch = /(<sheetData[^>]*>)([\s\S]*?)(<\/sheetData>)/.exec(xml);
    if (!sdMatch) return xml;
    const head = xml.slice(0, sdMatch.index) + sdMatch[1];
    const tail = sdMatch[3] + xml.slice(sdMatch.index + sdMatch[0].length);
    let body = sdMatch[2];

    // mapa wierszy: rowNum -> {attrs, cells:[{col,ref,xml}], raw}
    const rows = [];
    const rowRe = /<row\b([^>]*)>([\s\S]*?)<\/row>|<row\b([^>]*)\/>/g;
    let rm;
    while ((rm = rowRe.exec(body)) !== null) {
      const attrs = (rm[1] != null ? rm[1] : rm[3]) || '';
      const inner = rm[2] || '';
      const rn = parseInt((/\br="(\d+)"/.exec(attrs) || [])[1], 10);
      const cells = [];
      const cRe = /<c\b[^>]*?(?:\/>|>[\s\S]*?<\/c>)/g;
      let cm;
      while ((cm = cRe.exec(inner)) !== null) {
        const ctag = cm[0];
        const cref = (/\br="([A-Z]+\d+)"/.exec(ctag) || [])[1];
        cells.push({ col: cref ? colToNum(splitRef(cref).colLetter) : 0, ref: cref, xml: ctag });
      }
      rows.push({ rn, attrs, cells });
    }

    // pogrupuj wpisy po wierszu  (writes[ref] = { value, style })
    const byRow = {};
    for (const ref in writes) {
      const p = splitRef(ref);
      const w = writes[ref];
      (byRow[p.row] = byRow[p.row] || []).push({ ref, col: p.col, value: w.value, style: w.style });
    }

    // styl wzorcowy dla kolumny: poszukaj istniejącej komórki w tej kolumnie
    function columnStyle(col) {
      for (const row of rows) {
        for (const c of row.cells) {
          if (c.col === col && c.ref) {
            const st = styleOf(c.xml);
            if (st) return st;
          }
        }
      }
      return null;
    }

    for (const rnStr in byRow) {
      const rn = parseInt(rnStr, 10);
      let row = rows.find((r) => r.rn === rn);
      if (!row) {
        row = { rn, attrs: ' r="' + rn + '"', cells: [] };
        // wstaw w kolejności numerycznej
        let idx = rows.findIndex((r) => r.rn > rn);
        if (idx < 0) idx = rows.length;
        rows.splice(idx, 0, row);
      }
      for (const w of byRow[rnStr]) {
        const existing = row.cells.find((c) => c.ref === w.ref);
        // wybór stylu: jawny (kolor) > istniejący > wzorzec kolumny
        let style;
        if (w.style != null) style = String(w.style);
        else if (existing) style = styleOf(existing.xml);
        else style = columnStyle(w.col);
        const newXml = buildCellXml(w.ref, w.value, style);
        if (existing) {
          existing.xml = newXml;
        } else {
          let idx = row.cells.findIndex((c) => c.col > w.col);
          if (idx < 0) idx = row.cells.length;
          row.cells.splice(idx, 0, { col: w.col, ref: w.ref, xml: newXml });
        }
      }
    }

    // serializacja
    let out = '';
    for (const row of rows) {
      if (!row.cells.length) {
        out += '<row' + row.attrs + '/>';
      } else {
        out += '<row' + row.attrs + '>' + row.cells.map((c) => c.xml).join('') + '</row>';
      }
    }
    return head + out + tail;
  }

  // wygeneruj Blob/Uint8 zmodyfikowanego pliku (zachowuje resztę zipa)
  async function toBlob(wb, type) {
    const zip = wb.zip;
    for (const sheetName in (wb.pending || {})) {
      const writes = wb.pending[sheetName];
      if (!Object.keys(writes).length) continue;
      const sh = wb.sheets[sheetName];
      const newXml = applyToSheetXml(sh.xml, writes);
      sh.xml = newXml; // utrwal
      zip.file(sh.path, newXml);
    }
    wb.pending = {};
    const opts = { type: type || 'uint8array', compression: 'DEFLATE' };
    return zip.generateAsync(opts);
  }

  return {
    colToNum, numToCol, splitRef, xmlEncode, xmlDecode,
    parseSharedStrings, parseWorkbookSheets, parseSheet, parseStyles,
    colorCategory, effortCategory, detectEffortStyles,
    load, get, setCell, hasPending, toBlob, applyToSheetXml
  };
});
