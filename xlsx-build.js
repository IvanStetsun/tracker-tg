/* Builds the monthly "resource tracker" workbook.
 *
 * data = { y, m, name, role, nw, wr, p: [ { n: "Project", t: [ { n: "Task", w: [min x nw], e: [min x nw] } ] } ] }
 *   nw - number of calendar weeks (Mon-Sun) the month touches, 4..6; week 1 runs from the 1st to the first Sunday
 *   wr - date range label of every week, e.g. ["1-4", "5-11", ...]
 *
 * Week cells hold time as ГОД.ХВ (1.20 = 1 h 20 min), like the original template; every total is a
 * formula that converts to minutes and back, so editing a cell by hand keeps the sums correct.
 */
(function (root) {
  var C = {
    navy: 'FF1F3A6E', blue: 'FF2E5CB8', orange: 'FFC55A11', green: 'FF38761D',
    blueBand: 'FFDDEBF7', blueSum: 'FFC9DDF3', orangeSum: 'FFFCE4D6', greenSum: 'FFE2EFDA', yellow: 'FFFFF2CC',
    border: 'FFBFBFBF', white: 'FFFFFFFF', grayText: 'FF595959',
  };
  var MONTHS = ['січень', 'лютий', 'березень', 'квітень', 'травень', 'червень',
    'липень', 'серпень', 'вересень', 'жовтень', 'листопад', 'грудень'];
  var NUM = '0.00';

  function hm(min) { return Math.floor(min / 60) + (min % 60) / 100; }
  function term(c) { return '(INT(' + c + ')*60+ROUND((' + c + '-INT(' + c + '))*100,0))'; }
  function hmFormula(m) { return 'INT((' + m + ')/60)+MOD(' + m + ',60)/100'; }
  function sumMinutes(arr) { return arr.reduce(function (a, b) { return a + b; }, 0); }

  /** 1 -> A, 27 -> AA */
  function L(n) {
    var s = '';
    while (n > 0) { var r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); }
    return s;
  }

  function style(cell, o) {
    if (o.fill) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: o.fill } };
    cell.font = { name: 'Calibri', size: o.size || 11, bold: !!o.bold, italic: !!o.italic, color: { argb: o.color || 'FF000000' } };
    cell.alignment = { horizontal: o.align || 'center', vertical: 'middle', wrapText: !!o.wrap };
    var line = { style: 'thin', color: { argb: C.border } };
    cell.border = { top: line, left: line, bottom: line, right: line };
    if (o.num) cell.numFmt = o.num;
  }

  function put(ws, ref, value, o) {
    var cell = ws.getCell(ref);
    cell.value = value;
    style(cell, o);
    return cell;
  }

  /** Column numbers for a month with nW week columns per section. */
  function layout(nW) {
    var k = {};
    k.workFirst = 4;
    k.workLast = 3 + nW;
    k.sumW = 4 + nW;
    k.editFirst = 5 + nW;
    k.editLast = 4 + 2 * nW;
    k.sumE = 5 + 2 * nW;
    k.total = 6 + 2 * nW;
    k.pct = 7 + 2 * nW;
    k.avg = 8 + 2 * nW;
    k.count = 9 + 2 * nW;
    k.minW = 10 + 2 * nW;   // hidden helper: work minutes
    k.minE = 11 + 2 * nW;   // hidden helper: edit minutes
    k.last = k.count;       // last visible column
    return k;
  }

  function block(ws, project, r0, data, K, nW) {
    var tasks = project.t;
    var n = tasks.length;
    var first = r0 + 4;
    var last = first + n - 1;
    var totalRow = first + n;
    var monthText = MONTHS[data.m - 1] + ' ' + data.y;
    var ranges = data.wr || [];

    // Title + hint
    ws.mergeCells(r0, 1, r0, K.last);
    put(ws, 'A' + r0, '📁 ПРОЄКТ: ' + project.n.toUpperCase() + ' — ТРЕКЕР РЕСУРСОЗАТРАТНОСТІ',
      { fill: C.navy, color: C.white, bold: true, size: 14 });
    ws.getRow(r0).height = 32;

    ws.mergeCells(r0 + 1, 1, r0 + 1, K.last);
    put(ws, 'A' + (r0 + 1),
      'Дані за ' + monthText + ' заповнені автоматично з трекера. Місячні підсумки рахуються автоматично. ' +
      '⏱ Час у форматі ГОД.ХВ: 0.40 = 40 хв, 1.20 = 1 год 20 хв (хвилини 00–59). ' +
      'Тижні календарні (пн–нд): Тиж. 1 — від 1-го числа до першої неділі, далі з понеділка.',
      { italic: true, color: C.grayText, align: 'left', wrap: true, size: 10 });
    ws.getRow(r0 + 1).height = 30;

    // Group headers
    var g = r0 + 2;
    [
      [1, 3, '👤 Учасник та формат роботи', C.navy],
      [K.workFirst, K.sumW, '⏱ ОСНОВНА РОБОТА (год)', C.blue],
      [K.editFirst, K.sumE, '✏️ ПРАВКИ (год)', C.orange],
      [K.total, K.count, '📊 ПІДСУМКИ МІСЯЦЯ', C.green],
    ].forEach(function (x) {
      for (var c = x[0]; c <= x[1]; c++) style(ws.getCell(g, c), { fill: x[3], color: C.white, bold: true });
      ws.mergeCells(g, x[0], g, x[1]);
      put(ws, L(x[0]) + g, x[2], { fill: x[3], color: C.white, bold: true });
    });
    ws.getRow(g).height = 24;

    // Column headers
    var h = r0 + 3;
    var head = function (col, text, fill) {
      put(ws, L(col) + h, text, { fill: fill, color: C.white, bold: true, size: 10, wrap: true });
    };
    head(1, "Ім'я", C.navy); head(2, 'Роль', C.navy); head(3, 'Формат роботи', C.navy);
    for (var i = 0; i < nW; i++) {
      var label = 'Тиж. ' + (i + 1) + (ranges[i] ? '\n' + ranges[i] : '');
      head(K.workFirst + i, label, C.blue);
      head(K.editFirst + i, label, C.orange);
    }
    head(K.sumW, 'Σ Осн. (міс.)', C.blue);
    head(K.sumE, 'Σ Правок (міс.)', C.orange);
    head(K.total, 'Σ Всього (міс.)', C.green);
    head(K.pct, '% Правок', C.green);
    head(K.avg, 'Сер./тиж. (осн.)', C.green);
    head(K.count, 'К-сть форматів', C.green);
    ws.getRow(h).height = 36;

    // Task rows
    var sumW = 0, sumE = 0, formats = 0;
    var colW = [], colE = [];
    for (i = 0; i < nW; i++) { colW.push(0); colE.push(0); }
    var cMinW = L(K.minW), cMinE = L(K.minE);

    tasks.forEach(function (t, idx) {
      var r = first + idx;
      var band = idx % 2 === 0 ? C.blueBand : C.white;
      var minW = sumMinutes(t.w), minE = sumMinutes(t.e);
      sumW += minW; sumE += minE; formats += minW + minE > 0 ? 1 : 0;
      for (var k = 0; k < nW; k++) { colW[k] += t.w[k] || 0; colE[k] += t.e[k] || 0; }

      put(ws, 'A' + r, idx === 0 ? data.name : '', { fill: band, bold: true, color: C.navy, align: 'left' });
      put(ws, 'B' + r, idx === 0 ? data.role : '', { fill: band });
      put(ws, 'C' + r, t.n, { fill: band, align: 'left' });

      var termsW = [], termsE = [];
      for (k = 0; k < nW; k++) {
        put(ws, L(K.workFirst + k) + r, hm(t.w[k] || 0), { fill: band, num: NUM, color: 'FF1F4E9C' });
        put(ws, L(K.editFirst + k) + r, hm(t.e[k] || 0), { fill: band, num: NUM, color: 'FFB45309' });
        termsW.push(term(L(K.workFirst + k) + r));
        termsE.push(term(L(K.editFirst + k) + r));
      }
      ws.getCell(cMinW + r).value = { formula: termsW.join('+'), result: minW };
      ws.getCell(cMinE + r).value = { formula: termsE.join('+'), result: minE };
      summary(ws, r, K, nW, minW, minE, band, false);
    });

    // Total row
    var T = totalRow;
    var dark = { fill: C.navy, color: C.white, bold: true };
    put(ws, 'A' + T, data.name, Object.assign({ align: 'left' }, dark));
    put(ws, 'B' + T, data.role === 'Лід' ? 'Менеджер' : data.role, dark);
    put(ws, 'C' + T, 'РАЗОМ', dark);

    for (var w = 0; w < nW; w++) {
      [[K.workFirst + w, colW[w], C.blueSum, 'FF1F4E9C'], [K.editFirst + w, colE[w], C.orangeSum, 'FFB45309']].forEach(function (x) {
        var c = L(x[0]);
        var cell = put(ws, c + T, hm(x[1]), { fill: x[2], bold: true, num: NUM, color: x[3] });
        if (n > 0) {
          var rng = c + first + ':' + c + last;
          var sp = 'SUMPRODUCT(INT(' + rng + ')*60+ROUND((' + rng + '-INT(' + rng + '))*100,0))';
          cell.value = { formula: hmFormula(sp), result: hm(x[1]) };
        }
      });
    }
    if (n > 0) {
      ws.getCell(cMinW + T).value = { formula: 'SUM(' + cMinW + first + ':' + cMinW + last + ')', result: sumW };
      ws.getCell(cMinE + T).value = { formula: 'SUM(' + cMinE + first + ':' + cMinE + last + ')', result: sumE };
    } else {
      ws.getCell(cMinW + T).value = 0;
      ws.getCell(cMinE + T).value = 0;
    }
    summary(ws, T, K, nW, sumW, sumE, null, true);
    var cnt = L(K.count);
    ws.getCell(cnt + T).value = n > 0
      ? { formula: 'SUM(' + cnt + first + ':' + cnt + last + ')', result: formats }
      : 0;
    ws.getRow(T).height = 24;

    return T + 2; // next block starts after one blank row
  }

  /** Σ columns for one row, all derived from the hidden minute columns. */
  function summary(ws, r, K, nW, minW, minE, band, isTotal) {
    var all = minW + minE;
    var plain = isTotal ? C.yellow : band;
    var R = L(K.minW) + r, S = L(K.minE) + r;

    put(ws, L(K.sumW) + r, 0, { fill: C.blueSum, bold: true, num: NUM, color: 'FF1F4E9C' })
      .value = { formula: hmFormula(R), result: hm(minW) };
    put(ws, L(K.sumE) + r, 0, { fill: C.orangeSum, bold: true, num: NUM, color: 'FFB45309' })
      .value = { formula: hmFormula(S), result: hm(minE) };
    put(ws, L(K.total) + r, 0, { fill: C.greenSum, bold: true, num: NUM, color: 'FF2E6B1F' })
      .value = { formula: hmFormula(R + '+' + S), result: hm(all) };
    put(ws, L(K.pct) + r, 0, { fill: plain, bold: isTotal, num: '0%' })
      .value = { formula: 'IF(' + R + '+' + S + '=0,0,' + S + '/(' + R + '+' + S + '))', result: all ? minE / all : 0 };
    put(ws, L(K.avg) + r, 0, { fill: plain, bold: isTotal, italic: !isTotal, num: NUM })
      .value = { formula: hmFormula('ROUND(' + R + '/' + nW + ',0)'), result: hm(Math.round(minW / nW)) };
    put(ws, L(K.count) + r, 0, { fill: plain, bold: isTotal, num: '0' });
    if (!isTotal) {
      ws.getCell(L(K.count) + r).value = { formula: 'IF(' + R + '+' + S + '>0,1,0)', result: all > 0 ? 1 : 0 };
    }
    // the two minute columns are hidden helpers, left unstyled
  }

  function buildWorkbook(ExcelJS, data) {
    var nW = data.nw || (data.p[0] && data.p[0].t[0] ? data.p[0].t[0].w.length : 4);
    var K = layout(nW);
    var wb = new ExcelJS.Workbook();
    wb.creator = 'Трекер';
    var ws = wb.addWorksheet('Трекер', {
      views: [{ showGridLines: false }],
      pageSetup: { orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0 },
    });

    var widths = [14, 12, 30];
    var i;
    for (i = 0; i < nW; i++) widths.push(9);
    widths.push(13);
    for (i = 0; i < nW; i++) widths.push(9);
    widths.push(14, 13, 10, 14, 12, 8, 8);
    widths.forEach(function (w, idx) { ws.getColumn(idx + 1).width = w; });
    ws.getColumn(K.minW).hidden = true;
    ws.getColumn(K.minE).hidden = true;

    var row = 1;
    data.p.forEach(function (project) { row = block(ws, project, row, data, K, nW); });
    return wb;
  }

  var api = { buildWorkbook: buildWorkbook, hm: hm };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.TrackerXlsx = api;
})(typeof window !== 'undefined' ? window : globalThis);
