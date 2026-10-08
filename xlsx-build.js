/* Builds the monthly "resource tracker" workbook.
 *
 * data = { y, m, name, role, p: [ { n: "Project", t: [ { n: "Task", w: [min x4], e: [min x4] } ] } ] }
 * w / e are minutes per week (weeks 1-3 are 7 days each, week 4 is day 22 to the end of the month).
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

  function block(ws, project, r0, data) {
    var tasks = project.t;
    var n = tasks.length;
    var first = r0 + 4;
    var last = first + n - 1;
    var totalRow = first + n;
    var monthText = MONTHS[data.m - 1] + ' ' + data.y;

    // Title + hint
    ws.mergeCells(r0, 1, r0, 17);
    put(ws, 'A' + r0, '📁 ПРОЄКТ: ' + project.n.toUpperCase() + ' — ТРЕКЕР РЕСУРСОЗАТРАТНОСТІ',
      { fill: C.navy, color: C.white, bold: true, size: 14 });
    ws.getRow(r0).height = 32;

    ws.mergeCells(r0 + 1, 1, r0 + 1, 17);
    put(ws, 'A' + (r0 + 1),
      'Дані за ' + monthText + ' заповнені автоматично з трекера. Місячні підсумки рахуються автоматично. ' +
      '⏱ Час у форматі ГОД.ХВ: 0.40 = 40 хв, 1.20 = 1 год 20 хв (хвилини 00–59). ' +
      'Тиж. 1–3 = по 7 днів, Тиж. 4 = з 22-го числа до кінця місяця.',
      { italic: true, color: C.grayText, align: 'left', wrap: true, size: 10 });
    ws.getRow(r0 + 1).height = 30;

    // Group headers
    var g = r0 + 2;
    [
      [1, 3, 'A', '👤 Учасник та формат роботи', C.navy],
      [4, 8, 'D', '⏱ ОСНОВНА РОБОТА (год)', C.blue],
      [9, 13, 'I', '✏️ ПРАВКИ (год)', C.orange],
      [14, 17, 'N', '📊 ПІДСУМКИ МІСЯЦЯ', C.green],
    ].forEach(function (x) {
      for (var c = x[0]; c <= x[1]; c++) style(ws.getCell(g, c), { fill: x[4], color: C.white, bold: true });
      ws.mergeCells(g, x[0], g, x[1]);
      put(ws, x[2] + g, x[3], { fill: x[4], color: C.white, bold: true });
    });
    ws.getRow(g).height = 24;

    // Column headers
    var h = r0 + 3;
    var heads = [
      ["Ім'я", C.navy], ['Роль', C.navy], ['Формат роботи', C.navy],
      ['Тиж. 1', C.blue], ['Тиж. 2', C.blue], ['Тиж. 3', C.blue], ['Тиж. 4', C.blue], ['Σ Осн. (міс.)', C.blue],
      ['Тиж. 1', C.orange], ['Тиж. 2', C.orange], ['Тиж. 3', C.orange], ['Тиж. 4', C.orange], ['Σ Правок (міс.)', C.orange],
      ['Σ Всього (міс.)', C.green], ['% Правок', C.green], ['Сер./тиж. (осн.)', C.green], ['К-сть форматів', C.green],
    ];
    heads.forEach(function (x, i) {
      put(ws, ws.getCell(h, i + 1).address, x[0], { fill: x[1], color: C.white, bold: true, size: 10, wrap: true });
    });
    ws.getRow(h).height = 34;

    // Task rows
    var sumW = 0, sumE = 0, formats = 0;
    var colW = [0, 0, 0, 0], colE = [0, 0, 0, 0];
    tasks.forEach(function (t, i) {
      var r = first + i;
      var band = i % 2 === 0 ? C.blueBand : C.white;
      var minW = sumMinutes(t.w), minE = sumMinutes(t.e);
      sumW += minW; sumE += minE; formats += minW + minE > 0 ? 1 : 0;
      t.w.forEach(function (v, k) { colW[k] += v; });
      t.e.forEach(function (v, k) { colE[k] += v; });

      put(ws, 'A' + r, i === 0 ? data.name : '', { fill: band, bold: true, color: C.navy, align: 'left' });
      put(ws, 'B' + r, i === 0 ? data.role : '', { fill: band });
      put(ws, 'C' + r, t.n, { fill: band, align: 'left' });

      ['D', 'E', 'F', 'G'].forEach(function (c, k) {
        put(ws, c + r, hm(t.w[k]), { fill: band, num: NUM, color: 'FF1F4E9C' });
      });
      ['I', 'J', 'K', 'L'].forEach(function (c, k) {
        put(ws, c + r, hm(t.e[k]), { fill: band, num: NUM, color: 'FFB45309' });
      });

      ws.getCell('R' + r).value = {
        formula: [term('D' + r), term('E' + r), term('F' + r), term('G' + r)].join('+'), result: minW,
      };
      ws.getCell('S' + r).value = {
        formula: [term('I' + r), term('J' + r), term('K' + r), term('L' + r)].join('+'), result: minE,
      };
      summary(ws, r, minW, minE, band, false);
    });

    // Total row
    var T = totalRow;
    var dark = { fill: C.navy, color: C.white, bold: true };
    put(ws, 'A' + T, data.name, Object.assign({ align: 'left' }, dark));
    put(ws, 'B' + T, data.role === 'Лід' ? 'Менеджер' : data.role, dark);
    put(ws, 'C' + T, 'РАЗОМ', dark);

    ['D', 'E', 'F', 'G', 'I', 'J', 'K', 'L'].forEach(function (c) {
      var isWork = 'DEFG'.indexOf(c) >= 0;
      var k = isWork ? 'DEFG'.indexOf(c) : 'IJKL'.indexOf(c);
      var minutes = isWork ? colW[k] : colE[k];
      var cell = put(ws, c + T, hm(minutes), {
        fill: isWork ? C.blueSum : C.orangeSum, bold: true, num: NUM, color: isWork ? 'FF1F4E9C' : 'FFB45309',
      });
      if (n > 0) {
        var rng = c + first + ':' + c + last;
        var sp = 'SUMPRODUCT(INT(' + rng + ')*60+ROUND((' + rng + '-INT(' + rng + '))*100,0))';
        cell.value = { formula: hmFormula(sp), result: hm(minutes) };
      }
    });
    if (n > 0) {
      ws.getCell('R' + T).value = { formula: 'SUM(R' + first + ':R' + last + ')', result: sumW };
      ws.getCell('S' + T).value = { formula: 'SUM(S' + first + ':S' + last + ')', result: sumE };
    } else {
      ws.getCell('R' + T).value = 0;
      ws.getCell('S' + T).value = 0;
    }
    summary(ws, T, sumW, sumE, null, true);
    ws.getCell('Q' + T).value = n > 0
      ? { formula: 'SUM(Q' + first + ':Q' + last + ')', result: formats }
      : 0;
    ws.getRow(T).height = 24;

    return T + 2; // next block starts after one blank row
  }

  /** Σ columns for one row, all derived from the hidden minute columns R (work) and S (edits). */
  function summary(ws, r, minW, minE, band, isTotal) {
    var all = minW + minE;
    var plain = isTotal ? C.yellow : band;
    var R = 'R' + r, S = 'S' + r;

    put(ws, 'H' + r, 0, { fill: C.blueSum, bold: true, num: NUM, color: 'FF1F4E9C' })
      .value = { formula: hmFormula(R), result: hm(minW) };
    put(ws, 'M' + r, 0, { fill: C.orangeSum, bold: true, num: NUM, color: 'FFB45309' })
      .value = { formula: hmFormula(S), result: hm(minE) };
    put(ws, 'N' + r, 0, { fill: C.greenSum, bold: true, num: NUM, color: 'FF2E6B1F' })
      .value = { formula: hmFormula(R + '+' + S), result: hm(all) };
    put(ws, 'O' + r, 0, { fill: plain, bold: isTotal, num: '0%' })
      .value = { formula: 'IF(' + R + '+' + S + '=0,0,' + S + '/(' + R + '+' + S + '))', result: all ? minE / all : 0 };
    put(ws, 'P' + r, 0, { fill: plain, bold: isTotal, italic: !isTotal, num: NUM })
      .value = { formula: hmFormula('ROUND(' + R + '/4,0)'), result: hm(Math.round(minW / 4)) };
    put(ws, 'Q' + r, 0, { fill: plain, bold: isTotal, num: '0' });
    if (!isTotal) {
      ws.getCell('Q' + r).value = { formula: 'IF(' + R + '+' + S + '>0,1,0)', result: all > 0 ? 1 : 0 };
    }
    // R and S (minutes) are hidden helper columns, left unstyled
  }

  function buildWorkbook(ExcelJS, data) {
    var wb = new ExcelJS.Workbook();
    wb.creator = 'Трекер';
    var ws = wb.addWorksheet('Трекер', {
      views: [{ showGridLines: false }],
      pageSetup: { orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0 },
    });
    var widths = [14, 12, 30, 9, 9, 9, 9, 13, 9, 9, 9, 9, 14, 13, 10, 14, 12, 8, 8];
    widths.forEach(function (w, i) { ws.getColumn(i + 1).width = w; });
    ws.getColumn(18).hidden = true;
    ws.getColumn(19).hidden = true;

    var row = 1;
    data.p.forEach(function (project) { row = block(ws, project, row, data); });
    return wb;
  }

  var api = { buildWorkbook: buildWorkbook, hm: hm };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.TrackerXlsx = api;
})(typeof window !== 'undefined' ? window : globalThis);
