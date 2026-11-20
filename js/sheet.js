/* 스프레드시트(Excel) 편집기 화면 */
window.DV = window.DV || {};
(function (DV) {
  'use strict';

  const F = function () { return DV.formula; };
  const ROWHDR_W = 48, HDR_H = 24;
  const MAX_ROWS = 1048576, MAX_COLS = 16384;
  const NUM_FORMATS = [['', '일반'], ['#,##0', '숫자 (1,234)'], ['#,##0.00', '숫자 (1,234.00)'], ['0', '정수 (1234)'], ['"₩"#,##0', '통화 (₩)'], ['$#,##0.00', '통화 ($)'],
    ['0%', '백분율 (12%)'], ['0.00%', '백분율 (12.34%)'], ['yyyy-mm-dd', '날짜 (2024-01-31)'], ['yyyy"년" m"월" d"일"', '날짜 (2024년 1월 31일)'], ['h:mm:ss', '시간'], ['0.00E+00', '지수'], ['@', '텍스트']];
  const COLORS = ['000000', '444444', '888888', 'CCCCCC', 'FFFFFF', 'C00000', 'FF0000', 'FFC000', 'FFFF00', '92D050',
    '00B050', '00B0F0', '0070C0', '002060', '7030A0', 'F8CBAD', 'FFE699', 'E2EFDA', 'DDEBF7', 'EDE7F6'];

  /* ---------- 입력 해석 ---------- */
  DV.sheetInput = {
    parse: function (text, base) {
      const cell = base ? Object.assign({}, base) : {};
      delete cell.v; delete cell.f; delete cell.e;
      if (text == null || text === '') return Object.keys(cell).length ? cell : undefined;
      const t = String(text);
      if (cell.z === '@') { cell.v = t; return cell; }
      if (t[0] === "'") { cell.v = t.slice(1); return cell; }
      if (t[0] === '=' && t.length > 1) { cell.f = t.slice(1); return cell; }
      const s = t.trim();
      let m;
      if (/^[-+]?(\d{1,3}(,\d{3})+|\d+)?(\.\d+)?([eE][-+]?\d+)?$/.test(s) && /\d/.test(s)) { cell.v = parseFloat(s.replace(/,/g, '')); if (/,/.test(s) && !cell.z) cell.z = /\./.test(s) ? '#,##0.00' : '#,##0'; return cell; }
      if ((m = /^([-+]?[\d,]*\.?\d+)\s*%$/.exec(s))) { cell.v = parseFloat(m[1].replace(/,/g, '')) / 100; if (!cell.z) cell.z = /\./.test(m[1]) ? '0.00%' : '0%'; return cell; }
      if ((m = /^([₩$])\s?([-+]?[\d,]*\.?\d+)$/.exec(s))) { cell.v = parseFloat(m[2].replace(/,/g, '')); if (!cell.z) cell.z = m[1] === '₩' ? '"₩"#,##0' : '$#,##0.00'; return cell; }
      if (/^(TRUE|FALSE)$/i.test(s)) { cell.v = s.toUpperCase() === 'TRUE'; return cell; }
      if ((m = /^(\d{4})[-./](\d{1,2})[-./](\d{1,2})\.?$/.exec(s)) && +m[2] >= 1 && +m[2] <= 12 && +m[3] >= 1 && +m[3] <= 31) {
        cell.v = F().dateSerial(+m[1], +m[2], +m[3]); if (!cell.z || !/[ymd]/i.test(cell.z)) cell.z = 'yyyy-mm-dd'; return cell;
      }
      if ((m = /^(\d{1,2}):(\d{2})(?::(\d{2}))?$/.exec(s)) && +m[1] < 24 && +m[2] < 60) {
        cell.v = (+m[1] * 3600 + +m[2] * 60 + (+m[3] || 0)) / 86400; if (!cell.z) cell.z = m[3] ? 'h:mm:ss' : 'h:mm'; return cell;
      }
      cell.v = t;
      return cell;
    }
  };

  function clone(x) { return x === undefined ? undefined : (typeof structuredClone === 'function' ? structuredClone(x) : JSON.parse(JSON.stringify(x))); }

  function fmtGeneral(n) {
    if (!isFinite(n)) return '#NUM!';
    const a = Math.abs(n);
    if (a !== 0 && (a >= 1e11 || a < 1e-9)) return n.toExponential(5).replace(/\.?0+e/, 'E').replace('e', 'E').replace(/E\+?/, 'E+').replace('E+-', 'E-');
    if (Number.isInteger(n)) return String(n);
    let s = String(parseFloat(n.toPrecision(11)));
    return s;
  }

  function createSheetEditor(container, book, opts) {
    opts = opts || {};
    const el = document.createElement('div');
    el.className = 'sheet-view';
    const tb = function (cmd, title, inner) { return '<button type="button" class="tb" data-cmd="' + cmd + '" title="' + title + '">' + inner + '</button>'; };
    el.innerHTML =
      '<div class="toolbar" role="toolbar">' +
        '<div class="tb-group">' + tb('undo', '실행 취소 (Ctrl+Z)', DV.icon('undo')) + tb('redo', '다시 실행 (Ctrl+Y)', DV.icon('redo')) + '</div>' +
        '<div class="tb-group"><select class="tb-select size" data-cmd="size" title="글자 크기"><option value="">크기</option>' + [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36].map(function (s) { return '<option>' + s + '</option>'; }).join('') + '</select>' +
          tb('b', '굵게 (Ctrl+B)', '<b>가</b>') + tb('i', '기울임 (Ctrl+I)', '<i style="font-family:serif">가</i>') + tb('u', '밑줄 (Ctrl+U)', '<u>가</u>') + tb('strike', '취소선', '<s>가</s>') +
          '<span class="color-btn"><button type="button" class="tb" data-cmd="color" title="글자 색"><span class="glyph">가</span><span class="bar fore" style="background:#C00000"></span></button><button type="button" class="tb caret-sm" data-cmd="pal-color">▾</button></span>' +
          '<span class="color-btn"><button type="button" class="tb" data-cmd="bg" title="채우기 색">' + DV.icon('fill') + '<span class="bar back" style="background:#FFE699"></span></button><button type="button" class="tb caret-sm" data-cmd="pal-bg">▾</button></span>' +
          tb('border', '모든 테두리 켜기/끄기', '<svg viewBox="0 0 24 24"><rect x="4" y="4" width="16" height="16"/><path d="M4 12h16M12 4v16"/></svg>') +
        '</div>' +
        '<div class="tb-group">' + tb('left', '왼쪽 맞춤', DV.icon('left')) + tb('center', '가운데 맞춤', DV.icon('center')) + tb('right', '오른쪽 맞춤', DV.icon('right')) +
          tb('wrap', '자동 줄 바꿈', DV.icon('wrap')) + tb('merge', '셀 병합/해제', DV.icon('merge')) + '</div>' +
        '<div class="tb-group"><select class="tb-select fmt" data-cmd="fmt" title="표시 형식">' + NUM_FORMATS.map(function (f) { return '<option value="' + DV.esc(f[0]) + '">' + f[1] + '</option>'; }).join('') + '</select>' +
          tb('dec+', '소수 자릿수 늘리기', '.0<sup>+</sup>') + tb('dec-', '소수 자릿수 줄이기', '.0<sup>−</sup>') + '</div>' +
        '<div class="tb-group">' + tb('sum', '자동 합계', DV.icon('sum')) + tb('sortasc', '오름차순 정렬', DV.icon('sortasc')) + tb('sortdesc', '내림차순 정렬', DV.icon('sortdesc')) + '</div>' +
        '<div class="tb-group">' + tb('insrow', '행 삽입', '행+') + tb('delrow', '행 삭제', '행−') + tb('inscol', '열 삽입', '열+') + tb('delcol', '열 삭제', '열−') + '</div>' +
      '</div>' +
      '<div class="formula-bar"><input class="namebox" type="text" spellcheck="false" title="셀 주소 (입력 후 Enter로 이동)"><span class="fx">fx</span><input class="fbar" type="text" spellcheck="false" autocomplete="off" placeholder="값 또는 =수식 입력"></div>' +
      '<div class="grid-wrap" tabindex="-1">' +
        '<div class="colhdr-row"><div class="corner" title="전체 선택"></div><div class="colhdrs"></div></div>' +
        '<div class="gbody"><div class="rowhdrs"></div><div class="cells"></div><div class="sel-box"></div><div class="ref-box" hidden></div>' +
          '<textarea class="cell-input" spellcheck="false" autocomplete="off" autocapitalize="off" aria-label="셀 입력"></textarea></div>' +
      '</div>' +
      '<div class="sheet-tabs"><button type="button" class="tab-add" title="시트 추가">+</button><div class="tab-list"></div></div>' +
      '<div class="statusbar"><span class="st-kind"></span><span class="st-sel"></span></div>' +
      '<div class="palette" hidden></div><div class="ctxmenu" hidden></div>';
    container.appendChild(el);

    const wrap = el.querySelector('.grid-wrap');
    const gbody = el.querySelector('.gbody');
    const cellsLayer = el.querySelector('.cells');
    const rowHdrs = el.querySelector('.rowhdrs');
    const colHdrs = el.querySelector('.colhdrs');
    const colHdrRow = el.querySelector('.colhdr-row');
    const selBox = el.querySelector('.sel-box');
    const refBox = el.querySelector('.ref-box');
    const input = el.querySelector('.cell-input');
    const fbar = el.querySelector('.fbar');
    const namebox = el.querySelector('.namebox');
    const tabList = el.querySelector('.tab-list');
    const palette = el.querySelector('.palette');
    const ctxmenu = el.querySelector('.ctxmenu');
    const toolbar = el.querySelector('.toolbar');
    const stSel = el.querySelector('.st-sel');
    el.querySelector('.st-kind').textContent = opts.kindLabel || '';

    let si = 0;
    let sel = { r1: 0, c1: 0, r2: 0, c2: 0 };
    let act = { r: 0, c: 0 };
    let editing = null;
    let rowCount = 100, colCount = 26;
    let rowTops = [], colLefts = [];
    let mergeMap = new Map();
    let calcDirty = true;
    const vcache = new Map();
    const evaluating = new Set();
    const history = [], future = [];
    const listeners = [];
    let clip = null;
    let foreColor = 'C00000', backColor = 'FFE699';
    let lastWin = null;

    function S() { return book.sheets[si]; }
    function sheetIndex(name) {
      if (name == null) return -1;
      const n = name.toLowerCase();
      for (let i = 0; i < book.sheets.length; i++) if (book.sheets[i].name.toLowerCase() === n) return i;
      return -1;
    }
    function getCell(sh, r, c) { const row = sh.rows[r]; return row ? row[c] : undefined; }
    function isEmptyCell(c) { return !c || ((c.v === undefined || c.v === null || c.v === '') && !c.f && !c.s && !c.z); }
    function setCellRaw(sh, r, c, cell) {
      if (isEmptyCell(cell)) { if (sh.rows[r]) delete sh.rows[r][c]; }
      else (sh.rows[r] = sh.rows[r] || [])[c] = cell;
    }
    function used(sh) {
      let rows = 0, cols = 0;
      sh.rows.forEach(function (row, r) {
        if (!row) return;
        let has = false;
        row.forEach(function (cell, c) { if (cell) { has = true; if (c + 1 > cols) cols = c + 1; } });
        if (has) rows = r + 1;
      });
      sh.merges.forEach(function (m) { rows = Math.max(rows, m.r2 + 1); cols = Math.max(cols, m.c2 + 1); });
      return { rows: rows, cols: cols };
    }

    /* ---------- 계산 ---------- */
    function rawValue(cell) {
      if (!cell) return null;
      if (cell.e) return new (F().Err)(cell.v);
      if (cell.v === '' || cell.v === undefined) return null;
      return cell.v;
    }
    function hasUnknownFn(ast) {
      if (!ast || typeof ast !== 'object') return false;
      if (ast.k === 'fn' && F().functions.indexOf(ast.name) < 0) return true;
      return ['l', 'r', 'x'].some(function (k) { return ast[k] && hasUnknownFn(ast[k]); }) || (ast.args || []).some(hasUnknownFn);
    }
    const extCache = {};
    function makeCtx(sIdx, r, c) {
      return {
        curRow: r, curCol: c,
        hasSheet: function (name) { return sheetIndex(name) >= 0; },
        cell: function (name, rr, cc) { return value(name == null ? sIdx : sheetIndex(name), rr, cc); },
        extent: function (name) {
          const i = name == null ? sIdx : sheetIndex(name);
          if (!extCache[i]) extCache[i] = used(book.sheets[i]);
          return extCache[i];
        }
      };
    }
    function value(sIdx, r, c) {
      if (sIdx < 0 || !book.sheets[sIdx]) return F().E.REF;
      const cell = getCell(book.sheets[sIdx], r, c);
      if (!cell || !cell.f) return rawValue(cell);
      const key = sIdx * 1e10 + r * 20000 + c;
      if (vcache.has(key)) return vcache.get(key);
      if (evaluating.has(key)) return F().E.CIRC;
      evaluating.add(key);
      let v;
      try {
        const ast = F().compile(cell.f);
        if (hasUnknownFn(ast) && cell.v !== undefined) v = rawValue(cell);
        else {
          v = F().evaluate(ast, makeCtx(sIdx, r, c));
          if (v instanceof F().Range) v = v.rows() === 1 && v.cols() === 1 ? v.get(0, 0) : F().E.VALUE;
          // 해석할 수 없는 수식(다른 프로그램 전용 문법 등)은 파일에 저장된 값을 보여 줍니다
          if (v === F().E.NAME && cell.v !== undefined) v = rawValue(cell);
        }
      } catch (e) {
        v = e instanceof RangeError ? F().E.CIRC : F().E.VALUE;
      }
      evaluating.delete(key);
      if (v === undefined) v = null;
      vcache.set(key, v);
      return v;
    }
    function recalc() {
      if (!calcDirty) return;
      vcache.clear();
      Object.keys(extCache).forEach(function (k) { delete extCache[k]; });
      calcDirty = false;
      // 위에서 아래로 미리 계산해 깊은 재귀를 피합니다
      book.sheets.forEach(function (sh, i) {
        sh.rows.forEach(function (row, r) {
          if (row) row.forEach(function (cell, c) { if (cell && cell.f) value(i, r, c); });
        });
      });
    }
    function invalidate() { calcDirty = true; }

    function display(cell, v) {
      if (v == null) return '';
      if (v instanceof F().Err) return v.e;
      if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
      if (typeof v === 'number') {
        if (cell && cell.z && cell.z !== '@') {
          try { return XLSX.SSF.format(cell.z, v); } catch (e) { return fmtGeneral(v); }
        }
        return fmtGeneral(v);
      }
      return String(v);
    }
    function rawText(cell) {
      if (!cell) return '';
      if (cell.f) return '=' + cell.f;
      if (cell.e) return cell.v;
      if (typeof cell.v === 'boolean') return cell.v ? 'TRUE' : 'FALSE';
      if (typeof cell.v === 'number') {
        if (cell.z && /[ymd]/i.test(cell.z.replace(/"[^"]*"/g, '')) && !/^\[?h/i.test(cell.z)) { try { return XLSX.SSF.format('yyyy-mm-dd', cell.v); } catch (e) { /* 무시 */ } }
        if (cell.z && /%/.test(cell.z)) return fmtGeneral(Math.round(cell.v * 100 * 1e9) / 1e9) + '%';
        return fmtGeneral(cell.v);
      }
      return cell.v == null ? '' : String(cell.v);
    }

    /* ---------- 배치 ---------- */
    function colW(c) { const w = S().colWidths[c]; return w == null ? DV.xlsxIO.DEFAULT_COL_PX : w; }
    function rowH(r) { const h = S().rowHeights[r]; return h == null ? DV.xlsxIO.DEFAULT_ROW_PX : h; }
    function layout() {
      const u = used(S());
      rowCount = Math.min(MAX_ROWS, Math.max(rowCount, u.rows + 50, 100, act.r + 30));
      colCount = Math.min(MAX_COLS, Math.max(colCount, u.cols + 6, 26, act.c + 6));
      rowTops = new Array(rowCount + 1);
      rowTops[0] = 0;
      for (let r = 0; r < rowCount; r++) rowTops[r + 1] = rowTops[r] + rowH(r);
      colLefts = new Array(colCount + 1);
      colLefts[0] = 0;
      for (let c = 0; c < colCount; c++) colLefts[c + 1] = colLefts[c] + colW(c);
      const W = ROWHDR_W + colLefts[colCount], H = rowTops[rowCount];
      gbody.style.width = W + 'px';
      gbody.style.height = H + 'px';
      colHdrRow.style.width = W + 'px';
      rowHdrs.style.height = H + 'px';
      mergeMap = new Map();
      S().merges.forEach(function (m) {
        for (let r = m.r1; r <= m.r2; r++) for (let c = m.c1; c <= m.c2; c++) mergeMap.set(r * 20000 + c, m);
      });
    }
    function bsearch(arr, x) {
      let lo = 0, hi = arr.length - 2;
      while (lo < hi) { const mid = (lo + hi + 1) >> 1; if (arr[mid] <= x) lo = mid; else hi = mid - 1; }
      return lo;
    }
    function cellRect(r1, c1, r2, c2) {
      if (r2 == null) { r2 = r1; c2 = c1; }
      const m = mergeMap.get(r1 * 20000 + c1);
      if (m && r1 === r2 && c1 === c2) { r1 = m.r1; c1 = m.c1; r2 = m.r2; c2 = m.c2; }
      ensureSize(Math.max(r1, r2), Math.max(c1, c2));
      return { left: ROWHDR_W + colLefts[Math.min(c1, c2)], top: rowTops[Math.min(r1, r2)], width: colLefts[Math.max(c1, c2) + 1] - colLefts[Math.min(c1, c2)], height: rowTops[Math.max(r1, r2) + 1] - rowTops[Math.min(r1, r2)] };
    }
    function ensureSize(r, c) {
      if (r >= rowCount - 5 || c >= colCount - 2) {
        rowCount = Math.min(MAX_ROWS, Math.max(rowCount, r + 50));
        colCount = Math.min(MAX_COLS, Math.max(colCount, c + 6));
        layout();
      }
    }

    /* ---------- 그리기 ---------- */
    function cellStyle(cell, v) {
      const css = [];
      const s = cell && cell.s;
      let align = s && s.align;
      if (!align) align = typeof v === 'number' ? 'right' : (typeof v === 'boolean' || v instanceof F().Err) ? 'center' : null;
      if (align === 'right') css.push('justify-content:flex-end;text-align:right');
      else if (align === 'center') css.push('justify-content:center;text-align:center');
      if (s) {
        if (s.b) css.push('font-weight:700');
        if (s.i) css.push('font-style:italic');
        if (s.u || s.strike) css.push('text-decoration:' + (s.u ? 'underline ' : '') + (s.strike ? 'line-through' : ''));
        if (s.color) css.push('color:#' + s.color);
        if (s.bg) css.push('background:#' + s.bg);
        if (s.size) css.push('font-size:' + (s.size * 4 / 3).toFixed(1) + 'px');
        if (s.valign === 'top') css.push('align-items:flex-start');
        else if (s.valign === 'middle') css.push('align-items:center');
        if (s.wrap) css.push('white-space:pre-wrap;word-break:break-all');
        if (s.border) {
          ['left', 'right', 'top', 'bottom'].forEach(function (k) { if (s.border[k]) css.push('border-' + k + ':1px solid #222'); });
        }
      }
      return css.join(';');
    }

    function render() {
      recalc();
      const vw = wrap.clientWidth, vh = wrap.clientHeight - HDR_H;
      const st = wrap.scrollTop, sl = wrap.scrollLeft;
      if (st + vh > rowTops[rowCount] - 200 && rowCount < MAX_ROWS) { rowCount = Math.min(MAX_ROWS, rowCount + 200); layout(); }
      if (sl + vw > colLefts[colCount] + ROWHDR_W - 200 && colCount < MAX_COLS) { colCount = Math.min(MAX_COLS, colCount + 10); layout(); }
      const r0 = Math.max(0, bsearch(rowTops, st) - 3);
      const r1 = Math.min(rowCount - 1, bsearch(rowTops, st + vh) + 3);
      const c0 = Math.max(0, bsearch(colLefts, Math.max(0, sl - ROWHDR_W)) - 1);
      const c1 = Math.min(colCount - 1, bsearch(colLefts, sl + vw) + 1);
      lastWin = { r0: r0, r1: r1, c0: c0, c1: c1 };
      const sh = S();
      let h = '';
      const drawnMerges = new Set();
      for (let r = r0; r <= r1; r++) {
        const top = rowTops[r], hh = rowTops[r + 1] - top;
        if (hh <= 0) continue;
        const row = sh.rows[r];
        for (let c = c0; c <= c1; c++) {
          const ww = colLefts[c + 1] - colLefts[c];
          if (ww <= 0) continue;
          let m = mergeMap.get(r * 20000 + c);
          let left = ROWHDR_W + colLefts[c], width = ww, height = hh, rr = r, cc = c, ttop = top;
          if (m) {
            if (drawnMerges.has(m)) continue;
            drawnMerges.add(m);
            rr = m.r1; cc = m.c1;
            left = ROWHDR_W + colLefts[m.c1]; ttop = rowTops[m.r1];
            width = colLefts[m.c2 + 1] - colLefts[m.c1]; height = rowTops[m.r2 + 1] - rowTops[m.r1];
          }
          const cell = m ? getCell(sh, rr, cc) : (row ? row[c] : undefined);
          const v = cell ? value(si, rr, cc) : null;
          const text = cell ? display(cell, v) : '';
          let style = 'left:' + left + 'px;top:' + ttop + 'px;width:' + width + 'px;height:' + height + 'px;';
          let cls = 'c';
          if (m) cls += ' merged';
          if (cell) {
            style += cellStyle(cell, v);
            // 오른쪽이 비어 있으면 글자가 넘쳐 보이게
            if (text && typeof v === 'string' && !(cell.s && (cell.s.wrap || cell.s.align === 'center' || cell.s.align === 'right')) && !m) {
              let k = c + 1, extra = 0;
              while (k < colCount && k <= c + 12 && !(row && row[k] && (row[k].v != null && row[k].v !== '' || row[k].f)) && !mergeMap.has(r * 20000 + k)) { extra += colLefts[k + 1] - colLefts[k]; k++; }
              if (extra) cls += ' spill" data-spill="' + (width + extra);
            }
          }
          h += '<div class="' + cls + '" data-r="' + rr + '" data-c="' + cc + '" style="' + style + '">' + (text ? '<span>' + DV.esc(text) + '</span>' : '') + '</div>';
        }
      }
      cellsLayer.innerHTML = h;
      cellsLayer.querySelectorAll('.spill').forEach(function (d) {
        const sp = d.firstChild;
        if (sp && sp.scrollWidth > d.clientWidth) { d.style.width = d.dataset.spill + 'px'; d.style.zIndex = 1; d.classList.add('spilled'); }
      });
      // 머리글
      let ch = '';
      for (let c = c0; c <= c1; c++) {
        const w = colLefts[c + 1] - colLefts[c];
        if (w <= 0) continue;
        const on = c >= sel.c1 && c <= sel.c2;
        ch += '<div class="ch' + (on ? ' on' : '') + '" data-c="' + c + '" style="left:' + (ROWHDR_W + colLefts[c]) + 'px;width:' + w + 'px">' + F().colName(c) + '<i class="rz" data-rzc="' + c + '"></i></div>';
      }
      colHdrs.innerHTML = ch;
      let rh = '';
      for (let r = r0; r <= r1; r++) {
        const hh = rowTops[r + 1] - rowTops[r];
        if (hh <= 0) continue;
        const on = r >= sel.r1 && r <= sel.r2;
        rh += '<div class="rh' + (on ? ' on' : '') + '" data-r="' + r + '" style="top:' + rowTops[r] + 'px;height:' + hh + 'px">' + (r + 1) + '<i class="rz" data-rzr="' + r + '"></i></div>';
      }
      rowHdrs.innerHTML = rh;
      paintSel();
    }
    let raf = 0;
    function schedule() { if (!raf) raf = requestAnimationFrame(function () { raf = 0; render(); }); }
    wrap.addEventListener('scroll', function () { schedule(); });
    const ro = typeof ResizeObserver !== 'undefined' ? new ResizeObserver(function () { schedule(); }) : null;
    if (ro) ro.observe(wrap);

    function paintSel() {
      const r = cellRect(sel.r1, sel.c1, sel.r2, sel.c2);
      selBox.style.cssText = 'left:' + (r.left - 1) + 'px;top:' + (r.top - 1) + 'px;width:' + (r.width + 1) + 'px;height:' + (r.height + 1) + 'px';
      selBox.classList.toggle('multi', !(sel.r1 === sel.r2 && sel.c1 === sel.c2) && !isSingleMerged());
      const a = cellRect(act.r, act.c);
      selBox.style.setProperty('--ax', (a.left - r.left + 1) + 'px');
      selBox.style.setProperty('--ay', (a.top - r.top + 1) + 'px');
      selBox.style.setProperty('--aw', a.width + 'px');
      selBox.style.setProperty('--ah', a.height + 'px');
      colHdrs.querySelectorAll('.ch').forEach(function (d) { const c = +d.dataset.c; d.classList.toggle('on', c >= sel.c1 && c <= sel.c2); });
      rowHdrs.querySelectorAll('.rh').forEach(function (d) { const rr = +d.dataset.r; d.classList.toggle('on', rr >= sel.r1 && rr <= sel.r2); });
      namebox.value = sel.r1 === sel.r2 && sel.c1 === sel.c2 || isSingleMerged() ? F().addr(act.r, act.c) : F().addr(sel.r1, sel.c1) + ':' + F().addr(sel.r2, sel.c2);
      if (!editing) {
        fbar.value = rawText(getCell(S(), act.r, act.c));
        placeInput();
      }
      updateToolbarState();
      updateStatus();
    }
    function isSingleMerged() {
      const m = mergeMap.get(sel.r1 * 20000 + sel.c1);
      return m && m.r1 === sel.r1 && m.c1 === sel.c1 && m.r2 === sel.r2 && m.c2 === sel.c2;
    }
    function placeInput() {
      const r = cellRect(act.r, act.c);
      input.style.left = r.left + 'px';
      input.style.top = r.top + 'px';
      input.style.minWidth = r.width + 'px';
      input.style.width = r.width + 'px';
      input.style.height = r.height + 'px';
    }
    const updateStatus = DV.debounce(function () {
      if (sel.r1 === sel.r2 && sel.c1 === sel.c2) { stSel.textContent = ''; return; }
      const u = used(S());
      let sum = 0, n = 0, cnt = 0;
      for (let r = sel.r1; r <= Math.min(sel.r2, u.rows - 1); r++) for (let c = sel.c1; c <= Math.min(sel.c2, u.cols - 1); c++) {
        const v = value(si, r, c);
        if (v != null && v !== '') cnt++;
        if (typeof v === 'number') { sum += v; n++; }
      }
      stSel.textContent = n ? '평균: ' + fmtGeneral(sum / n) + '   개수: ' + cnt + '   합계: ' + fmtGeneral(sum) : (cnt ? '개수: ' + cnt : '');
    }, 120);

    function updateToolbarState() {
      const cell = getCell(S(), act.r, act.c);
      const s = (cell && cell.s) || {};
      ['b', 'i', 'u', 'strike', 'wrap'].forEach(function (k) { const b = toolbar.querySelector('[data-cmd="' + k + '"]'); if (b) b.classList.toggle('on', !!s[k]); });
      ['left', 'center', 'right'].forEach(function (k) { toolbar.querySelector('[data-cmd="' + k + '"]').classList.toggle('on', s.align === k); });
      toolbar.querySelector('[data-cmd="merge"]').classList.toggle('on', !!mergeMap.get(act.r * 20000 + act.c));
      const fsel = toolbar.querySelector('[data-cmd="fmt"]');
      const z = (cell && cell.z) || '';
      fsel.value = NUM_FORMATS.some(function (f) { return f[0] === z; }) ? z : '';
      toolbar.querySelector('[data-cmd="size"]').value = s.size ? String(s.size) : '';
    }

    function scrollIntoView(r, c) {
      const rc = cellRect(r, c);
      const vw = wrap.clientWidth, vh = wrap.clientHeight;
      if (rc.top < wrap.scrollTop) wrap.scrollTop = rc.top;
      else if (rc.top + rc.height > wrap.scrollTop + vh - HDR_H - 2) wrap.scrollTop = rc.top + rc.height - vh + HDR_H + 4;
      if (rc.left - ROWHDR_W < wrap.scrollLeft) wrap.scrollLeft = rc.left - ROWHDR_W;
      else if (rc.left + rc.width > wrap.scrollLeft + vw - 2) wrap.scrollLeft = rc.left + rc.width - vw + 4;
    }

    function expandMerges(s) {
      let changed = true;
      while (changed) {
        changed = false;
        S().merges.forEach(function (m) {
          if (m.r2 < s.r1 || m.r1 > s.r2 || m.c2 < s.c1 || m.c1 > s.c2) return;
          if (m.r1 < s.r1) { s.r1 = m.r1; changed = true; }
          if (m.c1 < s.c1) { s.c1 = m.c1; changed = true; }
          if (m.r2 > s.r2) { s.r2 = m.r2; changed = true; }
          if (m.c2 > s.c2) { s.c2 = m.c2; changed = true; }
        });
      }
      return s;
    }
    function setSel(r1, c1, r2, c2, keepAct) {
      r1 = Math.max(0, Math.min(MAX_ROWS - 1, r1)); r2 = Math.max(0, Math.min(MAX_ROWS - 1, r2));
      c1 = Math.max(0, Math.min(MAX_COLS - 1, c1)); c2 = Math.max(0, Math.min(MAX_COLS - 1, c2));
      if (!keepAct) act = { r: r1, c: c1 };
      sel = expandMerges({ r1: Math.min(r1, r2), c1: Math.min(c1, c2), r2: Math.max(r1, r2), c2: Math.max(c1, c2) });
      ensureSize(Math.max(r1, r2), Math.max(c1, c2));
      paintSel();
    }
    function moveAct(dr, dc, extend) {
      if (extend) {
        // 활성 셀 기준으로 범위의 반대쪽 끝을 움직입니다
        let er = sel.r1 === act.r ? sel.r2 : sel.r1;
        let ec = sel.c1 === act.c ? sel.c2 : sel.c1;
        er = Math.max(0, er + dr); ec = Math.max(0, ec + dc);
        sel = expandMerges({ r1: Math.min(act.r, er), c1: Math.min(act.c, ec), r2: Math.max(act.r, er), c2: Math.max(act.c, ec) });
        ensureSize(er, ec);
        paintSel();
        scrollIntoView(er, ec);
        return;
      }
      let r = act.r, c = act.c;
      const m = mergeMap.get(r * 20000 + c);
      if (m) { if (dr > 0) r = m.r2; if (dc > 0) c = m.c2; if (dr < 0) r = m.r1; if (dc < 0) c = m.c1; }
      r = Math.max(0, r + dr); c = Math.max(0, c + dc);
      const m2 = mergeMap.get(r * 20000 + c);
      if (m2) { r = m2.r1; c = m2.c1; }
      setSel(r, c, r, c);
      scrollIntoView(r, c);
    }
    function jumpEdge(dr, dc) {
      const sh = S();
      const has = function (r, c) { const x = getCell(sh, r, c); return x && (x.f || (x.v != null && x.v !== '')); };
      let r = act.r, c = act.c;
      const u = used(sh);
      const maxR = Math.max(u.rows, act.r + 1), maxC = Math.max(u.cols, act.c + 1);
      const step = function () { r += dr; c += dc; };
      const inside = function () { return r >= 0 && c >= 0 && r < maxR && c < maxC; };
      step();
      if (!inside()) { r -= dr; c -= dc; }
      else if (has(act.r, act.c) && has(r, c)) { while (inside() && has(r, c)) step(); r -= dr; c -= dc; }
      else { while (inside() && !has(r, c)) step(); if (!inside()) { r -= dr; c -= dc; } }
      return { r: Math.max(0, r), c: Math.max(0, c) };
    }

    /* ---------- 변경 / 실행 취소 ---------- */
    function emit() { listeners.forEach(function (f) { f(); }); }
    function record(e) { history.push(e); if (history.length > 200) history.shift(); future.length = 0; }
    function afterChange() { invalidate(); layout(); render(); emit(); }
    function changeCells(updates) {
      const sh = S();
      const list = updates.map(function (u) { return { si: si, r: u.r, c: u.c, before: clone(getCell(sh, u.r, u.c)), after: clone(u.cell) }; });
      list.forEach(function (ch) { setCellRaw(sh, ch.r, ch.c, clone(ch.after)); });
      record({ type: 'cells', list: list, si: si, sel: Object.assign({}, sel), act: Object.assign({}, act) });
      afterChange();
    }
    function snapshot(fn) {
      const before = clone(book.sheets);
      const beforeSi = si;
      fn();
      record({ type: 'snap', before: before, si: beforeSi });
      afterChange();
      renderTabs();
    }
    function undo() {
      const e = history.pop();
      if (!e) return;
      if (e.type === 'cells') e.list.forEach(function (ch) { setCellRaw(book.sheets[ch.si], ch.r, ch.c, clone(ch.before)); });
      else { const cur = clone(book.sheets); book.sheets.splice.apply(book.sheets, [0, book.sheets.length].concat(e.before)); e.after = cur; }
      future.push(e);
      if (e.si != null && book.sheets[e.si]) si = e.si;
      if (e.sel) { sel = e.sel; act = e.act; }
      afterChange(); renderTabs();
    }
    function redo() {
      const e = future.pop();
      if (!e) return;
      if (e.type === 'cells') e.list.forEach(function (ch) { setCellRaw(book.sheets[ch.si], ch.r, ch.c, clone(ch.after)); });
      else { const cur = clone(book.sheets); book.sheets.splice.apply(book.sheets, [0, book.sheets.length].concat(e.after)); e.before = cur; }
      history.push(e);
      if (e.si != null && book.sheets[e.si]) si = e.si;
      afterChange(); renderTabs();
    }

    function forSel(fn) {
      const out = [];
      const u = used(S());
      // 아주 큰 선택(열 전체 등)은 사용 중인 범위까지만 처리
      const r2 = sel.r2 - sel.r1 < 5000 ? sel.r2 : Math.min(sel.r2, Math.max(u.rows - 1, sel.r1));
      const c2 = sel.c2 - sel.c1 < 200 ? sel.c2 : Math.min(sel.c2, Math.max(u.cols - 1, sel.c1));
      for (let r = sel.r1; r <= r2; r++) for (let c = sel.c1; c <= c2; c++) {
        const cell = fn(clone(getCell(S(), r, c)) || {}, r, c);
        if (cell !== false) out.push({ r: r, c: c, cell: cell });
      }
      if (out.length) changeCells(out);
    }
    function setStyle(key, val, toggle) {
      const cur = getCell(S(), act.r, act.c);
      const on = toggle ? !(cur && cur.s && cur.s[key]) : null;
      forSel(function (cell) {
        cell.s = Object.assign({}, cell.s);
        const v = toggle ? on : val;
        if (v === null || v === false || v === '' || v === undefined) delete cell.s[key]; else cell.s[key] = v;
        if (!Object.keys(cell.s).length) delete cell.s;
        return cell;
      });
    }
    function clearSel(what) {
      forSel(function (cell) {
        if (what !== 'format') { delete cell.v; delete cell.f; delete cell.e; }
        if (what !== 'content') { delete cell.s; delete cell.z; }
        return cell;
      });
    }

    /* ---------- 편집 ---------- */
    function startEdit(mode, initial, source) {
      if (editing) return;
      const cell = getCell(S(), act.r, act.c);
      editing = { r: act.r, c: act.c, mode: mode, source: source || 'cell' };
      placeInput();
      input.classList.add('editing');
      const cs = cell ? cellStyle(cell, null) : '';
      input.style.fontWeight = /font-weight:700/.test(cs) ? '700' : '';
      input.style.fontStyle = /italic/.test(cs) ? 'italic' : '';
      input.style.textAlign = cell && cell.s && cell.s.align && !/^=/.test(rawText(cell)) ? cell.s.align : 'left';
      input.style.background = cell && cell.s && cell.s.bg ? '#' + cell.s.bg : '';
      if (initial != null) input.value = initial;
      else input.value = rawText(cell);
      fbar.value = input.value;
      autoGrow();
      if (source !== 'bar') {
        input.focus();
        const n = input.value.length;
        input.setSelectionRange(n, n);
      }
      updateRefBox();
    }
    function autoGrow() {
      if (!editing) return;
      const r = cellRect(editing.r, editing.c);
      input.style.width = r.width + 'px';
      const need = Math.min(600, input.scrollWidth + 4);
      if (need > r.width) input.style.width = need + 'px';
      input.style.height = r.height + 'px';
      if (input.scrollHeight > r.height) input.style.height = Math.min(300, input.scrollHeight) + 'px';
    }
    function commit(dr, dc) {
      if (!editing) return;
      const text = input.value.replace(/​/g, '');
      const ed = editing;
      editing = null;
      input.classList.remove('editing');
      refBox.hidden = true;
      const cur = getCell(S(), ed.r, ed.c);
      if (text !== rawText(cur)) {
        let nc = DV.sheetInput.parse(text, cur);
        if (nc && nc.f) nc.f = autoClose(nc.f);
        changeCells([{ r: ed.r, c: ed.c, cell: nc }]);
      }
      resetInput();
      if (dr || dc) moveAct(dr, dc);
      else paintSel();
      focusGrid();
    }
    function autoClose(f) {
      let depth = 0, q = false;
      for (const ch of f) { if (ch === '"') q = !q; else if (!q && ch === '(') depth++; else if (!q && ch === ')') depth--; }
      return depth > 0 ? f + ')'.repeat(depth) : f;
    }
    function cancelEdit() {
      if (!editing) return;
      editing = null;
      input.classList.remove('editing');
      refBox.hidden = true;
      resetInput();
      paintSel();
      focusGrid();
    }
    function resetInput() {
      input.value = '​';
      input.style.width = '';
      if (document.activeElement === input) input.select();
    }
    function focusGrid() {
      if (opts.isTouch && opts.isTouch()) return;
      if (document.activeElement !== input) input.focus({ preventScroll: true });
      if (!editing) input.select();
    }
    resetInput();

    // 수식 입력 중 셀을 클릭하면 참조를 넣습니다
    function formulaRefMode() {
      if (!editing) return false;
      const src = editing.source === 'bar' ? fbar : input;
      const v = src.value;
      if (v[0] !== '=') return false;
      const pos = src.selectionStart == null ? v.length : src.selectionStart;
      const before = v.slice(0, pos).replace(/\s+$/, '');
      return /[=(,+\-*/&^<>:;]$/.test(before) || (refInsert && refInsert.end === pos);
    }
    let refInsert = null;
    function insertRef(r1, c1, r2, c2) {
      const src = editing.source === 'bar' ? fbar : input;
      const v = src.value;
      let start, end;
      if (refInsert) { start = refInsert.start; end = refInsert.end; }
      else { start = end = src.selectionStart == null ? v.length : src.selectionStart; }
      const t = (r1 === r2 && c1 === c2) ? F().addr(r1, c1) : F().addr(Math.min(r1, r2), Math.min(c1, c2)) + ':' + F().addr(Math.max(r1, r2), Math.max(c1, c2));
      src.value = v.slice(0, start) + t + v.slice(end);
      refInsert = { start: start, end: start + t.length };
      src.setSelectionRange(refInsert.end, refInsert.end);
      if (src === fbar) input.value = fbar.value; else fbar.value = input.value;
      autoGrow();
      updateRefBox();
    }
    function updateRefBox() {
      if (!editing) { refBox.hidden = true; return; }
      const v = input.value;
      const m = /([A-Za-z]{1,3}\$?\d+)(?::\$?([A-Za-z]{1,3}\$?\d+))?(?![\w(])/g;
      let last = null, mm;
      if (v[0] === '=') while ((mm = m.exec(v))) last = mm;
      if (!last) { refBox.hidden = true; return; }
      const a = F().parseAddr(last[1].replace(/\$/g, '')), b = last[2] ? F().parseAddr(last[2].replace(/\$/g, '')) : a;
      if (!a || !b) { refBox.hidden = true; return; }
      const r = cellRect(a.r, a.c, b.r, b.c);
      refBox.style.cssText = 'left:' + r.left + 'px;top:' + r.top + 'px;width:' + r.width + 'px;height:' + r.height + 'px';
      refBox.hidden = false;
    }

    input.addEventListener('compositionstart', function () {
      if (!editing) { startEdit('enter', '', 'cell'); }
    });
    input.addEventListener('input', function () {
      if (!editing) {
        const v = input.value.replace(/​/g, '');
        if (!v) { resetInput(); return; }
        startEdit('enter', v, 'cell');
        return;
      }
      refInsert = null;
      fbar.value = input.value;
      autoGrow();
      updateRefBox();
    });
    input.addEventListener('keydown', function (e) {
      const mod = e.ctrlKey || e.metaKey;
      if (e.isComposing || e.keyCode === 229) return;
      if (editing) {
        if (e.key === 'Enter' && e.altKey) {
          e.preventDefault();
          const s = input.selectionStart, en = input.selectionEnd;
          input.value = input.value.slice(0, s) + '\n' + input.value.slice(en);
          input.setSelectionRange(s + 1, s + 1);
          fbar.value = input.value; autoGrow();
          return;
        }
        if (e.key === 'Enter') { e.preventDefault(); commit(e.shiftKey ? -1 : 1, 0); return; }
        if (e.key === 'Tab') { e.preventDefault(); commit(0, e.shiftKey ? -1 : 1); return; }
        if (e.key === 'Escape') { e.preventDefault(); cancelEdit(); return; }
        if (editing.mode === 'enter' && /^Arrow/.test(e.key)) {
          if (formulaRefMode()) {
            e.preventDefault();
            const base = refInsert ? refInsert.cur : { r: editing.r, c: editing.c };
            const d = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] }[e.key];
            const nr = Math.max(0, base.r + d[0]), nc = Math.max(0, base.c + d[1]);
            insertRef(nr, nc, nr, nc);
            refInsert.cur = { r: nr, c: nc };
            return;
          }
          e.preventDefault();
          const d = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] }[e.key];
          commit(d[0], d[1]);
          return;
        }
        return;
      }
      // 편집 중이 아닐 때
      const k = e.key;
      if (/^Arrow/.test(k)) {
        e.preventDefault();
        const d = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] }[k];
        if (mod) {
          const t = jumpEdge(d[0], d[1]);
          if (e.shiftKey) { sel = { r1: Math.min(act.r, t.r), c1: Math.min(act.c, t.c), r2: Math.max(act.r, t.r), c2: Math.max(act.c, t.c) }; paintSel(); scrollIntoView(t.r, t.c); }
          else { setSel(t.r, t.c, t.r, t.c); scrollIntoView(t.r, t.c); }
        } else moveAct(d[0], d[1], e.shiftKey);
        return;
      }
      if (k === 'Enter') { e.preventDefault(); moveWithin(e.shiftKey ? -1 : 1, 0); return; }
      if (k === 'Tab') { e.preventDefault(); moveWithin(0, e.shiftKey ? -1 : 1); return; }
      if (k === 'F2') { e.preventDefault(); startEdit('edit', null, 'cell'); return; }
      if (k === 'Delete' || k === 'Backspace') { e.preventDefault(); clearSel('content'); return; }
      if (k === 'Home') { e.preventDefault(); const r = mod ? 0 : act.r; setSel(r, 0, r, 0); scrollIntoView(r, 0); return; }
      if (k === 'End' && mod) { e.preventDefault(); const u = used(S()); setSel(Math.max(0, u.rows - 1), Math.max(0, u.cols - 1), Math.max(0, u.rows - 1), Math.max(0, u.cols - 1)); scrollIntoView(act.r, act.c); return; }
      if (k === 'PageDown' || k === 'PageUp') {
        e.preventDefault();
        const n = Math.max(1, Math.floor((wrap.clientHeight - HDR_H) / DV.xlsxIO.DEFAULT_ROW_PX) - 1);
        moveAct(k === 'PageDown' ? n : -n, 0, e.shiftKey);
        return;
      }
      if (mod) {
        const kk = k.toLowerCase();
        if (kk === 'z') { e.preventDefault(); if (e.shiftKey) redo(); else undo(); return; }
        if (kk === 'y') { e.preventDefault(); redo(); return; }
        if (kk === 'b') { e.preventDefault(); setStyle('b', null, true); return; }
        if (kk === 'i') { e.preventDefault(); setStyle('i', null, true); return; }
        if (kk === 'u') { e.preventDefault(); setStyle('u', null, true); return; }
        if (kk === 'a') { e.preventDefault(); act = { r: 0, c: 0 }; sel = { r1: 0, c1: 0, r2: rowCount - 1, c2: colCount - 1 }; paintSel(); return; }
        if (kk === 'd') { e.preventDefault(); fill('down'); return; }
        if (kk === 'r') { e.preventDefault(); fill('right'); return; }
      }
    });
    function moveWithin(dr, dc) {
      if (sel.r1 === sel.r2 && sel.c1 === sel.c2 || isSingleMerged()) { moveAct(dr, dc); return; }
      let r = act.r + dr, c = act.c + dc;
      if (r > sel.r2) { r = sel.r1; c = c + 1 > sel.c2 ? sel.c1 : c + 1; }
      if (r < sel.r1) { r = sel.r2; c = c - 1 < sel.c1 ? sel.c2 : c - 1; }
      if (c > sel.c2) { c = sel.c1; r = r + 1 > sel.r2 ? sel.r1 : r + 1; }
      if (c < sel.c1) { c = sel.c2; r = r - 1 < sel.r1 ? sel.r2 : r - 1; }
      act = { r: r, c: c };
      paintSel();
      scrollIntoView(r, c);
    }
    function fill(dir) {
      const out = [];
      if (dir === 'down') {
        const r0 = sel.r1 === sel.r2 ? sel.r1 - 1 : sel.r1;
        if (r0 < 0) return;
        for (let c = sel.c1; c <= sel.c2; c++) {
          const src = getCell(S(), r0, c);
          for (let r = r0 + 1; r <= sel.r2; r++) out.push({ r: r, c: c, cell: shiftCell(src, r - r0, 0) });
        }
      } else {
        const c0 = sel.c1 === sel.c2 ? sel.c1 - 1 : sel.c1;
        if (c0 < 0) return;
        for (let r = sel.r1; r <= sel.r2; r++) {
          const src = getCell(S(), r, c0);
          for (let c = c0 + 1; c <= sel.c2; c++) out.push({ r: r, c: c, cell: shiftCell(src, 0, c - c0) });
        }
      }
      if (out.length) changeCells(out);
    }
    function shiftCell(cell, dr, dc) {
      const n = clone(cell);
      if (n && n.f) n.f = F().shift(n.f, dr, dc);
      return n;
    }

    fbar.addEventListener('focus', function () {
      if (!editing) startEdit('edit', null, 'bar');
      else editing.source = 'bar';
    });
    fbar.addEventListener('input', function () {
      if (!editing) startEdit('edit', null, 'bar');
      input.value = fbar.value;
      refInsert = null;
      autoGrow();
      updateRefBox();
    });
    fbar.addEventListener('keydown', function (e) {
      if (e.isComposing || e.keyCode === 229) return;
      if (e.key === 'Enter') { e.preventDefault(); commit(1, 0); }
      else if (e.key === 'Escape') { e.preventDefault(); cancelEdit(); }
      else if (e.key === 'Tab') { e.preventDefault(); commit(0, 1); }
    });
    namebox.addEventListener('keydown', function (e) {
      if (e.key !== 'Enter') return;
      e.preventDefault();
      const v = namebox.value.trim().toUpperCase().replace(/\$/g, '');
      const parts = v.split(':');
      const a = F().parseAddr(parts[0]), b = parts[1] ? F().parseAddr(parts[1]) : a;
      if (a && b) { setSel(a.r, a.c, b.r, b.c); scrollIntoView(a.r, a.c); }
      focusGrid();
    });

    /* ---------- 마우스 ---------- */
    function cellAt(e) {
      const rect = gbody.getBoundingClientRect();
      const x = e.clientX - rect.left - ROWHDR_W, y = e.clientY - rect.top;
      return { r: Math.max(0, Math.min(rowCount - 1, bsearch(rowTops, Math.max(0, y)))), c: Math.max(0, Math.min(colCount - 1, bsearch(colLefts, Math.max(0, x)))) };
    }
    let drag = null;
    let lastTap = 0;
    gbody.addEventListener('pointerdown', function (e) {
      if (e.target === input || e.button === 2 && inSel(cellAt(e))) return;
      if (e.target.closest('.rowhdrs')) return;
      const p = cellAt(e);
      if (editing && formulaRefMode() && !(p.r === editing.r && p.c === editing.c)) {
        e.preventDefault();
        refInsert = null;
        insertRef(p.r, p.c, p.r, p.c);
        drag = { mode: 'ref', r: p.r, c: p.c };
        return;
      }
      if (editing) commit(0, 0);
      const touch = e.pointerType === 'touch';
      if (touch) {
        const now = Date.now();
        const same = p.r === act.r && p.c === act.c;
        if (same && now - lastTap < 400) { startEdit('edit', null, 'cell'); lastTap = 0; return; }
        lastTap = now;
        setSel(p.r, p.c, p.r, p.c);
        return;
      }
      e.preventDefault();
      if (e.shiftKey) { sel = expandMerges({ r1: Math.min(act.r, p.r), c1: Math.min(act.c, p.c), r2: Math.max(act.r, p.r), c2: Math.max(act.c, p.c) }); paintSel(); }
      else { const m = mergeMap.get(p.r * 20000 + p.c); const q = m ? { r: m.r1, c: m.c1 } : p; setSel(q.r, q.c, q.r, q.c); }
      drag = { mode: 'sel' };
      focusGrid();
    });
    gbody.addEventListener('dblclick', function (e) {
      if (e.target === input) return;
      startEdit('edit', null, 'cell');
    });
    function inSel(p) { return p.r >= sel.r1 && p.r <= sel.r2 && p.c >= sel.c1 && p.c <= sel.c2; }
    window.addEventListener('pointermove', function (e) {
      if (!drag) return;
      if (drag.mode === 'sel' || drag.mode === 'ref') {
        const p = cellAt(e);
        if (drag.mode === 'ref') { insertRef(drag.r, drag.c, p.r, p.c); return; }
        sel = expandMerges({ r1: Math.min(act.r, p.r), c1: Math.min(act.c, p.c), r2: Math.max(act.r, p.r), c2: Math.max(act.c, p.c) });
        paintSel();
        autoScroll(e);
      } else if (drag.mode === 'col' || drag.mode === 'row') {
        const p = cellAt(e);
        if (drag.mode === 'col') { sel = { r1: 0, r2: rowCount - 1, c1: Math.min(drag.start, p.c), c2: Math.max(drag.start, p.c) }; }
        else { sel = { c1: 0, c2: colCount - 1, r1: Math.min(drag.start, p.r), r2: Math.max(drag.start, p.r) }; }
        paintSel();
      } else if (drag.mode === 'rzc') {
        const w = Math.max(8, drag.w0 + e.clientX - drag.x0);
        S().colWidths[drag.i] = Math.round(w);
        layout(); schedule();
      } else if (drag.mode === 'rzr') {
        const h = Math.max(8, drag.h0 + e.clientY - drag.y0);
        S().rowHeights[drag.i] = Math.round(h);
        layout(); schedule();
      }
    });
    window.addEventListener('pointerup', function () {
      if (!drag) return;
      if (drag.mode === 'rzc' || drag.mode === 'rzr') { emit(); }
      if (drag.mode === 'ref') { input.focus(); }
      drag = null;
    });
    function autoScroll(e) {
      const r = wrap.getBoundingClientRect();
      if (e.clientY > r.bottom - 20) wrap.scrollTop += 20;
      else if (e.clientY < r.top + HDR_H + 10) wrap.scrollTop -= 20;
      if (e.clientX > r.right - 20) wrap.scrollLeft += 20;
      else if (e.clientX < r.left + ROWHDR_W + 10) wrap.scrollLeft -= 20;
    }
    colHdrRow.addEventListener('pointerdown', function (e) {
      if (editing) commit(0, 0);
      const rz = e.target.closest('[data-rzc]');
      if (rz) { e.preventDefault(); const i = +rz.dataset.rzc; drag = { mode: 'rzc', i: i, x0: e.clientX, w0: colW(i) }; return; }
      if (e.target.closest('.corner')) { act = { r: 0, c: 0 }; sel = { r1: 0, c1: 0, r2: rowCount - 1, c2: colCount - 1 }; paintSel(); focusGrid(); return; }
      const h = e.target.closest('.ch');
      if (!h) return;
      if (e.button === 2 && +h.dataset.c >= sel.c1 && +h.dataset.c <= sel.c2) return;
      e.preventDefault();
      const c = +h.dataset.c;
      const start = e.shiftKey ? act.c : c;
      if (!e.shiftKey) act = { r: Math.max(0, bsearch(rowTops, wrap.scrollTop)), c: c };
      sel = { r1: 0, r2: rowCount - 1, c1: Math.min(start, c), c2: Math.max(start, c) };
      paintSel();
      drag = { mode: 'col', start: start };
      focusGrid();
    });
    rowHdrs.addEventListener('pointerdown', function (e) {
      if (editing) commit(0, 0);
      const rz = e.target.closest('[data-rzr]');
      if (rz) { e.preventDefault(); const i = +rz.dataset.rzr; drag = { mode: 'rzr', i: i, y0: e.clientY, h0: rowH(i) }; return; }
      const h = e.target.closest('.rh');
      if (!h) return;
      if (e.button === 2 && +h.dataset.r >= sel.r1 && +h.dataset.r <= sel.r2) return;
      e.preventDefault();
      const r = +h.dataset.r;
      const start = e.shiftKey ? act.r : r;
      if (!e.shiftKey) act = { r: r, c: Math.max(0, bsearch(colLefts, wrap.scrollLeft)) };
      sel = { c1: 0, c2: colCount - 1, r1: Math.min(start, r), r2: Math.max(start, r) };
      paintSel();
      drag = { mode: 'row', start: start };
      focusGrid();
    });
    colHdrRow.addEventListener('dblclick', function (e) {
      const rz = e.target.closest('[data-rzc]');
      if (rz) autoFitCol(+rz.dataset.rzc);
    });
    function autoFitCol(c) {
      const sh = S();
      const cv = document.createElement('canvas').getContext('2d');
      cv.font = '13px ' + getComputedStyle(cellsLayer).fontFamily;
      let w = 40;
      sh.rows.forEach(function (row, r) {
        const cell = row && row[c];
        if (!cell || mergeMap.has(r * 20000 + c)) return;
        const t = display(cell, value(si, r, c));
        const scale = cell.s && cell.s.size ? cell.s.size / 11 : 1;
        w = Math.max(w, cv.measureText(t).width * scale * (cell.s && cell.s.b ? 1.08 : 1) + 12);
      });
      sh.colWidths[c] = Math.min(600, Math.ceil(w));
      layout(); render(); emit();
    }

    /* ---------- 복사 / 붙여넣기 ---------- */
    function selText() {
      const u = used(S());
      const r2 = Math.min(sel.r2, Math.max(u.rows - 1, sel.r1)), c2 = Math.min(sel.c2, Math.max(u.cols - 1, sel.c1));
      const lines = [];
      let html = '<table>';
      for (let r = sel.r1; r <= r2; r++) {
        const line = [];
        html += '<tr>';
        for (let c = sel.c1; c <= c2; c++) {
          const cell = getCell(S(), r, c);
          const t = cell ? display(cell, value(si, r, c)) : '';
          line.push(/[\t\n"]/.test(t) ? '"' + t.replace(/"/g, '""') + '"' : t);
          html += '<td>' + DV.esc(t) + '</td>';
        }
        lines.push(line.join('\t'));
        html += '</tr>';
      }
      return { text: lines.join('\n'), html: html + '</table>', r2: r2, c2: c2 };
    }
    function doCopy(e, cut) {
      const t = selText();
      const cells = [];
      for (let r = sel.r1; r <= t.r2; r++) { const row = []; for (let c = sel.c1; c <= t.c2; c++) row.push(clone(getCell(S(), r, c))); cells.push(row); }
      clip = { text: t.text, cells: cells, si: si, r1: sel.r1, c1: sel.c1, r2: t.r2, c2: t.c2, cut: !!cut };
      if (e && e.clipboardData) {
        e.clipboardData.setData('text/plain', t.text);
        e.clipboardData.setData('text/html', t.html);
        e.preventDefault();
      } else if (navigator.clipboard) navigator.clipboard.writeText(t.text).catch(function () { /* 무시 */ });
      DV.toast(cut ? '잘라냈습니다. 붙여넣을 위치를 선택하세요.' : '복사했습니다.', 'info', 1500);
    }
    function parseTSV(text) {
      const rows = [];
      let row = [], cell = '', q = false;
      text = text.replace(/\r\n?/g, '\n').replace(/\n$/, '');
      for (let i = 0; i < text.length; i++) {
        const ch = text[i];
        if (q) { if (ch === '"') { if (text[i + 1] === '"') { cell += '"'; i++; } else q = false; } else cell += ch; }
        else if (ch === '"' && cell === '') q = true;
        else if (ch === '\t') { row.push(cell); cell = ''; }
        else if (ch === '\n') { row.push(cell); rows.push(row); row = []; cell = ''; }
        else cell += ch;
      }
      row.push(cell); rows.push(row);
      return rows;
    }
    function doPaste(text) {
      const out = [];
      if (clip && text != null && text.replace(/\r/g, '') === clip.text.replace(/\r/g, '')) {
        const dr = act.r - clip.r1, dc = act.c - clip.c1;
        const h = clip.cells.length, w = clip.cells[0] ? clip.cells[0].length : 0;
        // 선택 영역이 원본의 배수면 반복해서 채웁니다
        const reps = { r: Math.max(1, Math.floor((sel.r2 - sel.r1 + 1) / h)), c: Math.max(1, Math.floor((sel.c2 - sel.c1 + 1) / w)) };
        for (let rr = 0; rr < reps.r; rr++) for (let cc = 0; cc < reps.c; cc++) {
          clip.cells.forEach(function (row, i) {
            row.forEach(function (cell, j) {
              const r = act.r + rr * h + i, c = act.c + cc * w + j;
              out.push({ r: r, c: c, cell: clip.cut ? clone(cell) : shiftCell(cell, dr + rr * h, dc + cc * w) });
            });
          });
        }
        if (clip.cut && clip.si === si) {
          for (let r = clip.r1; r <= clip.r2; r++) for (let c = clip.c1; c <= clip.c2; c++) {
            const inDest = r >= act.r && r < act.r + h && c >= act.c && c < act.c + w;
            if (!inDest) out.unshift({ r: r, c: c, cell: undefined });
          }
          clip = null;
        }
      } else if (text != null && text !== '') {
        const rows = parseTSV(text);
        rows.forEach(function (row, i) {
          row.forEach(function (t, j) {
            const base = getCell(S(), act.r + i, act.c + j);
            out.push({ r: act.r + i, c: act.c + j, cell: DV.sheetInput.parse(t, base && { s: base.s, z: base.z }) });
          });
        });
      }
      if (!out.length) return;
      changeCells(out);
      let maxR = act.r, maxC = act.c;
      out.forEach(function (o) { if (o.r >= act.r && o.c >= act.c) { maxR = Math.max(maxR, o.r); maxC = Math.max(maxC, o.c); } });
      sel = { r1: act.r, c1: act.c, r2: maxR, c2: maxC };
      paintSel();
    }
    input.addEventListener('copy', function (e) { if (!editing) doCopy(e, false); });
    input.addEventListener('cut', function (e) { if (!editing) doCopy(e, true); });
    input.addEventListener('paste', function (e) {
      if (editing) return;
      e.preventDefault();
      doPaste(e.clipboardData.getData('text/plain'));
    });

    /* ---------- 구조 변경 ---------- */
    function adjustAll(targetName, axis, at, count) {
      book.sheets.forEach(function (sh) {
        sh.rows.forEach(function (row) {
          if (!row) return;
          row.forEach(function (cell) { if (cell && cell.f) cell.f = F().adjust(cell.f, sh.name, targetName, axis, at, count); });
        });
      });
    }
    function insertRows(at, n) {
      snapshot(function () {
        const sh = S();
        adjustAll(sh.name, 'row', at, n);
        if (sh.rows.length > at) sh.rows.splice.apply(sh.rows, [at, 0].concat(new Array(n)));
        if (sh.rowHeights.length > at) sh.rowHeights.splice.apply(sh.rowHeights, [at, 0].concat(new Array(n)));
        sh.merges.forEach(function (m) { if (m.r1 >= at) { m.r1 += n; m.r2 += n; } else if (m.r2 >= at) m.r2 += n; });
      });
    }
    function deleteRows(at, n) {
      snapshot(function () {
        const sh = S();
        adjustAll(sh.name, 'row', at, -n);
        sh.rows.splice(at, n);
        sh.rowHeights.splice(at, n);
        sh.merges = sh.merges.filter(function (m) { return !(m.r1 >= at && m.r2 < at + n); });
        sh.merges.forEach(function (m) {
          if (m.r1 >= at + n) { m.r1 -= n; m.r2 -= n; }
          else if (m.r2 >= at) { m.r2 = Math.max(m.r1 >= at ? at : m.r1, m.r2 - n); if (m.r1 > at) m.r1 = at; }
        });
      });
    }
    function insertCols(at, n) {
      snapshot(function () {
        const sh = S();
        adjustAll(sh.name, 'col', at, n);
        sh.rows.forEach(function (row) { if (row && row.length > at) row.splice.apply(row, [at, 0].concat(new Array(n))); });
        if (sh.colWidths.length > at) sh.colWidths.splice.apply(sh.colWidths, [at, 0].concat(new Array(n)));
        sh.merges.forEach(function (m) { if (m.c1 >= at) { m.c1 += n; m.c2 += n; } else if (m.c2 >= at) m.c2 += n; });
      });
    }
    function deleteCols(at, n) {
      snapshot(function () {
        const sh = S();
        adjustAll(sh.name, 'col', at, -n);
        sh.rows.forEach(function (row) { if (row) row.splice(at, n); });
        sh.colWidths.splice(at, n);
        sh.merges = sh.merges.filter(function (m) { return !(m.c1 >= at && m.c2 < at + n); });
        sh.merges.forEach(function (m) {
          if (m.c1 >= at + n) { m.c1 -= n; m.c2 -= n; }
          else if (m.c2 >= at) { m.c2 = Math.max(m.c1 >= at ? at : m.c1, m.c2 - n); if (m.c1 > at) m.c1 = at; }
        });
      });
    }
    function wholeRows() { return sel.c1 === 0 && sel.c2 >= colCount - 1; }
    function wholeCols() { return sel.r1 === 0 && sel.r2 >= rowCount - 1; }
    function toggleMerge() {
      const existing = S().merges.filter(function (m) { return !(m.r2 < sel.r1 || m.r1 > sel.r2 || m.c2 < sel.c1 || m.c1 > sel.c2); });
      if (existing.length) {
        snapshot(function () { S().merges = S().merges.filter(function (m) { return existing.indexOf(m) < 0; }); });
        return;
      }
      if (sel.r1 === sel.r2 && sel.c1 === sel.c2) return;
      if ((sel.r2 - sel.r1 + 1) * (sel.c2 - sel.c1 + 1) > 100000) { DV.toast('병합할 범위가 너무 큽니다.', 'warn'); return; }
      let filled = 0;
      for (let r = sel.r1; r <= sel.r2; r++) for (let c = sel.c1; c <= sel.c2; c++) { const x = getCell(S(), r, c); if (x && (x.v != null && x.v !== '' || x.f) && !(r === sel.r1 && c === sel.c1)) filled++; }
      if (filled && !confirm('병합하면 왼쪽 위 셀의 값만 남습니다. 계속할까요?')) return;
      snapshot(function () {
        const sh = S();
        for (let r = sel.r1; r <= sel.r2; r++) for (let c = sel.c1; c <= sel.c2; c++) if (!(r === sel.r1 && c === sel.c1)) setCellRaw(sh, r, c, undefined);
        sh.merges.push({ r1: sel.r1, c1: sel.c1, r2: sel.r2, c2: sel.c2 });
        const tl = getCell(sh, sel.r1, sel.c1) || {};
        tl.s = Object.assign({}, tl.s, { align: 'center', valign: 'middle' });
        setCellRaw(sh, sel.r1, sel.c1, tl);
      });
      act = { r: sel.r1, c: sel.c1 };
    }
    function sortSel(desc) {
      const u = used(S());
      let r1 = sel.r1, r2 = Math.min(sel.r2, u.rows - 1), c1 = sel.c1, c2 = Math.min(sel.c2, Math.max(u.cols - 1, sel.c1));
      if (r1 === r2 || r2 < r1) {
        // 한 셀만 선택: 현재 열이 포함된 데이터 영역 전체(첫 행은 머리글로 간주)
        r1 = 1; r2 = u.rows - 1; c1 = 0; c2 = Math.max(0, u.cols - 1);
        if (r2 <= r1) { DV.toast('정렬할 데이터가 없습니다.', 'warn'); return; }
      }
      if (S().merges.some(function (m) { return !(m.r2 < r1 || m.r1 > r2 || m.c2 < c1 || m.c1 > c2); })) { DV.toast('병합된 셀이 있는 범위는 정렬할 수 없습니다.', 'warn'); return; }
      const key = act.c;
      snapshot(function () {
        const sh = S();
        const rows = [];
        for (let r = r1; r <= r2; r++) {
          const cells = [];
          for (let c = c1; c <= c2; c++) cells.push(getCell(sh, r, c));
          rows.push({ cells: cells, v: value(si, r, key), r: r });
        }
        rows.sort(function (a, b) {
          const x = a.v, y = b.v;
          if (x == null && y == null) return a.r - b.r;
          if (x == null) return 1;
          if (y == null) return -1;
          const tx = typeof x === 'number' ? 0 : 1, ty = typeof y === 'number' ? 0 : 1;
          if (tx !== ty) return tx - ty;
          let cmp = tx === 0 ? x - y : String(x).localeCompare(String(y), 'ko');
          if (desc) cmp = -cmp;
          return cmp || a.r - b.r;
        });
        rows.forEach(function (row, i) {
          const r = r1 + i;
          row.cells.forEach(function (cell, j) {
            let n = cell;
            if (n && n.f && row.r !== r) { n = clone(n); n.f = F().shift(n.f, r - row.r, 0); }
            setCellRaw(sh, r, c1 + j, n);
          });
        });
      });
    }
    function autoSum() {
      // 선택 영역 아래(또는 위쪽 연속 숫자)에 SUM 넣기
      if (sel.r1 !== sel.r2 || sel.c1 !== sel.c2) {
        const out = [];
        for (let c = sel.c1; c <= sel.c2; c++) out.push({ r: sel.r2 + 1, c: c, cell: Object.assign({}, getCell(S(), sel.r2 + 1, c), { f: 'SUM(' + F().addr(sel.r1, c) + ':' + F().addr(sel.r2, c) + ')' }) });
        changeCells(out);
        return;
      }
      let r = act.r - 1;
      while (r >= 0 && typeof value(si, r, act.c) === 'number') r--;
      if (r === act.r - 1) {
        let c = act.c - 1;
        while (c >= 0 && typeof value(si, act.r, c) === 'number') c--;
        if (c === act.c - 1) { startEdit('enter', '=SUM()', 'cell'); input.setSelectionRange(5, 5); return; }
        startEdit('enter', '=SUM(' + F().addr(act.r, c + 1) + ':' + F().addr(act.r, act.c - 1) + ')', 'cell');
        return;
      }
      startEdit('enter', '=SUM(' + F().addr(r + 1, act.c) + ':' + F().addr(act.r - 1, act.c) + ')', 'cell');
    }
    function changeDecimals(delta) {
      const v = value(si, act.r, act.c);
      if (typeof v !== 'number') return;
      const cell = getCell(S(), act.r, act.c);
      let z = cell && cell.z || '';
      let base, dec;
      if (!z || z === 'General') {
        const s = fmtGeneral(v);
        dec = s.indexOf('.') >= 0 ? s.length - s.indexOf('.') - 1 : 0;
        base = '0';
      } else {
        const m = /(0+)(?:\.(0+))?/.exec(z);
        dec = m && m[2] ? m[2].length : 0;
        base = z;
      }
      dec = Math.max(0, Math.min(10, dec + delta));
      let nz;
      if (base === '0') nz = dec ? '0.' + '0'.repeat(dec) : '0';
      else nz = base.replace(/(0+)(?:\.0+)?/, function (m, a) { return a + (dec ? '.' + '0'.repeat(dec) : ''); });
      forSel(function (c) { if (typeof c.v === 'number' || c.f) { c.z = nz; return c; } return false; });
    }

    /* ---------- 도구 모음 ---------- */
    function showPalette(anchor, kind) {
      palette.innerHTML = '<div class="pal-grid">' + COLORS.map(function (c) { return '<button type="button" style="background:#' + c + '" data-color="' + c + '"></button>'; }).join('') + '</div>' +
        '<button type="button" class="pal-none" data-color="">' + (kind === 'bg' ? '채우기 없음' : '자동(검정)') + '</button>';
      palette.dataset.kind = kind;
      const r = anchor.getBoundingClientRect(), er = el.getBoundingClientRect();
      palette.style.left = Math.min(r.left - er.left, er.width - 220) + 'px';
      palette.style.top = (r.bottom - er.top + 4) + 'px';
      palette.hidden = false;
    }
    palette.addEventListener('click', function (e) {
      const b = e.target.closest('[data-color]');
      if (!b) return;
      const c = b.dataset.color || null;
      if (palette.dataset.kind === 'bg') { if (c) { backColor = c; el.querySelector('.bar.back').style.background = '#' + c; } setStyle('bg', c); }
      else { if (c) { foreColor = c; el.querySelector('.bar.fore').style.background = '#' + c; } setStyle('color', c); }
      palette.hidden = true;
      focusGrid();
    });
    toolbar.addEventListener('mousedown', function (e) { if (e.target.closest('button')) e.preventDefault(); });
    toolbar.addEventListener('click', function (e) {
      const b = e.target.closest('button[data-cmd]');
      if (!b) return;
      if (editing) commit(0, 0);
      palette.hidden = true;
      const cmd = b.dataset.cmd;
      switch (cmd) {
        case 'undo': undo(); break;
        case 'redo': redo(); break;
        case 'b': case 'i': case 'u': case 'strike': case 'wrap': setStyle(cmd, null, true); break;
        case 'left': case 'center': case 'right': {
          const cur = getCell(S(), act.r, act.c);
          setStyle('align', cur && cur.s && cur.s.align === cmd ? null : cmd);
          break;
        }
        case 'color': setStyle('color', foreColor); break;
        case 'bg': setStyle('bg', backColor); break;
        case 'pal-color': showPalette(b, 'color'); return;
        case 'pal-bg': showPalette(b, 'bg'); return;
        case 'border': {
          const cur = getCell(S(), act.r, act.c);
          const on = !(cur && cur.s && cur.s.border);
          setStyle('border', on ? { left: 1, right: 1, top: 1, bottom: 1 } : null);
          break;
        }
        case 'merge': toggleMerge(); break;
        case 'sum': autoSum(); return;
        case 'sortasc': sortSel(false); break;
        case 'sortdesc': sortSel(true); break;
        case 'insrow': insertRows(sel.r1, wholeRows() ? sel.r2 - sel.r1 + 1 : 1); break;
        case 'delrow': deleteRows(sel.r1, wholeRows() ? sel.r2 - sel.r1 + 1 : 1); break;
        case 'inscol': insertCols(sel.c1, wholeCols() ? sel.c2 - sel.c1 + 1 : 1); break;
        case 'delcol': deleteCols(sel.c1, wholeCols() ? sel.c2 - sel.c1 + 1 : 1); break;
        case 'dec+': changeDecimals(1); break;
        case 'dec-': changeDecimals(-1); break;
      }
      focusGrid();
    });
    toolbar.addEventListener('change', function (e) {
      const s = e.target.closest('select');
      if (!s) return;
      if (editing) commit(0, 0);
      if (s.dataset.cmd === 'fmt') {
        const z = s.value;
        forSel(function (c) { if (z) c.z = z; else delete c.z; return c; });
      } else if (s.dataset.cmd === 'size') {
        const v = parseFloat(s.value);
        setStyle('size', v && v !== 11 ? v : null);
      }
      focusGrid();
    });

    /* ---------- 오른쪽 클릭 메뉴 ---------- */
    wrap.addEventListener('contextmenu', function (e) {
      e.preventDefault();
      if (editing) commit(0, 0);
      const onHdr = e.target.closest('.ch') ? 'col' : e.target.closest('.rh') ? 'row' : null;
      if (!onHdr) {
        const p = cellAt(e);
        if (!inSel(p)) setSel(p.r, p.c, p.r, p.c);
      }
      const merged = S().merges.some(function (m) { return !(m.r2 < sel.r1 || m.r1 > sel.r2 || m.c2 < sel.c1 || m.c1 > sel.c2); });
      const items = [['cut', '잘라내기'], ['copy', '복사'], ['paste', '붙여넣기'], ['-'],
        ['rowAbove', '위에 행 삽입'], ['rowBelow', '아래에 행 삽입'], ['colLeft', '왼쪽에 열 삽입'], ['colRight', '오른쪽에 열 삽입'], ['-'],
        ['delRow', '행 삭제'], ['delCol', '열 삭제'], ['clear', '내용 지우기'], ['clearFmt', '서식 지우기'], ['-'],
        ['merge', merged ? '셀 병합 해제' : '셀 병합'], ['sortAsc', '오름차순 정렬'], ['sortDesc', '내림차순 정렬']];
      if (onHdr === 'col') items.push(['-'], ['autofit', '열 너비 자동 맞춤']);
      ctxmenu.innerHTML = items.map(function (it) { return it[0] === '-' ? '<hr>' : '<button type="button" data-op="' + it[0] + '">' + it[1] + '</button>'; }).join('');
      const er = el.getBoundingClientRect();
      ctxmenu.hidden = false;
      const mh = ctxmenu.offsetHeight;
      ctxmenu.style.left = Math.min(e.clientX - er.left, er.width - 200) + 'px';
      ctxmenu.style.top = Math.max(0, Math.min(e.clientY - er.top, er.height - mh - 8)) + 'px';
    });
    ctxmenu.addEventListener('click', async function (e) {
      const b = e.target.closest('[data-op]');
      if (!b) return;
      ctxmenu.hidden = true;
      const n = function (a, b2) { return b2 - a + 1; };
      switch (b.dataset.op) {
        case 'cut': doCopy(null, true); break;
        case 'copy': doCopy(null, false); break;
        case 'paste': {
          let text = clip ? clip.text : '';
          try { if (navigator.clipboard && navigator.clipboard.readText) text = await navigator.clipboard.readText(); } catch (err) { /* 권한 없음 */ }
          doPaste(text);
          break;
        }
        case 'rowAbove': insertRows(sel.r1, wholeRows() ? n(sel.r1, sel.r2) : 1); break;
        case 'rowBelow': insertRows(sel.r2 + 1, wholeRows() ? n(sel.r1, sel.r2) : 1); break;
        case 'colLeft': insertCols(sel.c1, wholeCols() ? n(sel.c1, sel.c2) : 1); break;
        case 'colRight': insertCols(sel.c2 + 1, wholeCols() ? n(sel.c1, sel.c2) : 1); break;
        case 'delRow': deleteRows(sel.r1, Math.min(n(sel.r1, sel.r2), Math.max(1, used(S()).rows - sel.r1))); break;
        case 'delCol': deleteCols(sel.c1, Math.min(n(sel.c1, sel.c2), Math.max(1, used(S()).cols - sel.c1))); break;
        case 'clear': clearSel('content'); break;
        case 'clearFmt': clearSel('format'); break;
        case 'merge': toggleMerge(); break;
        case 'sortAsc': sortSel(false); break;
        case 'sortDesc': sortSel(true); break;
        case 'autofit': for (let c = sel.c1; c <= Math.min(sel.c2, sel.c1 + 50); c++) autoFitCol(c); break;
      }
      focusGrid();
    });
    document.addEventListener('mousedown', function (e) {
      if (!ctxmenu.hidden && !ctxmenu.contains(e.target)) ctxmenu.hidden = true;
      if (!palette.hidden && !palette.contains(e.target) && !e.target.closest('.color-btn')) palette.hidden = true;
    });

    /* ---------- 시트 탭 ---------- */
    function renderTabs() {
      tabList.innerHTML = book.sheets.map(function (sh, i) {
        return '<button type="button" class="stab' + (i === si ? ' on' : '') + '" data-i="' + i + '" title="두 번 눌러 이름 바꾸기">' + DV.esc(sh.name) + '</button>';
      }).join('');
    }
    function switchSheet(i) {
      if (editing) commit(0, 0);
      if (i === si || !book.sheets[i]) return;
      S()._view = { sel: sel, act: act, st: wrap.scrollTop, sl: wrap.scrollLeft };
      si = i;
      const v = S()._view;
      sel = v ? v.sel : { r1: 0, c1: 0, r2: 0, c2: 0 };
      act = v ? v.act : { r: 0, c: 0 };
      rowCount = 100; colCount = 26;
      layout();
      wrap.scrollTop = v ? v.st : 0;
      wrap.scrollLeft = v ? v.sl : 0;
      render();
      renderTabs();
      focusGrid();
    }
    function uniqueName(base) {
      let n = 1, name = base;
      while (sheetIndex(name) >= 0) name = base.replace(/\d+$/, '') + (++n);
      return name;
    }
    function renameSheet(i) {
      const old = book.sheets[i].name;
      const name = prompt('시트 이름', old);
      if (!name || name === old) return;
      if (/[\\/?*[\]:]/.test(name) || name.length > 31) { DV.toast('시트 이름에 \\ / ? * [ ] : 를 쓸 수 없고 31자 이하여야 합니다.', 'warn'); return; }
      if (sheetIndex(name) >= 0) { DV.toast('같은 이름의 시트가 이미 있습니다.', 'warn'); return; }
      snapshot(function () {
        book.sheets.forEach(function (sh) { sh.rows.forEach(function (row) { if (row) row.forEach(function (cell) { if (cell && cell.f) cell.f = F().renameSheet(cell.f, old, name); }); }); });
        book.sheets[i].name = name;
      });
    }
    el.querySelector('.tab-add').addEventListener('click', function () {
      snapshot(function () { book.sheets.push(DV.xlsxIO.newSheet(uniqueName('Sheet' + (book.sheets.length + 1)))); });
      switchSheet(book.sheets.length - 1);
    });
    tabList.addEventListener('click', function (e) { const b = e.target.closest('.stab'); if (b) switchSheet(+b.dataset.i); });
    tabList.addEventListener('dblclick', function (e) { const b = e.target.closest('.stab'); if (b) renameSheet(+b.dataset.i); });
    tabList.addEventListener('contextmenu', function (e) {
      const b = e.target.closest('.stab');
      if (!b) return;
      e.preventDefault();
      const i = +b.dataset.i;
      switchSheet(i);
      ctxmenu.innerHTML = '<button type="button" data-tab="rename">이름 바꾸기</button><button type="button" data-tab="dup">복제</button><button type="button" data-tab="left">왼쪽으로 이동</button><button type="button" data-tab="right">오른쪽으로 이동</button><hr><button type="button" data-tab="del">시트 삭제</button>';
      const er = el.getBoundingClientRect();
      ctxmenu.hidden = false;
      ctxmenu.style.left = Math.min(e.clientX - er.left, er.width - 200) + 'px';
      ctxmenu.style.top = Math.max(0, e.clientY - er.top - ctxmenu.offsetHeight) + 'px';
      ctxmenu.onclick = function (ev) {
        const t = ev.target.closest('[data-tab]');
        if (!t) return;
        ev.stopPropagation();
        ctxmenu.hidden = true;
        ctxmenu.onclick = null;
        const a = t.dataset.tab;
        if (a === 'rename') renameSheet(i);
        else if (a === 'dup') {
          snapshot(function () { const c = clone(book.sheets[i]); delete c._view; c.name = uniqueName(book.sheets[i].name + ' (2)'); book.sheets.splice(i + 1, 0, c); });
          switchSheet(i + 1);
        } else if (a === 'left' && i > 0) {
          snapshot(function () { const s = book.sheets.splice(i, 1)[0]; book.sheets.splice(i - 1, 0, s); si = i - 1; });
        } else if (a === 'right' && i < book.sheets.length - 1) {
          snapshot(function () { const s = book.sheets.splice(i, 1)[0]; book.sheets.splice(i + 1, 0, s); si = i + 1; });
        } else if (a === 'del') {
          if (book.sheets.length <= 1) { DV.toast('시트가 하나뿐이라 삭제할 수 없습니다.', 'warn'); return; }
          if (!confirm('"' + book.sheets[i].name + '" 시트를 삭제할까요?')) return;
          snapshot(function () { book.sheets.splice(i, 1); si = Math.max(0, Math.min(si, book.sheets.length - 1)); sel = { r1: 0, c1: 0, r2: 0, c2: 0 }; act = { r: 0, c: 0 }; });
        }
      };
    });

    /* ---------- 초기화 ---------- */
    layout();
    renderTabs();
    requestAnimationFrame(function () { render(); });
    setTimeout(function () { render(); focusGrid(); }, 30);

    function valueForExport(sIdx, r, c) {
      const v = value(sIdx, r, c);
      if (v instanceof F().Err) return { e: v.e };
      return v;
    }

    return {
      el: el,
      kind: 'sheet',
      onChange: function (f) { listeners.push(f); },
      refresh: function () { layout(); render(); },
      focus: focusGrid,
      commitEdit: function () { if (editing) commit(0, 0); },
      getBook: function () { return book; },
      exportXlsx: function () { if (editing) commit(0, 0); recalc(); return DV.xlsxIO.writeXlsx(book, valueForExport); },
      exportCSV: function () {
        if (editing) commit(0, 0);
        recalc();
        const sh = S();
        const u = used(sh);
        const rows = [];
        for (let r = 0; r < u.rows; r++) {
          const row = [];
          for (let c = 0; c < u.cols; c++) { const cell = getCell(sh, r, c); row.push(cell ? display(cell, value(si, r, c)) : ''); }
          rows.push(row);
        }
        return DV.xlsxIO.writeCSV(rows);
      },
      sheetName: function () { return S().name; },
      printHTML: function () {
        recalc();
        const sh = S();
        const u = used(sh);
        let h = '<table class="print-sheet"><colgroup>';
        for (let c = 0; c < u.cols; c++) h += '<col style="width:' + colW(c) + 'px">';
        h += '</colgroup>';
        const covered = new Set();
        for (let r = 0; r < u.rows; r++) {
          h += '<tr style="height:' + rowH(r) + 'px">';
          for (let c = 0; c < u.cols; c++) {
            if (covered.has(r * 20000 + c)) continue;
            const m = mergeMap.get(r * 20000 + c);
            let span = '';
            if (m) {
              for (let i = m.r1; i <= m.r2; i++) for (let j = m.c1; j <= m.c2; j++) if (i !== r || j !== c) covered.add(i * 20000 + j);
              span = ' rowspan="' + (m.r2 - m.r1 + 1) + '" colspan="' + (m.c2 - m.c1 + 1) + '"';
            }
            const cell = getCell(sh, r, c);
            const v = cell ? value(si, r, c) : null;
            const css = cell ? cellStyle(cell, v).replace(/justify-content:[^;]+;?/g, '').replace(/align-items:flex-start/, 'vertical-align:top').replace(/align-items:center/, 'vertical-align:middle') : '';
            h += '<td' + span + (css ? ' style="' + css + '"' : '') + '>' + DV.esc(cell ? display(cell, v) : '') + '</td>';
          }
          h += '</tr>';
        }
        return h + '</table>';
      },
      destroy: function () { if (ro) ro.disconnect(); el.remove(); }
    };
  }

  DV.createSheetEditor = createSheetEditor;
})(window.DV);
