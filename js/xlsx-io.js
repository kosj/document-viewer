/* 스프레드시트 파일 읽기(SheetJS + 서식 보강) / XLSX·CSV 쓰기 */
window.DV = window.DV || {};
(function (DV) {
  'use strict';

  const INDEXED = ['000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF', '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
    '800000', '008000', '000080', '808000', '800080', '008080', 'C0C0C0', '808080', '9999FF', '993366', 'FFFFCC', 'CCFFFF', '660066', 'FF8080', '0066CC', 'CCCCFF',
    '000080', 'FF00FF', 'FFFF00', '00FFFF', '800080', '800000', '008080', '0000FF', '00CCFF', 'CCFFFF', 'CCFFCC', 'FFFF99', '99CCFF', 'FF99CC', 'CC99FF', 'FFCC99',
    '3366FF', '33CCCC', '99CC00', 'FFCC00', 'FF9900', 'FF6600', '666699', '969696', '003366', '339966', '003300', '333300', '993300', '993366', '333399', '333333'];
  const DEFAULT_COL_PX = 72;
  const DEFAULT_ROW_PX = 22;

  function newSheet(name) {
    return { name: name, rows: [], colWidths: [], rowHeights: [], merges: [] };
  }

  function applyTint(hex, tint) {
    if (!tint) return hex;
    let r = parseInt(hex.slice(0, 2), 16) / 255, g = parseInt(hex.slice(2, 4), 16) / 255, b = parseInt(hex.slice(4, 6), 16) / 255;
    const max = Math.max(r, g, b), min = Math.min(r, g, b);
    let h = 0, s = 0, l = (max + min) / 2;
    if (max !== min) {
      const d = max - min;
      s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
      h = max === r ? (g - b) / d + (g < b ? 6 : 0) : max === g ? (b - r) / d + 2 : (r - g) / d + 4;
      h /= 6;
    }
    l = tint < 0 ? l * (1 + tint) : l * (1 - tint) + tint;
    function hue(p, q, t) { if (t < 0) t += 1; if (t > 1) t -= 1; if (t < 1 / 6) return p + (q - p) * 6 * t; if (t < 1 / 2) return q; if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6; return p; }
    if (s === 0) r = g = b = l;
    else { const q = l < 0.5 ? l * (1 + s) : l + s - l * s, p = 2 * l - q; r = hue(p, q, h + 1 / 3); g = hue(p, q, h); b = hue(p, q, h - 1 / 3); }
    return [r, g, b].map(function (x) { return ('0' + Math.round(x * 255).toString(16)).slice(-2); }).join('').toUpperCase();
  }

  function colorOf(el, theme) {
    if (!el) return null;
    const rgb = el.getAttribute('rgb');
    const tint = parseFloat(el.getAttribute('tint') || '0');
    if (rgb) return applyTint(rgb.slice(-6).toUpperCase(), tint);
    const th = el.getAttribute('theme');
    if (th != null && theme[+th]) return applyTint(theme[+th], tint);
    const ix = el.getAttribute('indexed');
    if (ix != null) { const i = +ix; if (i === 64) return null; if (INDEXED[i]) return applyTint(INDEXED[i], tint); }
    return null;
  }

  /* xlsx 안의 styles.xml로 셀 서식(글꼴·채우기·맞춤)을 보강 */
  async function readXlsxStyles(buf, sheetNames) {
    const zip = await JSZip.loadAsync(buf);
    const theme = [];
    const tf = zip.file('xl/theme/theme1.xml');
    if (tf) {
      const tdoc = DV.parseXml(await tf.async('string'));
      const cs = DV.desc(tdoc, 'clrScheme')[0];
      if (cs) {
        const map = {};
        DV.kids(cs).forEach(function (c) {
          const v = c.firstElementChild;
          if (!v) return;
          map[c.localName] = (v.getAttribute('lastClr') || v.getAttribute('val') || '000000').toUpperCase();
        });
        ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'].forEach(function (k, i) { theme[i] = map[k]; });
      }
    }
    const sf = zip.file('xl/styles.xml');
    if (!sf) return null;
    const sdoc = DV.parseXml(await sf.async('string'));
    const fonts = DV.kids(DV.desc(sdoc, 'fonts')[0], 'font').map(function (f) {
      const o = {};
      if (DV.kid(f, 'b') && DV.kid(f, 'b').getAttribute('val') !== '0') o.b = true;
      if (DV.kid(f, 'i') && DV.kid(f, 'i').getAttribute('val') !== '0') o.i = true;
      const u = DV.kid(f, 'u');
      if (u && u.getAttribute('val') !== 'none') o.u = true;
      if (DV.kid(f, 'strike') && DV.kid(f, 'strike').getAttribute('val') !== '0') o.strike = true;
      const c = colorOf(DV.kid(f, 'color'), theme);
      if (c && c !== '000000') o.color = c;
      const sz = DV.kid(f, 'sz');
      if (sz) o.size = parseFloat(sz.getAttribute('val'));
      const nm = DV.kid(f, 'name');
      if (nm) o.font = nm.getAttribute('val');
      return o;
    });
    const baseSize = fonts[0] && fonts[0].size || 11;
    const fills = DV.kids(DV.desc(sdoc, 'fills')[0], 'fill').map(function (f) {
      const pf = DV.kid(f, 'patternFill');
      if (!pf || (pf.getAttribute('patternType') || 'none') === 'none') return null;
      return colorOf(DV.kid(pf, 'fgColor'), theme) || colorOf(DV.kid(pf, 'bgColor'), theme);
    });
    const borders = DV.kids(DV.desc(sdoc, 'borders')[0], 'border').map(function (b) {
      const o = {};
      ['left', 'right', 'top', 'bottom'].forEach(function (s) {
        const e = DV.kid(b, s);
        if (e && e.getAttribute('style') && e.getAttribute('style') !== 'none') o[s] = 1;
      });
      return Object.keys(o).length ? o : null;
    });
    const xfs = DV.kids(DV.desc(sdoc, 'cellXfs')[0], 'xf').map(function (x) {
      const st = {};
      const font = fonts[+x.getAttribute('fontId') || 0] || {};
      Object.keys(font).forEach(function (k) { if (k !== 'size' && k !== 'font') st[k] = font[k]; });
      if (font.size && font.size !== baseSize) st.size = font.size;
      const fill = fills[+x.getAttribute('fillId') || 0];
      if (fill) st.bg = fill;
      const bd = borders[+x.getAttribute('borderId') || 0];
      if (bd) st.border = bd;
      const al = DV.kid(x, 'alignment');
      if (al) {
        const h = al.getAttribute('horizontal');
        if (h === 'center' || h === 'centerContinuous') st.align = 'center';
        else if (h === 'right') st.align = 'right';
        else if (h === 'left') st.align = 'left';
        const v = al.getAttribute('vertical');
        if (v === 'center') st.valign = 'middle';
        else if (v === 'top') st.valign = 'top';
        if (al.getAttribute('wrapText') === '1' || al.getAttribute('wrapText') === 'true') st.wrap = true;
      }
      return Object.keys(st).length ? st : null;
    });
    // 시트별 셀 → 스타일 번호
    const wb = DV.parseXml(await zip.file('xl/workbook.xml').async('string'));
    const relsF = zip.file('xl/_rels/workbook.xml.rels');
    const rels = {};
    if (relsF) DV.desc(DV.parseXml(await relsF.async('string')), 'Relationship').forEach(function (r) { rels[r.getAttribute('Id')] = r.getAttribute('Target'); });
    const result = {};
    for (const s of DV.desc(wb, 'sheet')) {
      const name = s.getAttribute('name');
      const rid = DV.attr(s, 'id');
      let target = rels[rid];
      if (!target) continue;
      target = target.replace(/^\//, '');
      if (!/^xl\//.test(target)) target = 'xl/' + target;
      const f = zip.file(target);
      if (!f) continue;
      const xml = await f.async('string');
      const map = {};
      const re = /<c\b([^>]*?)\/?>/g;
      let m;
      while ((m = re.exec(xml))) {
        const a = m[1];
        const rm = /\br="([A-Z]+\d+)"/.exec(a);
        const sm = /\bs="(\d+)"/.exec(a);
        if (rm && sm && xfs[+sm[1]]) map[rm[1]] = xfs[+sm[1]];
      }
      result[name] = map;
    }
    void sheetNames;
    return { cells: result, baseSize: baseSize };
  }

  function decodeText(bytes) {
    if (bytes[0] === 0xEF && bytes[1] === 0xBB && bytes[2] === 0xBF) return new TextDecoder('utf-8').decode(bytes.subarray(3));
    if (bytes[0] === 0xFF && bytes[1] === 0xFE) return new TextDecoder('utf-16le').decode(bytes.subarray(2));
    try { return new TextDecoder('utf-8', { fatal: true }).decode(bytes); }
    catch (e) { return new TextDecoder('euc-kr').decode(bytes); }
  }

  /* CSV: 직접 해석 (따옴표, 줄바꿈 포함 셀 지원) */
  function parseCSV(text) {
    const delim = (function () {
      const line = text.split(/\r?\n/)[0] || '';
      const counts = { ',': (line.match(/,/g) || []).length, '\t': (line.match(/\t/g) || []).length, ';': (line.match(/;/g) || []).length };
      return Object.keys(counts).sort(function (a, b) { return counts[b] - counts[a]; })[0] || ',';
    })();
    const rows = [];
    let row = [], cell = '', q = false;
    for (let i = 0; i < text.length; i++) {
      const ch = text[i];
      if (q) {
        if (ch === '"') { if (text[i + 1] === '"') { cell += '"'; i++; } else q = false; }
        else cell += ch;
      } else if (ch === '"' && cell === '') q = true;
      else if (ch === delim) { row.push(cell); cell = ''; }
      else if (ch === '\n' || ch === '\r') {
        if (ch === '\r' && text[i + 1] === '\n') i++;
        row.push(cell); rows.push(row); row = []; cell = '';
      } else cell += ch;
    }
    if (cell !== '' || row.length) { row.push(cell); rows.push(row); }
    return rows;
  }

  async function readWorkbook(buf, ext) {
    const bytes = new Uint8Array(buf);
    const book = { sheets: [] };
    if (ext === 'csv' || ext === 'tsv' || ext === 'txt') {
      const sh = newSheet('Sheet1');
      parseCSV(decodeText(bytes)).forEach(function (r, ri) {
        sh.rows[ri] = r.map(function (v) { return v === '' ? undefined : DV.sheetInput.parse(v); });
      });
      book.sheets.push(sh);
      return book;
    }
    const wb = XLSX.read(bytes, { type: 'array', cellFormula: true, cellNF: true, cellStyles: true, cellDates: false });
    let extra = null;
    if (bytes[0] === 0x50 && bytes[1] === 0x4B && (ext === 'xlsx' || ext === 'xlsm')) {
      try { extra = await readXlsxStyles(buf, wb.SheetNames); } catch (e) { extra = null; }
    }
    wb.SheetNames.forEach(function (name) {
      const ws = wb.Sheets[name];
      const sh = newSheet(name);
      const smap = extra && extra.cells[name] || {};
      if (ws && ws['!ref']) {
        const rg = XLSX.utils.decode_range(ws['!ref']);
        const maxR = Math.min(rg.e.r, 200000), maxC = Math.min(rg.e.c, 1000);
        for (let r = rg.s.r; r <= maxR; r++) {
          for (let c = rg.s.c; c <= maxC; c++) {
            const a = XLSX.utils.encode_cell({ r: r, c: c });
            const x = ws[a];
            const st = smap[a];
            if (!x && !st) continue;
            const cell = {};
            if (x) {
              if (x.t === 'n') cell.v = x.v;
              else if (x.t === 'b') cell.v = !!x.v;
              else if (x.t === 's' || x.t === 'str') cell.v = x.v == null ? '' : String(x.v);
              else if (x.t === 'e') { cell.v = x.w || '#N/A'; cell.e = 1; }
              else if (x.t === 'd') cell.v = DV.formula.dateSerial(x.v.getFullYear(), x.v.getMonth() + 1, x.v.getDate());
              if (x.f) cell.f = x.f;
              if (x.z && x.z !== 'General' && typeof x.z === 'string') cell.z = x.z;
              if (!extra && x.s && x.s.fgColor && x.s.fgColor.rgb && x.s.patternType === 'solid') cell.s = { bg: x.s.fgColor.rgb.slice(-6).toUpperCase() };
            }
            if (st) cell.s = Object.assign({}, st);
            if (cell.v === undefined && !cell.f && !cell.s) continue;
            (sh.rows[r] = sh.rows[r] || [])[c] = cell;
          }
        }
      }
      (ws && ws['!cols'] || []).forEach(function (col, i) {
        if (!col) return;
        const px = col.wpx || (col.wch ? Math.round(col.wch * 7 + 5) : col.width ? Math.round(col.width * 7 + 5) : 0);
        if (px) sh.colWidths[i] = col.hidden ? 0 : Math.max(16, Math.round(px));
      });
      (ws && ws['!rows'] || []).forEach(function (row, i) {
        if (!row) return;
        const px = row.hpx || (row.hpt ? row.hpt / 0.75 : 0);
        if (px) sh.rowHeights[i] = row.hidden ? 4 : Math.max(14, Math.round(px));
      });
      (ws && ws['!merges'] || []).forEach(function (m) { sh.merges.push({ r1: m.s.r, c1: m.s.c, r2: m.e.r, c2: m.e.c }); });
      book.sheets.push(sh);
    });
    if (!book.sheets.length) book.sheets.push(newSheet('Sheet1'));
    return book;
  }

  /* ---------------- 쓰기 ---------------- */
  const BUILTIN_FMT = { 'General': 0, '0': 1, '0.00': 2, '#,##0': 3, '#,##0.00': 4, '0%': 9, '0.00%': 10, '0.00E+00': 11, 'mm-dd-yy': 14, 'd-mmm-yy': 15, 'h:mm': 20, 'h:mm:ss': 21, 'm/d/yy h:mm': 22, '@': 49 };

  function writeXlsx(book, valueOf) {
    const esc = DV.xmlEsc;
    const sst = [], sstMap = new Map();
    function si(s) {
      if (!sstMap.has(s)) { sstMap.set(s, sst.length); sst.push(s); }
      return sstMap.get(s);
    }
    const fonts = ['<font><sz val="11"/><color theme="1"/><name val="맑은 고딕"/><family val="3"/><charset val="129"/><scheme val="minor"/></font>'], fontMap = { '': 0 };
    const fills = ['<fill><patternFill patternType="none"/></fill>', '<fill><patternFill patternType="gray125"/></fill>'], fillMap = { '': 0 };
    const borders = ['<border><left/><right/><top/><bottom/><diagonal/></border>'], borderMap = { '': 0 };
    const numFmts = [], fmtMap = {};
    const xfs = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'], xfMap = { '': 0 };

    function fontId(s) {
      if (!s || !(s.b || s.i || s.u || s.strike || s.color || s.size || s.font)) return 0;
      const key = [s.b, s.i, s.u, s.strike, s.color, s.size, s.font].join('|');
      if (fontMap[key] == null) {
        fontMap[key] = fonts.length;
        fonts.push('<font>' + (s.b ? '<b/>' : '') + (s.i ? '<i/>' : '') + (s.strike ? '<strike/>' : '') + (s.u ? '<u/>' : '') +
          '<sz val="' + (s.size || 11) + '"/>' + (s.color ? '<color rgb="FF' + s.color + '"/>' : '<color theme="1"/>') +
          '<name val="' + esc(s.font || '맑은 고딕') + '"/><family val="3"/><charset val="129"/></font>');
      }
      return fontMap[key];
    }
    function fillId(s) {
      if (!s || !s.bg) return 0;
      if (fillMap[s.bg] == null) {
        fillMap[s.bg] = fills.length;
        fills.push('<fill><patternFill patternType="solid"><fgColor rgb="FF' + s.bg + '"/><bgColor indexed="64"/></patternFill></fill>');
      }
      return fillMap[s.bg];
    }
    function borderId(s) {
      if (!s || !s.border) return 0;
      const b = s.border;
      const key = ['left', 'right', 'top', 'bottom'].map(function (k) { return b[k] ? 1 : 0; }).join('');
      if (key === '0000') return 0;
      if (borderMap[key] == null) {
        borderMap[key] = borders.length;
        borders.push('<border>' + ['left', 'right', 'top', 'bottom'].map(function (k) { return b[k] ? '<' + k + ' style="thin"><color indexed="64"/></' + k + '>' : '<' + k + '/>'; }).join('') + '<diagonal/></border>');
      }
      return borderMap[key];
    }
    function numFmtId(z) {
      if (!z || z === 'General') return 0;
      if (BUILTIN_FMT[z] != null) return BUILTIN_FMT[z];
      if (fmtMap[z] == null) { fmtMap[z] = 164 + numFmts.length; numFmts.push('<numFmt numFmtId="' + fmtMap[z] + '" formatCode="' + esc(z) + '"/>'); }
      return fmtMap[z];
    }
    function xfId(cell) {
      const s = cell.s, z = cell.z;
      if (!s && !z) return 0;
      const f = fontId(s), fi = fillId(s), b = borderId(s), n = numFmtId(z);
      const al = s && (s.align || s.valign || s.wrap) ? '<alignment' + (s.align ? ' horizontal="' + s.align + '"' : '') + (s.valign ? ' vertical="' + (s.valign === 'middle' ? 'center' : s.valign) + '"' : '') + (s.wrap ? ' wrapText="1"' : '') + '/>' : '';
      const key = [f, fi, b, n, al].join('|');
      if (key === '0|0|0|0|') return 0;
      if (xfMap[key] == null) {
        xfMap[key] = xfs.length;
        xfs.push('<xf numFmtId="' + n + '" fontId="' + f + '" fillId="' + fi + '" borderId="' + b + '" xfId="0"' + (n ? ' applyNumberFormat="1"' : '') + (f ? ' applyFont="1"' : '') + (fi ? ' applyFill="1"' : '') + (b ? ' applyBorder="1"' : '') + (al ? ' applyAlignment="1">' + al + '</xf>' : '/>'));
      }
      return xfMap[key];
    }

    const zip = new JSZip();
    const zfile = zip.file.bind(zip);
    zip.file = function (n, d, o) { return zfile(n, d, Object.assign({ createFolders: false }, o)); };
    const sheetXmls = [];
    book.sheets.forEach(function (sh, si_) {
      let maxR = 0, maxC = 0;
      let data = '';
      sh.rows.forEach(function (row, r) {
        if (!row && !sh.rowHeights[r]) return;
        let cells = '';
        (row || []).forEach(function (cell, c) {
          if (!cell) return;
          const ref = DV.formula.addr(r, c);
          const s = xfId(cell);
          const sa = s ? ' s="' + s + '"' : '';
          if (cell.f) {
            const v = valueOf(si_, r, c);
            let t = '', vx = '';
            if (typeof v === 'number' && isFinite(v)) vx = '<v>' + v + '</v>';
            else if (typeof v === 'boolean') { t = ' t="b"'; vx = '<v>' + (v ? 1 : 0) + '</v>'; }
            else if (v && v.e) { t = ' t="e"'; vx = '<v>' + esc(v.e) + '</v>'; }
            else if (v != null) { t = ' t="str"'; vx = '<v>' + esc(String(v)) + '</v>'; }
            cells += '<c r="' + ref + '"' + sa + t + '><f>' + esc(cell.f) + '</f>' + vx + '</c>';
          } else if (cell.e) {
            cells += '<c r="' + ref + '"' + sa + ' t="e"><v>' + esc(cell.v) + '</v></c>';
          } else if (typeof cell.v === 'number') {
            cells += '<c r="' + ref + '"' + sa + '><v>' + cell.v + '</v></c>';
          } else if (typeof cell.v === 'boolean') {
            cells += '<c r="' + ref + '"' + sa + ' t="b"><v>' + (cell.v ? 1 : 0) + '</v></c>';
          } else if (cell.v != null && cell.v !== '') {
            cells += '<c r="' + ref + '"' + sa + ' t="s"><v>' + si(String(cell.v)) + '</v></c>';
          } else if (s) {
            cells += '<c r="' + ref + '"' + sa + '/>';
          } else return;
          maxR = Math.max(maxR, r); maxC = Math.max(maxC, c);
        });
        const ht = sh.rowHeights[r];
        if (cells || ht) {
          data += '<row r="' + (r + 1) + '"' + (ht ? ' ht="' + (ht * 0.75).toFixed(2) + '" customHeight="1"' : '') + '>' + cells + '</row>';
          maxR = Math.max(maxR, r);
        }
      });
      let cols = '';
      sh.colWidths.forEach(function (w, c) {
        if (w == null) return;
        const width = w === 0 ? 0 : Math.max(1, (w - 5) / 7);
        cols += '<col min="' + (c + 1) + '" max="' + (c + 1) + '" width="' + width.toFixed(2) + '" customWidth="1"' + (w === 0 ? ' hidden="1"' : '') + '/>';
      });
      const merges = sh.merges.filter(function (m) { return m.r2 > m.r1 || m.c2 > m.c1; });
      const x = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
        '<dimension ref="A1:' + DV.formula.addr(maxR, maxC) + '"/><sheetViews><sheetView workbookViewId="0"' + (si_ === 0 ? ' tabSelected="1"' : '') + '/></sheetViews>' +
        '<sheetFormatPr defaultRowHeight="16.5"/>' + (cols ? '<cols>' + cols + '</cols>' : '') + '<sheetData>' + data + '</sheetData>' +
        (merges.length ? '<mergeCells count="' + merges.length + '">' + merges.map(function (m) { return '<mergeCell ref="' + DV.formula.addr(m.r1, m.c1) + ':' + DV.formula.addr(m.r2, m.c2) + '"/>'; }).join('') + '</mergeCells>' : '') +
        '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>';
      sheetXmls.push(x);
    });

    const n = book.sheets.length;
    let ct = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>';
    for (let i = 1; i <= n; i++) ct += '<Override PartName="/xl/worksheets/sheet' + i + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>';
    ct += '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
      '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>' +
      '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>';
    zip.file('[Content_Types].xml', ct);
    zip.file('_rels/.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>');
    let wbx = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><bookViews><workbookView xWindow="0" yWindow="0" windowWidth="28800" windowHeight="15000" activeTab="0"/></bookViews><sheets>';
    book.sheets.forEach(function (sh, i) { wbx += '<sheet name="' + esc(sh.name) + '" sheetId="' + (i + 1) + '" r:id="rId' + (i + 1) + '"/>'; });
    wbx += '</sheets><calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>';
    zip.file('xl/workbook.xml', wbx);
    let rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
    for (let i = 1; i <= n; i++) rels += '<Relationship Id="rId' + i + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' + i + '.xml"/>';
    rels += '<Relationship Id="rId' + (n + 1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId' + (n + 2) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>';
    zip.file('xl/_rels/workbook.xml.rels', rels);
    sheetXmls.forEach(function (x, i) { zip.file('xl/worksheets/sheet' + (i + 1) + '.xml', x); });
    zip.file('xl/sharedStrings.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="' + sst.length + '" uniqueCount="' + sst.length + '">' +
      sst.map(function (s) { return '<si><t xml:space="preserve">' + esc(s) + '</t></si>'; }).join('') + '</sst>');
    zip.file('xl/styles.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      (numFmts.length ? '<numFmts count="' + numFmts.length + '">' + numFmts.join('') + '</numFmts>' : '') +
      '<fonts count="' + fonts.length + '">' + fonts.join('') + '</fonts><fills count="' + fills.length + '">' + fills.join('') + '</fills>' +
      '<borders count="' + borders.length + '">' + borders.join('') + '</borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
      '<cellXfs count="' + xfs.length + '">' + xfs.join('') + '</cellXfs><cellStyles count="1"><cellStyle name="표준" xfId="0" builtinId="0"/></cellStyles><dxfs count="0"/><tableStyles count="0"/></styleSheet>');
    const now = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
    zip.file('docProps/core.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:creator>문서 편집기</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">' + now + '</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">' + now + '</dcterms:modified></cp:coreProperties>');
    zip.file('docProps/app.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Microsoft Excel</Application></Properties>');
    return zip.generateAsync({ type: 'blob', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', compression: 'DEFLATE' });
  }

  function writeCSV(rows) {
    const lines = rows.map(function (r) {
      return r.map(function (v) {
        v = v == null ? '' : String(v);
        return /[",\r\n]/.test(v) || /^\s|\s$/.test(v) ? '"' + v.replace(/"/g, '""') + '"' : v;
      }).join(',');
    });
    return new Blob(['﻿' + lines.join('\r\n')], { type: 'text/csv;charset=utf-8' });
  }

  DV.xlsxIO = {
    readWorkbook: readWorkbook, writeXlsx: writeXlsx, writeCSV: writeCSV, newSheet: newSheet,
    DEFAULT_COL_PX: DEFAULT_COL_PX, DEFAULT_ROW_PX: DEFAULT_ROW_PX
  };
})(window.DV);
