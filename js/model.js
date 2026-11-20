/*
 * 편집기 HTML(DOM) → 중간 문서 모델
 * DOCX/HWPX 내보내기에서 공통으로 사용합니다.
 *
 * Block = { type:'p', align, indentPt, firstLinePt, heading, list, pageBreakBefore, hr, runs:[Run] }
 *       | { type:'table', border, widthPx, colWidthsPx:[], rows:[[Cell]] }
 * Cell  = { colspan, rowspan, bg, valign, blocks:[Block] }
 * Run   = { text, b,i,u,s,sup,sub,color,bg,size,font,href } | { br:true } | { tab:true } | { img:{src,w,h} }
 */
window.DV = window.DV || {};
(function (DV) {
  'use strict';

  const BLOCK = {
    P: 1, DIV: 1, H1: 1, H2: 1, H3: 1, H4: 1, H5: 1, H6: 1, LI: 1, UL: 1, OL: 1, BLOCKQUOTE: 1,
    PRE: 1, TABLE: 1, HR: 1, SECTION: 1, ARTICLE: 1, HEADER: 1, FOOTER: 1, ADDRESS: 1, CENTER: 1,
    FIGURE: 1, FIGCAPTION: 1, DL: 1, DT: 1, DD: 1, NAV: 1, ASIDE: 1, MAIN: 1
  };
  const CONTAINER_ONLY = { UL: 1, OL: 1, DL: 1, SECTION: 1, ARTICLE: 1, NAV: 1, MAIN: 1 };
  const FONT_SIZES = { 1: 7.5, 2: 10, 3: 12, 4: 13.5, 5: 18, 6: 24, 7: 36 };
  const DEFAULT_PT = 10;

  function normAlign(a) {
    if (!a) return null;
    a = String(a).toLowerCase().trim();
    if (a === 'start') return 'left';
    if (a === 'end') return 'right';
    if (a === 'left' || a === 'right' || a === 'center' || a === 'justify') return a;
    return null;
  }

  function firstFont(ff) {
    if (!ff) return null;
    const f = String(ff).split(',')[0].trim().replace(/^["']|["']$/g, '');
    if (!f || /^(serif|sans-serif|monospace|inherit|initial|system-ui|-apple-system)$/i.test(f)) return null;
    return f;
  }

  function applyInline(el, fmt) {
    const f = Object.assign({}, fmt);
    const tag = el.tagName;
    switch (tag) {
      case 'B': case 'STRONG': f.b = true; break;
      case 'I': case 'EM': case 'CITE': case 'VAR': case 'DFN': f.i = true; break;
      case 'U': case 'INS': f.u = true; break;
      case 'S': case 'STRIKE': case 'DEL': f.s = true; break;
      case 'SUP': f.sup = true; f.sub = false; break;
      case 'SUB': f.sub = true; f.sup = false; break;
      case 'MARK': f.bg = 'FFFF00'; break;
      case 'CODE': case 'KBD': case 'SAMP': case 'TT': f.font = 'Consolas'; break;
      case 'A': {
        const h = el.getAttribute('href');
        if (h && /^(https?:|mailto:)/i.test(h)) { f.href = h; f.u = true; f.color = f.color || '0563C1'; }
        break;
      }
      case 'FONT': {
        const c = DV.colorHex(el.getAttribute('color'));
        if (c) f.color = c;
        const face = firstFont(el.getAttribute('face'));
        if (face) f.font = face;
        const sz = el.getAttribute('size');
        if (sz && FONT_SIZES[sz]) f.size = FONT_SIZES[sz];
        break;
      }
    }
    const st = el.style;
    if (st) {
      const fw = st.fontWeight;
      if (fw) {
        if (fw === 'bold' || fw === 'bolder' || +fw >= 600) f.b = true;
        else if (fw === 'normal' || fw === 'lighter' || +fw < 600) f.b = false;
      }
      if (st.fontStyle) f.i = /italic|oblique/.test(st.fontStyle);
      const td = st.textDecorationLine || st.textDecoration;
      if (td) {
        if (/none/.test(td)) { f.u = false; f.s = false; }
        if (/underline/.test(td)) f.u = true;
        if (/line-through/.test(td)) f.s = true;
      }
      const col = DV.colorHex(st.color);
      if (col) f.color = col;
      const bg = DV.colorHex(st.backgroundColor);
      if (bg) f.bg = bg;
      if (st.fontSize) {
        const pt = DV.toPt(st.fontSize, f.size || DEFAULT_PT);
        if (pt) f.size = Math.round(pt * 2) / 2;
      }
      const ff = firstFont(st.fontFamily);
      if (ff) f.font = ff;
      if (st.verticalAlign === 'super') { f.sup = true; f.sub = false; }
      if (st.verticalAlign === 'sub') { f.sub = true; f.sup = false; }
    }
    return f;
  }

  function sameFmt(a, b) {
    const keys = ['b', 'i', 'u', 's', 'sup', 'sub', 'color', 'bg', 'size', 'font', 'href'];
    for (let k = 0; k < keys.length; k++) {
      if ((a[keys[k]] || null) !== (b[keys[k]] || null)) return false;
    }
    return true;
  }

  function pxOf(v) {
    if (v == null || v === '') return null;
    const s = String(v).trim();
    if (/%$/.test(s)) return null;
    const pt = DV.toPt(/^[\d.]+$/.test(s) ? s + 'px' : s);
    return pt == null ? null : pt / 0.75;
  }

  function imageSize(img) {
    let w = pxOf(img.style.width) || pxOf(img.getAttribute('width'));
    let h = pxOf(img.style.height) || pxOf(img.getAttribute('height'));
    const pct = /%$/.test(img.style.width || '') ? parseFloat(img.style.width) : null;
    const nw = img.naturalWidth || 0, nh = img.naturalHeight || 0;
    if (!w && pct) w = DV.PAGE_TEXT_WIDTH_PX * pct / 100;
    if (!w && img.width) w = img.width;
    if (!w) w = nw || 200;
    if (!h || (img.style.height === 'auto')) h = nw ? w * nh / nw : (img.height || w * 0.75);
    if (w > DV.PAGE_TEXT_WIDTH_PX) { h = h * DV.PAGE_TEXT_WIDTH_PX / w; w = DV.PAGE_TEXT_WIDTH_PX; }
    return { w: Math.max(1, Math.round(w)), h: Math.max(1, Math.round(h)) };
  }

  let listSeq = 0;

  function build(root, opts) {
    opts = opts || {};
    const out = [];
    let cur = null;
    let pendingBreak = false;
    const pstack = [{ align: null, indentPt: 0, firstLinePt: 0, heading: 0, list: null }];

    function top() { return pstack[pstack.length - 1]; }

    function ensureP() {
      if (!cur) {
        const t = top();
        cur = { type: 'p', align: t.align, indentPt: t.indentPt, firstLinePt: t.firstLinePt, heading: t.heading, list: null, runs: [] };
        if (t.list) {
          if (!t.list.used) { cur.list = Object.assign({}, t.list); t.list.used = true; }
          else cur.indentPt = (t.list.level + 1) * 20;
        }
        if (pendingBreak) { cur.pageBreakBefore = true; pendingBreak = false; }
        out.push(cur);
      }
      return cur;
    }
    function flush() {
      if (cur) {
        const r = cur.runs;
        if (r.length && r[r.length - 1].br) r.pop();
      }
      cur = null;
    }

    function pushText(text, fmt) {
      if (!text) return;
      text = text.replace(/ /g, ' ').replace(/\r\n?/g, '\n').replace(/[​﻿]/g, '');
      const parts = text.split(/(\n|\t)/);
      for (let i = 0; i < parts.length; i++) {
        const p = parts[i];
        if (!p) continue;
        const para = ensureP();
        if (p === '\n') { para.runs.push({ br: true }); continue; }
        if (p === '\t') { para.runs.push({ tab: true, fmt: fmt }); continue; }
        const last = para.runs[para.runs.length - 1];
        if (last && last.text != null && sameFmt(last, fmt)) last.text += p;
        else para.runs.push(Object.assign({ text: p }, fmt));
      }
    }

    function blockCtx(el, parentCtx) {
      const tag = el.tagName;
      const ctx = Object.assign({}, parentCtx);
      ctx.heading = /^H[1-6]$/.test(tag) ? +tag[1] : 0;
      const a = normAlign(el.style && el.style.textAlign) || normAlign(el.getAttribute('align')) || (tag === 'CENTER' ? 'center' : null);
      if (a) ctx.align = a;
      if (tag !== 'UL' && tag !== 'OL' && tag !== 'LI') {
        let add = 0;
        try {
          const cs = window.getComputedStyle(el);
          add = ((parseFloat(cs.marginLeft) || 0) + (parseFloat(cs.paddingLeft) || 0)) * 0.75;
          const ti = parseFloat(cs.textIndent) || 0;
          ctx.firstLinePt = ti * 0.75;
        } catch (e) { /* 무시 */ }
        if (tag === 'BLOCKQUOTE' && !add) add = 30;
        ctx.indentPt = (parentCtx.indentPt || 0) + Math.max(0, add);
      }
      if (tag === 'LI') {
        const list = el.parentElement;
        const ordered = list && list.tagName === 'OL';
        let level = 0;
        for (let p = list && list.parentElement; p && p !== root; p = p.parentElement) {
          if (p.tagName === 'UL' || p.tagName === 'OL') level++;
        }
        if (!list.__dvListId) list.__dvListId = ++listSeq;
        let index = 0;
        for (let s = el; s; s = s.previousElementSibling) if (s.tagName === 'LI') index++;
        const start = parseInt(list.getAttribute('start'), 10) || 1;
        let style = (list.style && list.style.listStyleType) || list.getAttribute('type') || '';
        ctx.list = { ordered: ordered, level: Math.min(level, 8), listId: list.__dvListId, number: start + index - 1, start: start, style: style, used: false };
        ctx.indentPt = 0;
      } else if (tag !== 'UL' && tag !== 'OL') {
        // 목록 항목 안의 단락: 번호는 첫 단락에만
        if (parentCtx.list) ctx.list = parentCtx.list;
      }
      return ctx;
    }

    function tableModel(table, fmt) {
      const rows = [];
      const trs = Array.prototype.filter.call(table.querySelectorAll('tr'), function (tr) {
        return tr.closest('table') === table;
      });
      const occ = [];
      let ncols = 0;
      const measured = [];
      trs.forEach(function (tr, r) {
        const row = [];
        occ[r] = occ[r] || [];
        let c = 0;
        Array.prototype.forEach.call(tr.children, function (td) {
          if (td.tagName !== 'TD' && td.tagName !== 'TH') return;
          while (occ[r][c]) c++;
          const cs = Math.max(1, parseInt(td.getAttribute('colspan'), 10) || 1);
          const rs = Math.max(1, parseInt(td.getAttribute('rowspan'), 10) || 1);
          for (let i = 0; i < rs; i++) {
            occ[r + i] = occ[r + i] || [];
            for (let j = 0; j < cs; j++) occ[r + i][c + j] = true;
          }
          let cellFmt = applyInline(td, fmt);
          if (td.tagName === 'TH') cellFmt.b = true;
          delete cellFmt.bg;
          const inner = build(td, { fmt: cellFmt, align: td.tagName === 'TH' ? 'center' : null });
          const bg = DV.colorHex(td.style.backgroundColor) || DV.colorHex(td.getAttribute('bgcolor'));
          let va = (td.style.verticalAlign || td.getAttribute('valign') || '').toLowerCase();
          if (va === 'middle') va = 'center';
          const cell = { colspan: cs, rowspan: rs, bg: bg, valign: va || null, blocks: inner, col: c };
          row.push(cell);
          if (cs === 1) {
            const w = td.getBoundingClientRect ? td.getBoundingClientRect().width : 0;
            const sw = pxOf(td.style.width) || pxOf(td.getAttribute('width'));
            if (sw || w) measured[c] = Math.max(measured[c] || 0, sw || w);
          }
          c += cs;
          ncols = Math.max(ncols, c);
        });
        rows.push(row);
      });
      for (let r = 0; r < occ.length; r++) if (occ[r]) ncols = Math.max(ncols, occ[r].length);
      // 열 너비: colgroup → 측정값 → 균등
      const cols = Array.prototype.filter.call(table.querySelectorAll('col'), function (col) { return col.closest('table') === table; });
      const widths = [];
      for (let i = 0; i < ncols; i++) {
        const cw = cols[i] && (pxOf(cols[i].style.width) || pxOf(cols[i].getAttribute('width')));
        widths[i] = cw || measured[i] || 0;
      }
      const known = widths.filter(function (w) { return w > 0; });
      const avg = known.length ? known.reduce(function (a, b) { return a + b; }, 0) / known.length : DV.PAGE_TEXT_WIDTH_PX / Math.max(1, ncols);
      for (let i = 0; i < ncols; i++) if (!widths[i]) widths[i] = avg;
      let total = widths.reduce(function (a, b) { return a + b; }, 0);
      const maxW = opts.maxWidthPx || DV.PAGE_TEXT_WIDTH_PX;
      if (total > maxW || table.style.width === '100%') {
        const k = maxW / total;
        for (let i = 0; i < ncols; i++) widths[i] *= k;
        total = maxW;
      }
      const noBorder = table.classList.contains('dv-noborder') || table.getAttribute('border') === '0';
      return { type: 'table', border: !noBorder, colWidthsPx: widths, widthPx: total, ncols: ncols, rows: rows };
    }

    function walk(node, fmt) {
      if (node.nodeType === 3) {
        const t = node.nodeValue;
        if (!cur && /^[\s]*$/.test(t) && /\n/.test(t)) return; // 블록 사이 서식용 공백
        pushText(t, fmt);
        return;
      }
      if (node.nodeType !== 1) return;
      const el = node;
      const tag = el.tagName;
      if (tag === 'SCRIPT' || tag === 'STYLE' || tag === 'TEMPLATE') return;
      if (el.classList && el.classList.contains('page-break')) {
        flush();
        pendingBreak = true;
        return;
      }
      if (el.classList && el.classList.contains('dv-ui')) return;
      if (tag === 'BR') { ensureP().runs.push({ br: true }); return; }
      if (tag === 'IMG') {
        const src = el.getAttribute('src') || '';
        if (src) ensureP().runs.push({ img: Object.assign({ src: src }, imageSize(el)) });
        return;
      }
      if (tag === 'TABLE') {
        flush();
        const tm = tableModel(el, fmt);
        if (pendingBreak) { out.push({ type: 'p', runs: [], pageBreakBefore: true, align: null, indentPt: 0, heading: 0 }); pendingBreak = false; }
        out.push(tm);
        return;
      }
      if (tag === 'HR') {
        flush();
        out.push({ type: 'p', hr: true, runs: [], align: null, indentPt: top().indentPt, heading: 0 });
        return;
      }
      const f2 = applyInline(el, fmt);
      if (BLOCK[tag]) {
        flush();
        const ctx = blockCtx(el, top());
        if (/^H[1-6]$/.test(tag)) { delete f2.b; }
        pstack.push(ctx);
        const before = out.length;
        for (let ch = el.firstChild; ch; ch = ch.nextSibling) walk(ch, f2);
        flush();
        if (out.length === before && !CONTAINER_ONLY[tag]) { ensureP(); flush(); }
        pstack.pop();
        return;
      }
      for (let ch = el.firstChild; ch; ch = ch.nextSibling) walk(ch, f2);
    }

    if (opts.align) pstack[0].align = opts.align;
    const baseFmt = opts.fmt || {};
    for (let ch = root.firstChild; ch; ch = ch.nextSibling) walk(ch, baseFmt);
    flush();
    if (pendingBreak) out.push({ type: 'p', runs: [], pageBreakBefore: true, align: null, indentPt: 0, heading: 0 });
    return out;
  }

  DV.model = {
    fromDOM: function (root) {
      const blocks = build(root, {});
      if (!blocks.length) blocks.push({ type: 'p', runs: [], align: null, indentPt: 0, heading: 0 });
      return blocks;
    },
    /* 목록 번호 문자열 (HWPX 내보내기에서 글자로 표시) */
    listLabel: function (list) {
      if (!list.ordered) return ['●', '○', '■'][list.level % 3] + ' ';
      const n = list.number;
      const st = (list.style || '').toLowerCase();
      if (st === 'a' || st === 'lower-alpha' || st === 'lower-latin') return String.fromCharCode(96 + ((n - 1) % 26) + 1) + '. ';
      if (st === 'A' || st === 'upper-alpha' || st === 'upper-latin') return String.fromCharCode(64 + ((n - 1) % 26) + 1) + '. ';
      if (st === 'i' || st === 'lower-roman') return roman(n).toLowerCase() + '. ';
      if (st === 'I' || st === 'upper-roman') return roman(n) + '. ';
      if (st === 'hangul') return '가나다라마바사아자차카타파하'.charAt((n - 1) % 14) + '. ';
      if (st === 'hangul-consonant') return 'ㄱㄴㄷㄹㅁㅂㅅㅇㅈㅊㅋㅌㅍㅎ'.charAt((n - 1) % 14) + '. ';
      return n + '. ';
    },
    DEFAULT_PT: DEFAULT_PT
  };

  function roman(n) {
    const map = [[1000, 'M'], [900, 'CM'], [500, 'D'], [400, 'CD'], [100, 'C'], [90, 'XC'], [50, 'L'], [40, 'XL'], [10, 'X'], [9, 'IX'], [5, 'V'], [4, 'IV'], [1, 'I']];
    let s = '';
    map.forEach(function (m) { while (n >= m[0]) { s += m[1]; n -= m[0]; } });
    return s;
  }
})(window.DV);
