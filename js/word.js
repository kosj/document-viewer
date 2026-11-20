/* MS Word (.docx) 가져오기 / 내보내기 */
window.DV = window.DV || {};
(function (DV) {
  'use strict';

  const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const HL = {
    yellow: 'FFFF00', green: '00FF00', cyan: '00FFFF', magenta: 'FF00FF', blue: '0000FF', red: 'FF0000',
    darkBlue: '000080', darkCyan: '008080', darkGreen: '008000', darkMagenta: '800080', darkRed: '800000',
    darkYellow: '808000', darkGray: '808080', lightGray: 'C0C0C0', black: '000000', white: 'FFFFFF'
  };
  const LIST_CSS = {
    decimal: 'decimal', lowerLetter: 'lower-alpha', upperLetter: 'upper-alpha', lowerRoman: 'lower-roman',
    upperRoman: 'upper-roman', ganada: 'hangul', chosung: 'hangul-consonant', decimalEnclosedCircle: 'decimal',
    koreanDigital: 'korean-hangul-formal', koreanCounting: 'korean-hangul-formal'
  };

  const kid = DV.kid, kids = DV.kids, attr = DV.attr, esc = DV.esc;

  function val(el) { return el ? attr(el, 'val') : null; }
  function onOff(el) {
    if (!el) return undefined;
    const v = val(el);
    return !(v === '0' || v === 'false' || v === 'off' || v === 'none');
  }

  function resolvePath(base, target) {
    if (/^\//.test(target)) return target.slice(1);
    const parts = base.split('/');
    parts.pop();
    target.split('/').forEach(function (p) {
      if (p === '..') parts.pop(); else if (p && p !== '.') parts.push(p);
    });
    return parts.join('/');
  }

  async function readRels(zip, path) {
    const f = zip.file(path);
    const out = {};
    if (!f) return out;
    const doc = DV.parseXml(await f.async('string'));
    DV.desc(doc, 'Relationship').forEach(function (r) {
      out[r.getAttribute('Id')] = { target: r.getAttribute('Target'), type: r.getAttribute('Type') || '', external: r.getAttribute('TargetMode') === 'External' };
    });
    return out;
  }

  function parseRPr(rPr) {
    const f = {};
    if (!rPr) return f;
    const b = onOff(kid(rPr, 'b')); if (b !== undefined) f.b = b;
    const i = onOff(kid(rPr, 'i')); if (i !== undefined) f.i = i;
    const st = onOff(kid(rPr, 'strike')), dst = onOff(kid(rPr, 'dstrike'));
    if (st !== undefined || dst !== undefined) f.s = !!(st || dst);
    const u = kid(rPr, 'u');
    if (u) f.u = val(u) !== 'none' && val(u) !== null ? true : (val(u) === null);
    const color = val(kid(rPr, 'color'));
    if (color && color !== 'auto' && /^[0-9A-Fa-f]{6}$/.test(color)) f.color = color.toUpperCase();
    const sz = val(kid(rPr, 'sz'));
    if (sz) f.size = parseInt(sz, 10) / 2;
    const hl = val(kid(rPr, 'highlight'));
    if (hl && HL[hl]) f.bg = HL[hl];
    const shd = kid(rPr, 'shd');
    if (shd && !f.bg) {
      const fill = attr(shd, 'fill');
      if (fill && fill !== 'auto' && /^[0-9A-Fa-f]{6}$/.test(fill) && fill.toUpperCase() !== 'FFFFFF') f.bg = fill.toUpperCase();
    }
    const va = val(kid(rPr, 'vertAlign'));
    if (va === 'superscript') f.sup = true;
    if (va === 'subscript') f.sub = true;
    const fonts = kid(rPr, 'rFonts');
    if (fonts) {
      const name = attr(fonts, 'eastAsia') || attr(fonts, 'ascii') || attr(fonts, 'hAnsi');
      if (name) f.font = name;
    }
    if (onOff(kid(rPr, 'vanish'))) f.hidden = true;
    return f;
  }

  async function loadStyles(zip) {
    const styles = {};
    const f = zip.file('word/styles.xml');
    if (!f) return { styles: styles, defaultPt: null };
    const doc = DV.parseXml(await f.async('string'));
    let defaultPt = null;
    const dd = DV.desc(doc, 'docDefaults')[0];
    if (dd) {
      const sz = DV.desc(dd, 'sz')[0];
      if (sz) defaultPt = parseInt(val(sz), 10) / 2;
    }
    DV.desc(doc, 'style').forEach(function (s) {
      const id = attr(s, 'styleId');
      const nameEl = kid(s, 'name');
      const pPr = kid(s, 'pPr');
      const tblPr = kid(s, 'tblPr');
      styles[id] = {
        type: attr(s, 'type'),
        name: (val(nameEl) || '').toLowerCase(),
        basedOn: val(kid(s, 'basedOn')),
        rPr: parseRPr(kid(s, 'rPr')),
        jc: pPr ? val(kid(pPr, 'jc')) : null,
        outline: pPr && kid(pPr, 'outlineLvl') ? parseInt(val(kid(pPr, 'outlineLvl')), 10) : null,
        numPr: pPr ? kid(pPr, 'numPr') : null,
        borders: tblPr ? hasBorders(kid(tblPr, 'tblBorders')) : null,
        isDefault: attr(s, 'default') === '1'
      };
      const pStyleRpr = s;
      void pStyleRpr;
    });
    return { styles: styles, defaultPt: defaultPt };
  }

  function hasBorders(b) {
    if (!b) return null;
    const list = kids(b);
    if (!list.length) return null;
    return list.some(function (e) { const v = val(e); return v && v !== 'none' && v !== 'nil'; });
  }

  function styleChain(styles, id) {
    const chain = [];
    const seen = {};
    while (id && styles[id] && !seen[id]) {
      seen[id] = 1;
      chain.unshift(styles[id]);
      id = styles[id].basedOn;
    }
    return chain;
  }

  function headingLevel(styles, id) {
    const chain = styleChain(styles, id);
    for (let i = chain.length - 1; i >= 0; i--) {
      const s = chain[i];
      let m = /^heading\s*(\d)$/.exec(s.name);
      if (m) return Math.min(6, +m[1]);
      if (s.name === 'title') return 1;
      if (s.name === 'subtitle') return 2;
    }
    if (id) {
      const m = /^heading(\d)$/i.exec(id);
      if (m) return Math.min(6, +m[1]);
    }
    for (let i = chain.length - 1; i >= 0; i--) {
      if (chain[i].outline != null && chain[i].outline < 6) return chain[i].outline + 1;
    }
    return 0;
  }

  async function loadNumbering(zip) {
    const f = zip.file('word/numbering.xml');
    const abs = {}, nums = {};
    if (!f) return { abs: abs, nums: nums };
    const doc = DV.parseXml(await f.async('string'));
    DV.desc(doc, 'abstractNum').forEach(function (a) {
      const lv = {};
      kids(a, 'lvl').forEach(function (l) {
        lv[attr(l, 'ilvl')] = { fmt: val(kid(l, 'numFmt')) || 'decimal', start: parseInt(val(kid(l, 'start')) || '1', 10), text: val(kid(l, 'lvlText')) || '' };
      });
      abs[attr(a, 'abstractNumId')] = lv;
    });
    DV.desc(doc, 'num').forEach(function (n) {
      const overrides = {};
      kids(n, 'lvlOverride').forEach(function (o) {
        const so = kid(o, 'startOverride');
        if (so) overrides[attr(o, 'ilvl')] = parseInt(val(so), 10);
      });
      nums[attr(n, 'numId')] = { abs: val(kid(n, 'abstractNumId')), overrides: overrides };
    });
    return { abs: abs, nums: nums };
  }

  function fmtKey(f) {
    return [f.b ? 1 : 0, f.i ? 1 : 0, f.u ? 1 : 0, f.s ? 1 : 0, f.sup ? 1 : 0, f.sub ? 1 : 0, f.color || '', f.bg || '', f.size || '', f.font || ''].join('|');
  }

  function wrapFmt(f, inner, defaultPt) {
    let h = inner;
    if (f.sup) h = '<sup>' + h + '</sup>';
    if (f.sub) h = '<sub>' + h + '</sub>';
    if (f.s) h = '<s>' + h + '</s>';
    if (f.u) h = '<u>' + h + '</u>';
    if (f.i) h = '<i>' + h + '</i>';
    if (f.b) h = '<b>' + h + '</b>';
    const css = [];
    if (f.color && f.color !== '000000') css.push('color:#' + f.color);
    if (f.bg) css.push('background-color:#' + f.bg);
    if (f.size && f.size !== defaultPt) css.push('font-size:' + f.size + 'pt');
    if (f.font) css.push('font-family:' + cssFont(f.font));
    if (css.length) h = '<span style="' + esc(css.join(';')) + '">' + h + '</span>';
    return h;
  }

  function cssFont(name) {
    const generic = /바탕|명조|batang|myeongjo|times|serif|georgia|궁서|gungsuh|cambria|mincho|song/i.test(name) && !/sans/i.test(name) ? 'serif' : 'sans-serif';
    return "'" + String(name).replace(/['";]/g, '') + "'," + generic;
  }

  /* ---------------- DOCX → HTML ---------------- */
  async function docxToHtml(buf) {
    const zip = await JSZip.loadAsync(buf);
    let mainPath = 'word/document.xml';
    const rootRels = await readRels(zip, '_rels/.rels');
    Object.keys(rootRels).forEach(function (k) {
      if (/officeDocument$/.test(rootRels[k].type)) mainPath = rootRels[k].target.replace(/^\//, '');
    });
    const main = zip.file(mainPath);
    if (!main) throw new Error('Word 문서 본문(document.xml)을 찾을 수 없습니다. 올바른 .docx 파일인지 확인해 주세요.');
    const relsPath = mainPath.replace(/[^/]+$/, '_rels/') + mainPath.split('/').pop() + '.rels';
    const rels = await readRels(zip, relsPath);
    const st = await loadStyles(zip);
    const numbering = await loadNumbering(zip);
    const images = {};
    const warnings = [];
    await Promise.all(Object.keys(rels).map(async function (id) {
      const r = rels[id];
      if (!/\/image$/.test(r.type) || r.external) return;
      const p = resolvePath(mainPath, r.target);
      const f = zip.file(p);
      if (!f) return;
      const ext = DV.ext(p);
      const mime = DV.mimeFromExt(ext);
      if (!DV.isWebImage(mime)) { images[id] = null; return; }
      images[id] = DV.bytesToDataUrl(await f.async('uint8array'), mime);
    }));

    const doc = DV.parseXml(await main.async('string'));
    const body = DV.desc(doc, 'body')[0];
    const defaultPt = st.defaultPt || 10;
    const ctx = { rels: rels, styles: st.styles, numbering: numbering, images: images, warnings: warnings, defaultPt: defaultPt, listCounters: {}, missingImg: 0 };
    const html = blocksToHtml(kids(body), ctx);
    if (ctx.missingImg) warnings.push('일부 그림(EMF/WMF 등)은 브라우저에서 표시할 수 없어 생략했습니다.');
    return { html: html || '<p><br></p>', defaultPt: defaultPt, warnings: warnings };
  }

  function flattenBlocks(nodes) {
    const out = [];
    nodes.forEach(function (n) {
      const ln = n.localName;
      if (ln === 'sdt') { out.push.apply(out, flattenBlocks(kids(kid(n, 'sdtContent')))); }
      else if (ln === 'customXml' || ln === 'ins' || ln === 'moveTo') { out.push.apply(out, flattenBlocks(kids(n))); }
      else if (ln === 'AlternateContent') {
        const ch = kid(n, 'Choice') || kid(n, 'Fallback');
        if (ch) out.push.apply(out, flattenBlocks(kids(ch)));
      }
      else out.push(n);
    });
    return out;
  }

  function blocksToHtml(nodes, ctx) {
    const items = [];
    flattenBlocks(nodes).forEach(function (n) {
      if (n.localName === 'p') items.push(paragraph(n, ctx));
      else if (n.localName === 'tbl') items.push({ html: table(n, ctx) });
    });
    // 목록 묶기
    let out = '';
    const stack = [];
    function closeTo(level) {
      while (stack.length > level) { out += '</li></' + stack.pop().tag + '>'; }
    }
    items.forEach(function (it) {
      if (it.pageBreakBefore) { closeTo(0); out += '<div class="page-break" contenteditable="false"></div>'; }
      if (!it.list) { closeTo(0); out += it.html; }
      else {
        const L = it.list;
        const lvl = Math.min(8, L.ilvl);
        const tag = L.ordered ? 'ol' : 'ul';
        closeTo(lvl + 1);
        if (stack.length === lvl + 1) {
          const t = stack[lvl];
          if (t.tag === tag && (lvl > 0 || t.numId === L.numId)) {
            out += '</li><li>' + it.inner;
          } else {
            closeTo(lvl);
          }
        }
        if (stack.length < lvl + 1) {
          while (stack.length < lvl + 1) {
            const depth = stack.length;
            const isTarget = depth === lvl;
            const tg = isTarget ? tag : 'ul';
            let a = '';
            if (isTarget && L.ordered) {
              if (L.css && L.css !== 'decimal') a += ' style="list-style-type:' + L.css + '"';
              if (L.startAt && L.startAt !== 1) a += ' start="' + L.startAt + '"';
            }
            out += '<' + tg + a + '>';
            stack.push({ tag: tg, numId: L.numId });
            if (!isTarget) out += '<li style="list-style-type:none">';
          }
          out += '<li>' + it.inner;
        }
      }
      if (it.pageBreakAfter) { closeTo(0); out += '<div class="page-break" contenteditable="false"></div>'; }
    });
    closeTo(0);
    return out;
  }

  function paragraph(p, ctx) {
    const pPr = kid(p, 'pPr');
    const styleId = pPr ? val(kid(pPr, 'pStyle')) : null;
    const styles = ctx.styles;
    let defStyle = null;
    Object.keys(styles).forEach(function (k) { if (styles[k].isDefault && styles[k].type === 'paragraph') defStyle = k; });
    const sid = styleId || defStyle;
    const hl = headingLevel(styles, sid);
    const chain = styleChain(styles, sid);
    // 문단 스타일의 글자 속성 (제목 제외)
    let baseFmt = {};
    if (!hl) chain.forEach(function (s) { Object.assign(baseFmt, s.rPr); });
    if (pPr) {
      const prPr = kid(pPr, 'rPr');
      void prPr; // 문단 기호 서식은 무시
    }
    let jc = pPr ? val(kid(pPr, 'jc')) : null;
    if (!jc) for (let i = chain.length - 1; i >= 0; i--) if (chain[i].jc) { jc = chain[i].jc; break; }
    const css = [];
    const align = { center: 'center', right: 'right', end: 'right', both: 'justify', distribute: 'justify', left: null, start: null }[jc];
    if (align) css.push('text-align:' + align);

    // 목록
    let list = null;
    let numPr = pPr ? kid(pPr, 'numPr') : null;
    if (!numPr) for (let i = chain.length - 1; i >= 0; i--) if (chain[i].numPr) { numPr = chain[i].numPr; break; }
    if (numPr) {
      const numId = val(kid(numPr, 'numId'));
      const ilvl = parseInt(val(kid(numPr, 'ilvl')) || '0', 10);
      const num = numId && ctx.numbering.nums[numId];
      if (num && numId !== '0') {
        const lv = (ctx.numbering.abs[num.abs] || {})[ilvl] || { fmt: 'decimal', start: 1 };
        if (lv.fmt !== 'none') {
          const ordered = lv.fmt !== 'bullet';
          const key = numId + ':' + ilvl;
          const start = num.overrides[ilvl] != null ? num.overrides[ilvl] : lv.start;
          const count = ctx.listCounters[key] || 0;
          ctx.listCounters[key] = count + 1;
          // 상위 수준 새 항목이 나오면 하위 번호 초기화
          Object.keys(ctx.listCounters).forEach(function (k) {
            const parts = k.split(':');
            if (parts[0] === numId && +parts[1] > ilvl) delete ctx.listCounters[k];
          });
          list = { numId: numId, ilvl: ilvl, ordered: ordered, css: LIST_CSS[lv.fmt] || 'decimal', startAt: start + count };
        }
      }
    }
    if (!list && pPr) {
      const ind = kid(pPr, 'ind');
      if (ind) {
        const left = parseInt(attr(ind, 'left') || attr(ind, 'start') || '0', 10);
        const first = parseInt(attr(ind, 'firstLine') || '0', 10);
        const hang = parseInt(attr(ind, 'hanging') || '0', 10);
        if (left > 0) css.push('margin-left:' + Math.round(left / 20) + 'pt');
        if (first > 0) css.push('text-indent:' + Math.round(first / 20) + 'pt');
        else if (hang > 0) css.push('text-indent:-' + Math.round(hang / 20) + 'pt');
      }
    }

    const state = { fld: [], pageBreak: false, hasContent: false, breakAtStart: false };
    const inner = inlines(kids(p), ctx, state, baseFmt);
    const html = inner.html;
    const tag = hl ? 'h' + hl : 'p';
    const style = css.length ? ' style="' + css.join(';') + '"' : '';
    const content = html || '<br>';
    const res = { html: '<' + tag + style + '>' + content + '</' + tag + '>', inner: html || '<br>' };
    if (state.pageBreak) {
      if (state.breakAtStart) res.pageBreakBefore = true; else res.pageBreakAfter = true;
      if (!html.replace(/<br>/g, '')) { res.html = ''; res.pageBreakAfter = false; res.pageBreakBefore = true; if (list) list = null; }
    }
    if (pPr && onOff(kid(pPr, 'pageBreakBefore'))) res.pageBreakBefore = true;
    if (list && !res.html) list = null;
    res.list = list;
    return res;
  }

  /* 인라인 요소들을 HTML로: 같은 서식의 연속 런은 하나로 합칩니다 */
  function inlines(nodes, ctx, state, baseFmt) {
    let html = '';
    let curKey = null, curFmt = null, buf = '';
    function flush() {
      if (buf) html += wrapFmt(curFmt, buf, ctx.defaultPt);
      buf = ''; curKey = null; curFmt = null;
    }
    function add(fmt, h) {
      const k = fmtKey(fmt);
      if (k !== curKey) { flush(); curKey = k; curFmt = fmt; }
      buf += h;
      state.hasContent = true;
    }
    function raw(h) { flush(); html += h; state.hasContent = true; }

    function inField() { return state.fld.some(function (x) { return x === 'instr'; }); }

    function run(r) {
      const rPr = kid(r, 'rPr');
      let fmt = Object.assign({}, baseFmt);
      if (rPr) {
        const rs = val(kid(rPr, 'rStyle'));
        if (rs) styleChain(ctx.styles, rs).forEach(function (s) { Object.assign(fmt, s.rPr); });
        Object.assign(fmt, parseRPr(rPr));
      }
      if (fmt.hidden) return;
      kids(r).forEach(function (c) {
        const ln = c.localName;
        if (ln === 'fldChar') {
          const t = attr(c, 'fldCharType');
          if (t === 'begin') state.fld.push('instr');
          else if (t === 'separate') { if (state.fld.length) state.fld[state.fld.length - 1] = 'result'; }
          else if (t === 'end') state.fld.pop();
          return;
        }
        if (inField()) return;
        switch (ln) {
          case 't': add(fmt, esc(c.textContent)); break;
          case 'tab': case 'ptab': add(fmt, '\t'); break;
          case 'br': case 'cr': {
            if (attr(c, 'type') === 'page') { state.pageBreak = true; if (!state.hasContent) state.breakAtStart = true; }
            else add(fmt, '<br>');
            break;
          }
          case 'noBreakHyphen': add(fmt, '‑'); break;
          case 'sym': {
            const code = parseInt(attr(c, 'char') || '', 16);
            if (code) add(fmt, esc(String.fromCharCode(code >= 0xF000 ? code - 0xF000 : code)));
            break;
          }
          case 'drawing': raw(drawing(c, ctx)); break;
          case 'pict': case 'object': raw(vml(c, ctx)); break;
          case 'AlternateContent': {
            const ch = kid(c, 'Choice');
            if (ch) kids(ch).forEach(function (x) { if (x.localName === 'drawing') raw(drawing(x, ctx)); });
            break;
          }
          default: break;
        }
      });
    }

    nodes.forEach(function walk(n) {
      switch (n.localName) {
        case 'r': run(n); break;
        case 'hyperlink': {
          const id = attr(n, 'id');
          let href = null;
          if (id && ctx.rels[id] && ctx.rels[id].external) href = ctx.rels[id].target;
          const sub = inlines(kids(n), ctx, state, baseFmt);
          if (href && /^(https?:|mailto:)/i.test(href)) raw('<a href="' + esc(href) + '">' + sub.html + '</a>');
          else if (sub.html) raw(sub.html);
          break;
        }
        case 'ins': case 'smartTag': case 'customXml': case 'fldSimple': case 'moveTo': case 'dir': case 'bdo':
          kids(n).forEach(walk); break;
        case 'sdt': kids(kid(n, 'sdtContent')).forEach(walk); break;
        case 'oMath': case 'oMathPara': {
          const t = DV.desc(n, 't').map(function (x) { return x.textContent; }).join('');
          if (t) add(Object.assign({}, baseFmt, { i: true }), esc(t));
          break;
        }
        default: break;
      }
    });
    flush();
    return { html: html };
  }

  function drawing(d, ctx) {
    const blip = DV.desc(d, 'blip')[0];
    const ext = DV.desc(d, 'extent')[0];
    let w = ext ? Math.round(parseInt(attr(ext, 'cx'), 10) / 9525) : null;
    let h = ext ? Math.round(parseInt(attr(ext, 'cy'), 10) / 9525) : null;
    if (blip) {
      const id = attr(blip, 'embed');
      const src = ctx.images[id];
      if (src) {
        const sz = w && h ? ' style="width:' + w + 'px;height:' + h + 'px"' : '';
        return '<img src="' + src + '"' + sz + ' alt="">';
      }
      ctx.missingImg++;
      return '';
    }
    // 텍스트 상자
    const tx = DV.desc(d, 'txbxContent')[0];
    if (tx) {
      const inner = DV.desc(tx, 'p').map(function (p) { return inlines(kids(p), ctx, { fld: [] }, {}).html; }).filter(Boolean).join('<br>');
      if (inner) return '<span class="dv-textbox">' + inner + '</span>';
    }
    return '';
  }

  function vml(el, ctx) {
    const im = DV.desc(el, 'imagedata')[0];
    if (!im) return '';
    const id = attr(im, 'id');
    const src = ctx.images[id];
    if (!src) { ctx.missingImg++; return ''; }
    const shape = DV.desc(el, 'shape')[0];
    let sz = '';
    if (shape) {
      const st = shape.getAttribute('style') || '';
      const mw = /width:([\d.]+)pt/.exec(st), mh = /height:([\d.]+)pt/.exec(st);
      if (mw && mh) sz = ' style="width:' + Math.round(mw[1] / 0.75) + 'px;height:' + Math.round(mh[1] / 0.75) + 'px"';
    }
    return '<img src="' + src + '"' + sz + ' alt="">';
  }

  function table(tbl, ctx) {
    const tblPr = kid(tbl, 'tblPr');
    let bordered = tblPr ? hasBorders(kid(tblPr, 'tblBorders')) : null;
    if (bordered == null && tblPr) {
      const ts = val(kid(tblPr, 'tblStyle'));
      if (ts) {
        const chain = styleChain(ctx.styles, ts);
        for (let i = chain.length - 1; i >= 0; i--) if (chain[i].borders != null) { bordered = chain[i].borders; break; }
      }
    }
    if (bordered == null) {
      bordered = DV.desc(tbl, 'tcBorders').some(function (b) { return hasBorders(b); });
    }
    const grid = kid(tbl, 'tblGrid');
    const cols = grid ? kids(grid, 'gridCol').map(function (g) { return parseInt(attr(g, 'w') || '0', 10); }) : [];
    const totalTw = cols.reduce(function (a, b) { return a + b; }, 0);
    const owners = [];
    const rows = [];
    kids(tbl, 'tr').forEach(function (tr) {
      const row = [];
      let col = 0;
      const trPr = kid(tr, 'trPr');
      if (trPr) col += parseInt(val(kid(trPr, 'gridBefore')) || '0', 10);
      const cells = [];
      kids(tr).forEach(function (n) {
        if (n.localName === 'tc') cells.push(n);
        else if (n.localName === 'sdt') kids(kid(n, 'sdtContent')).forEach(function (x) { if (x.localName === 'tc') cells.push(x); });
      });
      cells.forEach(function (tc) {
        const tcPr = kid(tc, 'tcPr');
        const span = tcPr && kid(tcPr, 'gridSpan') ? parseInt(val(kid(tcPr, 'gridSpan')), 10) || 1 : 1;
        const vm = tcPr ? kid(tcPr, 'vMerge') : null;
        const vmv = vm ? (val(vm) || 'continue') : null;
        if (vmv === 'continue' && owners[col]) {
          owners[col].rowspan++;
          col += span;
          return;
        }
        const css = [];
        if (tcPr) {
          const shd = kid(tcPr, 'shd');
          const fill = shd ? attr(shd, 'fill') : null;
          if (fill && fill !== 'auto' && /^[0-9A-Fa-f]{6}$/.test(fill)) css.push('background-color:#' + fill.toUpperCase());
          const va = val(kid(tcPr, 'vAlign'));
          if (va === 'center') css.push('vertical-align:middle');
          else if (va === 'bottom') css.push('vertical-align:bottom');
        }
        const content = blocksToHtml(kids(tc).filter(function (x) { return x.localName !== 'tcPr'; }), ctx);
        const cell = { colspan: span, rowspan: 1, css: css, html: content || '<p><br></p>' };
        for (let k = 0; k < span; k++) owners[col + k] = null;
        owners[col] = cell;
        row.push(cell);
        col += span;
      });
      rows.push(row);
    });
    let h = '<table class="dv-table' + (bordered ? '' : ' dv-noborder') + '"';
    if (totalTw) h += ' style="width:' + Math.round(totalTw / 20) + 'pt"';
    h += '>';
    if (cols.length && totalTw) {
      h += '<colgroup>' + cols.map(function (w) { return '<col style="width:' + Math.round(w / 20) + 'pt">'; }).join('') + '</colgroup>';
    }
    h += '<tbody>';
    rows.forEach(function (row) {
      h += '<tr>';
      row.forEach(function (c) {
        h += '<td' + (c.colspan > 1 ? ' colspan="' + c.colspan + '"' : '') + (c.rowspan > 1 ? ' rowspan="' + c.rowspan + '"' : '') +
          (c.css.length ? ' style="' + c.css.join(';') + '"' : '') + '>' + c.html + '</td>';
      });
      h += '</tr>';
    });
    h += '</tbody></table>';
    return h;
  }

  /* ---------------- HTML(모델) → DOCX ---------------- */
  const PAGE_W = 11906, MARGIN_LR = 1440, TEXT_W = PAGE_W - MARGIN_LR * 2; // twips

  async function prepareImages(blocks) {
    const list = [];
    (function collect(bs) {
      bs.forEach(function (b) {
        if (b.type === 'p') b.runs.forEach(function (r) { if (r.img) list.push(r.img); });
        else if (b.type === 'table') b.rows.forEach(function (row) { row.forEach(function (c) { collect(c.blocks); }); });
      });
    })(blocks);
    for (const img of list) {
      let src = img.src;
      if (!/^data:/.test(src)) {
        try {
          const resp = await fetch(src);
          const blob = await resp.blob();
          src = await new Promise(function (res) { const fr = new FileReader(); fr.onload = function () { res(fr.result); }; fr.readAsDataURL(blob); });
        } catch (e) { img.skip = true; continue; }
      }
      let info = DV.dataUrlInfo(src);
      if (info && !/^image\/(png|jpeg|gif|bmp)$/.test(info.mime)) {
        const png = await DV.ensurePng(src);
        info = png ? DV.dataUrlInfo(png) : null;
      }
      if (!info) { img.skip = true; continue; }
      img.data = info;
    }
  }

  async function blocksToDocx(blocks, opts) {
    opts = opts || {};
    await prepareImages(blocks);
    const ctx = { rels: [], media: [], nums: [], numMap: {}, docPr: 1, relSeq: 10 };

    function relId(type, target, external) {
      const id = 'rId' + (ctx.relSeq++);
      ctx.rels.push({ id: id, type: type, target: target, external: external });
      return id;
    }

    function numIdFor(list) {
      const key = list.listId + (list.ordered ? 'o' : 'u');
      if (!ctx.numMap[key]) {
        const id = ctx.nums.length + 1;
        ctx.nums.push({ id: id, ordered: list.ordered, start: list.start || 1, style: list.style });
        ctx.numMap[key] = id;
      }
      return ctx.numMap[key];
    }

    function rPrXml(r) {
      let x = '';
      if (r.href) x += '<w:rStyle w:val="Hyperlink"/>';
      if (r.font) x += '<w:rFonts w:ascii="' + DV.xmlEsc(r.font) + '" w:hAnsi="' + DV.xmlEsc(r.font) + '" w:eastAsia="' + DV.xmlEsc(r.font) + '"/>';
      if (r.b) x += '<w:b/><w:bCs/>';
      if (r.i) x += '<w:i/><w:iCs/>';
      if (r.s) x += '<w:strike/>';
      if (r.color) x += '<w:color w:val="' + r.color + '"/>';
      if (r.size) x += '<w:sz w:val="' + Math.round(r.size * 2) + '"/><w:szCs w:val="' + Math.round(r.size * 2) + '"/>';
      if (r.u) x += '<w:u w:val="single"/>';
      if (r.bg) x += '<w:shd w:val="clear" w:color="auto" w:fill="' + r.bg + '"/>';
      if (r.sup) x += '<w:vertAlign w:val="superscript"/>';
      else if (r.sub) x += '<w:vertAlign w:val="subscript"/>';
      return x ? '<w:rPr>' + x + '</w:rPr>' : '';
    }

    function imageRun(img) {
      if (img.skip || !img.data) return '';
      const n = ctx.media.length + 1;
      const name = 'image' + n + '.' + img.data.ext;
      ctx.media.push({ name: name, bytes: img.data.bytes });
      const rid = relId('http://schemas.openxmlformats.org/officeDocument/2006/relationships/image', 'media/' + name);
      const cx = Math.round(img.w * 9525), cy = Math.round(img.h * 9525);
      const id = ctx.docPr++;
      return '<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="' + cx + '" cy="' + cy + '"/><wp:effectExtent l="0" t="0" r="0" b="0"/>' +
        '<wp:docPr id="' + id + '" name="그림 ' + id + '"/><wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/></wp:cNvGraphicFramePr>' +
        '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
        '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="' + id + '" name="' + name + '"/><pic:cNvPicPr/></pic:nvPicPr>' +
        '<pic:blipFill><a:blip r:embed="' + rid + '"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>' +
        '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="' + cx + '" cy="' + cy + '"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>' +
        '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>';
    }

    function runsXml(runs) {
      let x = '';
      let i = 0;
      while (i < runs.length) {
        const r = runs[i];
        if (r.href) {
          const href = r.href;
          let inner = '';
          while (i < runs.length && runs[i].href === href) { inner += oneRun(runs[i]); i++; }
          const rid = relId('http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink', href, true);
          x += '<w:hyperlink r:id="' + rid + '" w:history="1">' + inner + '</w:hyperlink>';
          continue;
        }
        x += oneRun(r);
        i++;
      }
      return x;
    }
    function oneRun(r) {
      if (r.img) return imageRun(r.img);
      if (r.br) return '<w:r><w:br/></w:r>';
      if (r.tab) return '<w:r>' + rPrXml(r.fmt || {}) + '<w:tab/></w:r>';
      if (r.text == null || r.text === '') return '';
      return '<w:r>' + rPrXml(r) + '<w:t xml:space="preserve">' + DV.xmlEsc(r.text) + '</w:t></w:r>';
    }

    function paraXml(p) {
      let pPr = '';
      if (p.heading) pPr += '<w:pStyle w:val="Heading' + p.heading + '"/>';
      else if (p.list) pPr += '<w:pStyle w:val="ListParagraph"/>';
      if (p.pageBreakBefore) pPr += '<w:pageBreakBefore/>';
      if (p.list) pPr += '<w:numPr><w:ilvl w:val="' + p.list.level + '"/><w:numId w:val="' + numIdFor(p.list) + '"/></w:numPr>';
      if (p.hr) pPr += '<w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="auto"/></w:pBdr>';
      if (!p.list && (p.indentPt > 0 || p.firstLinePt)) {
        pPr += '<w:ind' + (p.indentPt > 0 ? ' w:left="' + Math.round(p.indentPt * 20) + '"' : '') +
          (p.firstLinePt > 0 ? ' w:firstLine="' + Math.round(p.firstLinePt * 20) + '"' : '') +
          (p.firstLinePt < 0 ? ' w:hanging="' + Math.round(-p.firstLinePt * 20) + '"' : '') + '/>';
      }
      if (p.align) pPr += '<w:jc w:val="' + ({ left: 'left', center: 'center', right: 'right', justify: 'both' })[p.align] + '"/>';
      return '<w:p>' + (pPr ? '<w:pPr>' + pPr + '</w:pPr>' : '') + runsXml(p.runs) + '</w:p>';
    }

    function tableXml(t, availTw) {
      const totalPx = t.colWidthsPx.reduce(function (a, b) { return a + b; }, 0) || 1;
      let scale = 15;
      if (totalPx * scale > availTw) scale = availTw / totalPx;
      const colTw = t.colWidthsPx.map(function (w) { return Math.max(200, Math.round(w * scale)); });
      const tblW = colTw.reduce(function (a, b) { return a + b; }, 0);
      const bv = t.border ? 'single' : 'nil';
      let x = '<w:tbl><w:tblPr><w:tblStyle w:val="' + (t.border ? 'TableGrid' : 'TableNormal') + '"/><w:tblW w:w="' + tblW + '" w:type="dxa"/>' +
        '<w:tblBorders>' + ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map(function (s) {
          return '<w:' + s + ' w:val="' + bv + '"' + (t.border ? ' w:sz="4" w:space="0" w:color="auto"' : '') + '/>';
        }).join('') + '</w:tblBorders><w:tblLayout w:type="fixed"/>' +
        '<w:tblCellMar><w:left w:w="108" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar>' +
        '<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr>';
      x += '<w:tblGrid>' + colTw.map(function (w) { return '<w:gridCol w:w="' + w + '"/>'; }).join('') + '</w:tblGrid>';
      // 그리드 배치
      const nrows = t.rows.length;
      const grid = [];
      t.rows.forEach(function (row, r) {
        grid[r] = grid[r] || [];
        let c = 0;
        row.forEach(function (cell) {
          while (grid[r][c]) c++;
          const slot = { cell: cell, r0: r, c0: c };
          for (let i = 0; i < cell.rowspan && r + i < nrows; i++) {
            grid[r + i] = grid[r + i] || [];
            for (let j = 0; j < cell.colspan; j++) grid[r + i][c + j] = slot;
          }
          c += cell.colspan;
        });
      });
      const ncols = colTw.length;
      for (let r = 0; r < nrows; r++) {
        x += '<w:tr>';
        let c = 0;
        while (c < ncols) {
          const slot = grid[r] && grid[r][c];
          if (!slot) { x += '<w:tc><w:tcPr><w:tcW w:w="' + colTw[c] + '" w:type="dxa"/></w:tcPr><w:p/></w:tc>'; c++; continue; }
          const cell = slot.cell;
          const cs = Math.min(cell.colspan, ncols - c);
          let w = 0;
          for (let j = 0; j < cs; j++) w += colTw[c + j] || 0;
          let tcPr = '<w:tcW w:w="' + w + '" w:type="dxa"/>';
          if (cs > 1) tcPr += '<w:gridSpan w:val="' + cs + '"/>';
          const isStart = slot.r0 === r;
          if (cell.rowspan > 1) tcPr += isStart ? '<w:vMerge w:val="restart"/>' : '<w:vMerge/>';
          if (cell.bg) tcPr += '<w:shd w:val="clear" w:color="auto" w:fill="' + cell.bg + '"/>';
          if (cell.valign === 'center' || cell.valign === 'bottom') tcPr += '<w:vAlign w:val="' + cell.valign + '"/>';
          let content = '<w:p/>';
          if (isStart && slot.c0 === c) {
            content = blocksXml(cell.blocks, w - 216);
            if (!content || /<\/w:tbl>$/.test(content)) content += '<w:p/>';
          }
          x += '<w:tc><w:tcPr>' + tcPr + '</w:tcPr>' + content + '</w:tc>';
          c += cs;
        }
        x += '</w:tr>';
      }
      return x + '</w:tbl>';
    }

    function blocksXml(bs, availTw) {
      return bs.map(function (b) { return b.type === 'table' ? tableXml(b, availTw) : paraXml(b); }).join('');
    }

    let bodyXml = blocksXml(blocks, TEXT_W);
    if (!bodyXml || /<\/w:tbl>$/.test(bodyXml)) bodyXml += '<w:p/>';

    const NS = 'xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ' +
      'xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
      'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" ' +
      'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" ' +
      'xmlns:w="' + W_NS + '" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"';
    const documentXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document ' + NS + '><w:body>' + bodyXml +
      '<w:sectPr><w:pgSz w:w="' + PAGE_W + '" w:h="16838"/><w:pgMar w:top="1440" w:right="' + MARGIN_LR + '" w:bottom="1440" w:left="' + MARGIN_LR + '" w:header="851" w:footer="992" w:gutter="0"/><w:cols w:space="425"/><w:docGrid w:linePitch="360"/></w:sectPr></w:body></w:document>';

    const zip = new JSZip();
    const zfile = zip.file.bind(zip);
    zip.file = function (n, d, o) { return zfile(n, d, Object.assign({ createFolders: false }, o)); };
    const hasNum = ctx.nums.length > 0;
    let ct = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' +
      '<Default Extension="png" ContentType="image/png"/><Default Extension="jpeg" ContentType="image/jpeg"/><Default Extension="jpg" ContentType="image/jpeg"/>' +
      '<Default Extension="gif" ContentType="image/gif"/><Default Extension="bmp" ContentType="image/bmp"/>' +
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
      '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>' +
      (hasNum ? '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>' : '') +
      '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' +
      '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>';
    zip.file('[Content_Types].xml', ct);
    zip.file('_rels/.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>' +
      '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>');
    let rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>' +
      (hasNum ? '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>' : '');
    ctx.rels.forEach(function (r) {
      rels += '<Relationship Id="' + r.id + '" Type="' + r.type + '" Target="' + DV.xmlEsc(r.target) + '"' + (r.external ? ' TargetMode="External"' : '') + '/>';
    });
    rels += '</Relationships>';
    zip.file('word/_rels/document.xml.rels', rels);
    zip.file('word/document.xml', documentXml);
    zip.file('word/styles.xml', stylesXml());
    zip.file('word/settings.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:settings xmlns:w="' + W_NS + '"><w:defaultTabStop w:val="800"/><w:characterSpacingControl w:val="doNotCompress"/><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>');
    if (hasNum) zip.file('word/numbering.xml', numberingXml(ctx.nums));
    ctx.media.forEach(function (m) { zip.file('word/media/' + m.name, m.bytes); });
    const now = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
    zip.file('docProps/core.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
      '<dc:title>' + DV.xmlEsc(opts.title || '') + '</dc:title><dc:creator>문서 편집기</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">' + now + '</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">' + now + '</dcterms:modified></cp:coreProperties>');
    zip.file('docProps/app.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Document Editor</Application></Properties>');
    return zip.generateAsync({ type: 'blob', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', compression: 'DEFLATE' });
  }

  function stylesXml() {
    const font = '맑은 고딕';
    const heads = [[1, 40], [2, 32], [3, 28], [4, 24], [5, 22], [6, 20]];
    let s = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:styles xmlns:w="' + W_NS + '">' +
      '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="' + font + '" w:eastAsia="' + font + '" w:hAnsi="' + font + '" w:cs="' + font + '"/><w:sz w:val="20"/><w:szCs w:val="20"/><w:lang w:val="en-US" w:eastAsia="ko-KR" w:bidi="ar-SA"/></w:rPr></w:rPrDefault>' +
      '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="384" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>' +
      '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/><w:pPr><w:widowControl w:val="0"/><w:wordWrap w:val="0"/><w:autoSpaceDE w:val="0"/><w:autoSpaceDN w:val="0"/></w:pPr></w:style>';
    heads.forEach(function (h) {
      s += '<w:style w:type="paragraph" w:styleId="Heading' + h[0] + '"><w:name w:val="heading ' + h[0] + '"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="9"/><w:qFormat/>' +
        '<w:pPr><w:keepNext/><w:spacing w:before="' + (h[0] <= 2 ? 240 : 160) + '" w:after="80"/><w:outlineLvl w:val="' + (h[0] - 1) + '"/></w:pPr>' +
        '<w:rPr><w:b/><w:bCs/>' + (h[0] === 6 ? '<w:i/>' : '') + '<w:sz w:val="' + h[1] + '"/><w:szCs w:val="' + h[1] + '"/></w:rPr></w:style>';
    });
    s += '<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/><w:uiPriority w:val="34"/><w:qFormat/><w:pPr><w:ind w:left="800"/></w:pPr></w:style>' +
      '<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/><w:unhideWhenUsed/></w:style>' +
      '<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:basedOn w:val="DefaultParagraphFont"/><w:uiPriority w:val="99"/><w:unhideWhenUsed/><w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr></w:style>' +
      '<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>' +
      '<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:basedOn w:val="TableNormal"/><w:uiPriority w:val="39"/><w:pPr><w:spacing w:line="240" w:lineRule="auto"/></w:pPr><w:tblPr><w:tblBorders>' +
      ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map(function (x) { return '<w:' + x + ' w:val="single" w:sz="4" w:space="0" w:color="auto"/>'; }).join('') +
      '</w:tblBorders></w:tblPr></w:style></w:styles>';
    return s;
  }

  function numberingXml(nums) {
    const bullets = ['●', '○', '■'];
    const ordFmt = ['decimal', 'lowerLetter', 'lowerRoman'];
    const STYLE_FMT = { 'lower-alpha': 'lowerLetter', 'lower-latin': 'lowerLetter', a: 'lowerLetter', 'upper-alpha': 'upperLetter', 'upper-latin': 'upperLetter', A: 'upperLetter', 'lower-roman': 'lowerRoman', i: 'lowerRoman', 'upper-roman': 'upperRoman', I: 'upperRoman', hangul: 'ganada', 'hangul-consonant': 'chosung' };
    let x = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:numbering xmlns:w="' + W_NS + '">';
    nums.forEach(function (n) {
      x += '<w:abstractNum w:abstractNumId="' + n.id + '"><w:multiLevelType w:val="hybridMultilevel"/>';
      for (let l = 0; l < 9; l++) {
        const left = 800 + l * 400;
        if (!n.ordered) {
          x += '<w:lvl w:ilvl="' + l + '"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="' + bullets[l % 3] + '"/><w:lvlJc w:val="left"/>' +
            '<w:pPr><w:ind w:left="' + left + '" w:hanging="400"/></w:pPr><w:rPr><w:rFonts w:ascii="맑은 고딕" w:hAnsi="맑은 고딕" w:hint="eastAsia"/><w:sz w:val="14"/></w:rPr></w:lvl>';
        } else {
          const fmt = l === 0 && STYLE_FMT[n.style] ? STYLE_FMT[n.style] : ordFmt[l % 3];
          x += '<w:lvl w:ilvl="' + l + '"><w:start w:val="' + (l === 0 ? n.start : 1) + '"/><w:numFmt w:val="' + fmt + '"/><w:lvlText w:val="%' + (l + 1) + '."/><w:lvlJc w:val="left"/>' +
            '<w:pPr><w:ind w:left="' + left + '" w:hanging="400"/></w:pPr></w:lvl>';
        }
      }
      x += '</w:abstractNum>';
    });
    nums.forEach(function (n) { x += '<w:num w:numId="' + n.id + '"><w:abstractNumId w:val="' + n.id + '"/></w:num>'; });
    return x + '</w:numbering>';
  }

  DV.word = {
    docxToHtml: docxToHtml,
    htmlToDocx: function (root, opts) { return blocksToDocx(DV.model.fromDOM(root), opts); },
    blocksToDocx: blocksToDocx
  };
})(window.DV);
