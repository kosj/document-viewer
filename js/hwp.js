/*
 * 한글 HWP 5.0 (바이너리) 읽기 → HTML
 * 복합 문서(CFB) 구조는 SheetJS에 포함된 CFB 모듈로 읽고,
 * 레코드 구조는 한글과컴퓨터의 공개 문서(HWP 5.0 파일 구조)를 따릅니다.
 */
window.DV = window.DV || {};
(function (DV) {
  'use strict';

  const TAG = {
    FACE_NAME: 19, BORDER_FILL: 20, CHAR_SHAPE: 21, PARA_SHAPE: 25, BIN_DATA: 18, ID_MAPPINGS: 17,
    PARA_HEADER: 66, PARA_TEXT: 67, PARA_CHAR_SHAPE: 68, CTRL_HEADER: 71, LIST_HEADER: 72,
    TABLE: 77, SHAPE_COMPONENT: 76, SHAPE_COMPONENT_PICTURE: 85, EQEDIT: 88
  };
  const HWPUNIT_PER_PX = 75;

  function ctrlIdStr(n) {
    return String.fromCharCode((n >>> 24) & 255, (n >>> 16) & 255, (n >>> 8) & 255, n & 255);
  }

  function parseRecords(bytes) {
    const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    const recs = [];
    let off = 0;
    while (off + 4 <= bytes.length) {
      const h = dv.getUint32(off, true);
      off += 4;
      const tag = h & 0x3FF;
      const level = (h >>> 10) & 0x3FF;
      let size = (h >>> 20) & 0xFFF;
      if (size === 0xFFF) { size = dv.getUint32(off, true); off += 4; }
      if (off + size > bytes.length) size = bytes.length - off;
      recs.push({ tag: tag, level: level, data: bytes.subarray(off, off + size), children: [] });
      off += size;
    }
    return recs;
  }

  function buildTree(recs) {
    const root = { tag: -1, level: -1, children: [] };
    const stack = [root];
    recs.forEach(function (r) {
      while (stack.length > 1 && stack[stack.length - 1].level >= r.level) stack.pop();
      stack[stack.length - 1].children.push(r);
      stack.push(r);
    });
    return root;
  }

  function u16(d, o) { return o + 2 <= d.length ? d[o] | (d[o + 1] << 8) : 0; }
  function u32(d, o) { return o + 4 <= d.length ? (d[o] | (d[o + 1] << 8) | (d[o + 2] << 16) | (d[o + 3] << 24)) >>> 0 : 0; }
  function i32(d, o) { return o + 4 <= d.length ? (d[o] | (d[o + 1] << 8) | (d[o + 2] << 16) | (d[o + 3] << 24)) : 0; }
  function colorref(n) {
    const r = n & 255, g = (n >>> 8) & 255, b = (n >>> 16) & 255;
    return [r, g, b].map(function (x) { return ('0' + x.toString(16)).slice(-2); }).join('').toUpperCase();
  }
  function wstr(d, o) {
    const len = u16(d, o);
    let s = '';
    for (let i = 0; i < len; i++) s += String.fromCharCode(u16(d, o + 2 + i * 2));
    return { s: s, end: o + 2 + len * 2 };
  }

  function findEntry(cfb, path) {
    const want = ('Root Entry/' + path).toLowerCase();
    for (let i = 0; i < cfb.FullPaths.length; i++) {
      if (cfb.FullPaths[i].toLowerCase().replace(/\/$/, '') === want) return cfb.FileIndex[i];
    }
    return null;
  }

  async function hwpToHtml(buf) {
    const bytes = new Uint8Array(buf);
    if (bytes[0] !== 0xD0 || bytes[1] !== 0xCF) {
      const sig = new TextDecoder('ascii').decode(bytes.subarray(0, 17));
      if (/^HWP Document File/.test(sig)) throw new Error('한글 97 이전(HWP 3.0) 형식은 지원하지 않습니다. 한글에서 최신 형식으로 저장한 뒤 열어 주세요.');
      throw new Error('올바른 HWP 파일이 아닙니다.');
    }
    const cfb = XLSX.CFB.read(bytes, { type: 'array' });
    const fh = findEntry(cfb, 'FileHeader');
    if (!fh) throw new Error('HWP 파일 헤더를 찾을 수 없습니다.');
    const fhd = fh.content instanceof Uint8Array ? fh.content : new Uint8Array(fh.content);
    const sig = new TextDecoder('ascii').decode(fhd.subarray(0, 17));
    if (sig !== 'HWP Document File') throw new Error('올바른 HWP 5.0 파일이 아닙니다.');
    const flags = u32(fhd, 36);
    const compressed = !!(flags & 1);
    if (flags & 2) throw new Error('암호가 걸린 HWP 문서는 열 수 없습니다. 한글에서 암호를 해제한 뒤 다시 시도해 주세요.');
    if (flags & 4) throw new Error('배포용 문서(편집·복사 제한)는 열 수 없습니다.');

    async function stream(path) {
      const e = findEntry(cfb, path);
      if (!e) return null;
      const raw = e.content instanceof Uint8Array ? e.content : new Uint8Array(e.content);
      return compressed ? await DV.inflateRaw(raw) : raw;
    }

    // ---------- DocInfo ----------
    const info = { chars: [], paras: [], fontsHangul: [], bins: {}, borderFills: [] };
    const di = await stream('DocInfo');
    if (di) {
      const recs = parseRecords(di);
      let hangulFontCount = Infinity, faceIdx = 0;
      let binSeq = 0;
      recs.forEach(function (r) {
        const d = r.data;
        switch (r.tag) {
          case TAG.ID_MAPPINGS: hangulFontCount = i32(d, 4); break;
          case TAG.FACE_NAME:
            if (faceIdx < hangulFontCount) info.fontsHangul.push(wstr(d, 1).s);
            faceIdx++;
            break;
          case TAG.CHAR_SHAPE: {
            const size = i32(d, 42) / 100;
            const prop = u32(d, 46);
            const fmt = { size: size };
            if (prop & 1) fmt.i = true;
            if (prop & 2) fmt.b = true;
            if (((prop >>> 2) & 3) === 1) fmt.u = true;
            if ((prop >>> 18) & 7) fmt.s = true;
            if (prop & (1 << 15)) fmt.sup = true;
            if (prop & (1 << 16)) fmt.sub = true;
            const tc = colorref(u32(d, 52));
            if (tc !== '000000') fmt.color = tc;
            const sc = u32(d, 60);
            if (d.length >= 64 && sc !== 0xFFFFFFFF && (sc & 0xFFFFFF) !== 0xFFFFFF) fmt.bg = colorref(sc);
            info.chars.push(fmt);
            break;
          }
          case TAG.PARA_SHAPE: {
            const p1 = u32(d, 0);
            const al = (p1 >>> 2) & 7;
            info.paras.push({
              align: ['justify', 'left', 'right', 'center', 'justify', 'justify'][al] || null,
              leftPt: i32(d, 4) / 200,
              indentPt: i32(d, 12) / 200
            });
            break;
          }
          case TAG.BORDER_FILL: {
            // 채우기: 테두리(6+선 4개*6+대각선 6) 뒤에 채우기 정보
            let o = 2 + 6 * 4 + 6;
            const fillType = u32(d, o);
            let bg = null;
            if (fillType & 1) { const c = u32(d, o + 4); if (c !== 0xFFFFFFFF) bg = colorref(c); }
            info.borderFills.push({ bg: bg });
            break;
          }
          case TAG.BIN_DATA: {
            binSeq++;
            const attr = u16(d, 0);
            const type = attr & 0xF;
            if (type === 1 || type === 2) {
              const id = u16(d, 2);
              const ext = wstr(d, 4).s;
              info.bins[binSeq] = { id: id, ext: ext.toLowerCase(), compress: (attr >>> 4) & 3 };
            }
            break;
          }
        }
      });
    }
    const basePt = (info.chars[0] && info.chars[0].size) || 10;

    // 그림 데이터
    const images = {};
    for (const k of Object.keys(info.bins)) {
      const b = info.bins[k];
      const name = 'BIN' + ('0000' + b.id.toString(16).toUpperCase()).slice(-4) + '.' + b.ext;
      const e = findEntry(cfb, 'BinData/' + name);
      if (!e) continue;
      const mime = DV.mimeFromExt(b.ext);
      if (!DV.isWebImage(mime)) continue;
      let data = e.content instanceof Uint8Array ? e.content : new Uint8Array(e.content);
      const needInflate = b.compress === 1 || (b.compress === 0 && compressed);
      if (needInflate) {
        try { data = await DV.inflateRaw(data); } catch (err) { /* 압축되지 않은 데이터 */ }
      }
      images[k] = DV.bytesToDataUrl(data, mime);
    }

    const ctx = { info: info, images: images, basePt: basePt, skippedObj: 0 };
    let html = '';
    for (let s = 0; s < 1000; s++) {
      const sec = await stream('BodyText/Section' + s);
      if (!sec) break;
      const tree = buildTree(parseRecords(sec));
      if (s > 0) html += '<div class="page-break" contenteditable="false"></div>';
      html += paragraphs(tree.children, ctx);
    }
    if (!html) throw new Error('HWP 본문을 읽을 수 없습니다.');
    const warnings = [];
    if (ctx.skippedObj) warnings.push('일부 개체(도형·수식·OLE 등)는 표시하지 않았습니다.');
    return { html: html, defaultPt: basePt, warnings: warnings };
  }

  function paragraphs(nodes, ctx) {
    let out = '';
    for (let i = 0; i < nodes.length; i++) {
      const n = nodes[i];
      if (n.tag === TAG.PARA_HEADER) out += paragraph(n, ctx);
    }
    return out;
  }

  /* LIST_HEADER 뒤의 문단들을 모읍니다 */
  function listParagraphs(children, startIndex) {
    const lh = children[startIndex];
    const d = lh.data;
    const count = d.length === 30 ? u16(d, 0) : u16(d, 0);
    const paras = [];
    let j = startIndex + 1;
    while (j < children.length && paras.length < count) {
      if (children[j].tag === TAG.PARA_HEADER) paras.push(children[j]);
      else if (children[j].tag === TAG.LIST_HEADER) break;
      j++;
    }
    return { paras: paras, next: j };
  }

  function paragraph(node, ctx) {
    const hd = node.data;
    const paraShape = ctx.info.paras[u16(hd, 8)];
    const breakType = hd[11] || 0;
    let textRec = null, shapeRec = null;
    const ctrls = [];
    node.children.forEach(function (c) {
      if (c.tag === TAG.PARA_TEXT) textRec = c;
      else if (c.tag === TAG.PARA_CHAR_SHAPE) shapeRec = c;
      else if (c.tag === TAG.CTRL_HEADER) ctrls.push(c);
    });
    const css = [];
    if (paraShape) {
      if (paraShape.align) css.push('text-align:' + paraShape.align);
      if (paraShape.leftPt > 0) css.push('margin-left:' + Math.round(paraShape.leftPt) + 'pt');
      if (paraShape.indentPt > 0) css.push('text-indent:' + Math.round(paraShape.indentPt) + 'pt');
    }
    const pw = new DV.hwpCommon.ParaWriter('p', css.join(';'));
    const rb = new DV.hwpCommon.RunBuffer(ctx.basePt);
    let produced = false;
    let prefix = (breakType & 4) ? '<div class="page-break" contenteditable="false"></div>' : '';

    // 글자 모양 포인터
    const ptrs = [];
    if (shapeRec) {
      for (let o = 0; o + 8 <= shapeRec.data.length; o += 8) ptrs.push([u32(shapeRec.data, o), u32(shapeRec.data, o + 4)]);
    }
    let pi = 0;
    function fmtAt(pos) {
      while (pi + 1 < ptrs.length && ptrs[pi + 1][0] <= pos) pi++;
      const id = ptrs.length ? ptrs[pi][1] : 0;
      const f = Object.assign({}, ctx.info.chars[id] || {});
      if (f.size && Math.abs(f.size - ctx.basePt) < 0.01) delete f.size;
      return f;
    }

    let ci = 0;
    if (textRec) {
      const d = textRec.data;
      let pos = 0;
      let o = 0;
      while (o + 2 <= d.length) {
        const code = u16(d, o);
        if (code < 32) {
          const isExt = (code >= 1 && code <= 3) || code === 11 || code === 12 || (code >= 14 && code <= 18) || (code >= 21 && code <= 23);
          const isInline = (code >= 4 && code <= 9) || code === 19 || code === 20;
          if (isExt || isInline) {
            if (code === 9) rb.add(fmtAt(pos), '\t');
            if (isExt) {
              const ctrl = ctrls[ci++];
              if (ctrl) {
                const cid = ctrlIdStr(u32(ctrl.data, 0));
                if (cid === 'tbl ') {
                  const h = rb.take();
                  if (h) { pw.para(h, true); }
                  pw.out += table(ctrl, ctx);
                  produced = true;
                } else if (cid === 'gso ') {
                  const r = shape(ctrl, ctx);
                  if (r.img) rb.raw(r.img);
                  if (r.block) { const h = rb.take(); if (h) pw.para(h, true); pw.out += r.block; produced = true; }
                  if (!r.img && !r.block) ctx.skippedObj++;
                } else if (cid === 'eqed') {
                  const eq = ctrl.children.filter(function (c) { return c.tag === TAG.EQEDIT; })[0];
                  if (eq) { const s = wstr(eq.data, 4).s; if (s) rb.add(Object.assign({}, fmtAt(pos), { i: true }), DV.esc(s)); }
                }
              }
            }
            o += 16; pos += 8;
            continue;
          }
          if (code === 10) rb.add(fmtAt(pos), '<br>');
          else if (code === 24) rb.add(fmtAt(pos), '-');
          else if (code === 30 || code === 31) rb.add(fmtAt(pos), ' ');
          o += 2; pos += 1;
          continue;
        }
        // 서로게이트 쌍 처리
        let ch = String.fromCharCode(code);
        rb.add(fmtAt(pos), DV.esc(ch));
        o += 2; pos += 1;
      }
    }
    const rest = rb.take();
    if (rest || !produced) pw.para(rest, true);
    return prefix + pw.out;
  }

  function table(ctrl, ctx) {
    const ch = ctrl.children;
    const cells = [];
    let i = 0;
    while (i < ch.length) {
      if (ch[i].tag === TAG.LIST_HEADER) {
        const d = ch[i].data;
        const base = d.length === 30 ? 6 : 8;
        const cell = {
          col: u16(d, base), row: u16(d, base + 2), cs: u16(d, base + 4) || 1, rs: u16(d, base + 6) || 1,
          w: u32(d, base + 8), h: u32(d, base + 12), bf: u16(d, base + 24)
        };
        const lp = listParagraphs(ch, i);
        cell.html = lp.paras.map(function (p) { return paragraph(p, ctx); }).join('') || '<p><br></p>';
        cells.push(cell);
        i = lp.next;
      } else i++;
    }
    const rows = {};
    cells.forEach(function (c) { (rows[c.row] = rows[c.row] || []).push(c); });
    let h = '<table class="dv-table"><tbody>';
    Object.keys(rows).map(Number).sort(function (a, b) { return a - b; }).forEach(function (r) {
      h += '<tr>';
      rows[r].sort(function (a, b) { return a.col - b.col; }).forEach(function (c) {
        const css = [];
        if (c.w) css.push('width:' + Math.round(c.w / HWPUNIT_PER_PX) + 'px');
        const bf = ctx.info.borderFills[c.bf - 1];
        if (bf && bf.bg && bf.bg !== 'FFFFFF') css.push('background-color:#' + bf.bg);
        h += '<td' + (c.cs > 1 ? ' colspan="' + c.cs + '"' : '') + (c.rs > 1 ? ' rowspan="' + c.rs + '"' : '') + (css.length ? ' style="' + css.join(';') + '"' : '') + '>' + c.html + '</td>';
      });
      h += '</tr>';
    });
    return h + '</tbody></table>';
  }

  /* 그리기 개체: 그림 또는 글상자 */
  function shape(ctrl, ctx) {
    const res = { img: null, block: null };
    const wPx = Math.round(u32(ctrl.data, 16) / HWPUNIT_PER_PX);
    const hPx = Math.round(u32(ctrl.data, 20) / HWPUNIT_PER_PX);
    let textHtml = '';
    (function walk(nodes) {
      for (let i = 0; i < nodes.length; i++) {
        const n = nodes[i];
        if (n.tag === TAG.SHAPE_COMPONENT_PICTURE && !res.img) {
          const binId = u16(n.data, 71);
          const src = ctx.images[binId];
          if (src) res.img = '<img src="' + src + '" alt=""' + (wPx && hPx ? ' style="width:' + wPx + 'px;height:' + hPx + 'px"' : '') + '>';
        } else if (n.tag === TAG.LIST_HEADER) {
          const lp = listParagraphs(nodes, i);
          textHtml += lp.paras.map(function (p) { return paragraph(p, ctx); }).join('');
          i = lp.next - 1;
        } else if (n.children && n.children.length) walk(n.children);
      }
    })(ctrl.children);
    if (textHtml) res.block = '<div class="dv-textbox-block">' + textHtml + '</div>';
    return res;
  }

  DV.hwp = { hwpToHtml: hwpToHtml };
})(window.DV);
