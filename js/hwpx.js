/* 한글 HWPX (OWPML) 가져오기 / 내보내기 */
window.DV = window.DV || {};
(function (DV) {
  'use strict';

  const kid = DV.kid, kids = DV.kids, attr = DV.attr, esc = DV.esc;
  const HWPUNIT_PER_PX = 75; // 7200 HWPUNIT = 1 inch = 96 px

  /* ---------------- 공통: HTML 생성 도우미 (HWP 바이너리 변환에서도 사용) ---------------- */
  function spanFor(f, inner, basePt) {
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
    if (f.size && Math.abs(f.size - (basePt || 10)) > 0.01) css.push('font-size:' + f.size + 'pt');
    if (css.length) h = '<span style="' + css.join(';') + '">' + h + '</span>';
    return h;
  }

  /* 같은 서식의 연속 글자를 묶어 HTML로 만드는 버퍼 */
  function RunBuffer(basePt) {
    this.html = '';
    this.key = null;
    this.fmt = null;
    this.buf = '';
    this.basePt = basePt || 10;
  }
  RunBuffer.prototype.add = function (fmt, h) {
    const k = fmt ? JSON.stringify(fmt) : '';
    if (k !== this.key) { this.flush(); this.key = k; this.fmt = fmt || {}; }
    this.buf += h;
  };
  RunBuffer.prototype.raw = function (h) { this.flush(); this.html += h; };
  RunBuffer.prototype.flush = function () {
    if (this.buf) this.html += spanFor(this.fmt, this.buf, this.basePt);
    this.buf = ''; this.key = null;
  };
  RunBuffer.prototype.take = function () { this.flush(); const h = this.html; this.html = ''; return h; };

  /* 문단 하나를 HTML로 출력하는 도우미: 표 같은 블록 개체가 나오면 문단을 나눕니다 */
  function ParaWriter(tag, style) {
    this.tag = tag || 'p';
    this.style = style || '';
    this.out = '';
    this.pendingEmpty = true;
  }
  ParaWriter.prototype.para = function (inner, force) {
    if (inner || force) {
      this.out += '<' + this.tag + (this.style ? ' style="' + this.style + '"' : '') + '>' + (inner || '<br>') + '</' + this.tag + '>';
    }
  };

  DV.hwpCommon = { spanFor: spanFor, RunBuffer: RunBuffer, ParaWriter: ParaWriter };

  /* ---------------- HWPX → HTML ---------------- */
  async function hwpxToHtml(buf) {
    const zip = await JSZip.loadAsync(buf);
    const mt = zip.file('mimetype');
    if (mt) {
      const m = (await mt.async('string')).trim();
      if (m && m !== 'application/hwp+zip') throw new Error('HWPX 파일이 아닙니다 (mimetype: ' + m + ').');
    }
    // 매니페스트
    let hpfPath = 'Contents/content.hpf';
    const container = zip.file('META-INF/container.xml');
    if (container) {
      const cdoc = DV.parseXml(await container.async('string'));
      DV.desc(cdoc, 'rootfile').forEach(function (r) {
        const p = r.getAttribute('full-path');
        if (/\.hpf$/i.test(p || '')) hpfPath = p;
      });
    }
    const items = {};
    const sections = [];
    const hpf = zip.file(hpfPath);
    if (hpf) {
      const hdoc = DV.parseXml(await hpf.async('string'));
      DV.desc(hdoc, 'item').forEach(function (it) {
        items[it.getAttribute('id')] = { href: it.getAttribute('href'), type: it.getAttribute('media-type') };
      });
      DV.desc(hdoc, 'itemref').forEach(function (ir) {
        const it = items[ir.getAttribute('idref')];
        if (it && /section/i.test(it.href)) sections.push(it.href);
      });
    }
    if (!sections.length) {
      Object.keys(zip.files).forEach(function (n) { if (/^Contents\/section\d+\.xml$/i.test(n)) sections.push(n); });
      sections.sort(function (a, b) { return parseInt(a.replace(/\D/g, ''), 10) - parseInt(b.replace(/\D/g, ''), 10); });
    }
    if (!sections.length) throw new Error('HWPX 본문(section)을 찾을 수 없습니다.');

    const head = await parseHeader(zip);
    const ctx = { head: head, items: items, zip: zip, images: {}, warnings: [] };
    // 그림 미리 읽기
    await Promise.all(Object.keys(items).map(async function (id) {
      const it = items[id];
      if (!it.href || !/^BinData\//i.test(it.href)) return;
      const f = zip.file(it.href);
      if (!f) return;
      const mime = it.type && /^image\//.test(it.type) ? it.type : DV.mimeFromExt(DV.ext(it.href));
      if (!DV.isWebImage(mime)) return;
      ctx.images[id] = DV.bytesToDataUrl(await f.async('uint8array'), mime);
    }));

    let html = '';
    for (let i = 0; i < sections.length; i++) {
      const f = zip.file(sections[i]);
      if (!f) continue;
      const sdoc = DV.parseXml(await f.async('string'));
      if (i > 0) html += '<div class="page-break" contenteditable="false"></div>';
      html += paragraphsToHtml(kids(sdoc.documentElement, 'p'), ctx);
    }
    return { html: html || '<p><br></p>', defaultPt: head.basePt, warnings: ctx.warnings };
  }

  async function parseHeader(zip) {
    const res = { chars: {}, paras: {}, borderFills: {}, fonts: {}, basePt: 10, styles: {} };
    const f = zip.file('Contents/header.xml');
    if (!f) return res;
    const doc = DV.parseXml(await f.async('string'));
    // 글꼴 (한글)
    DV.desc(doc, 'fontface').forEach(function (ff) {
      if (attr(ff, 'lang') !== 'HANGUL') return;
      kids(ff, 'font').forEach(function (fo) { res.fonts[attr(fo, 'id')] = attr(fo, 'face'); });
    });
    DV.desc(doc, 'borderFill').forEach(function (bf) {
      const wb = DV.desc(bf, 'winBrush')[0];
      const face = wb ? attr(wb, 'faceColor') : null;
      res.borderFills[attr(bf, 'id')] = { bg: face && face !== 'none' ? DV.colorHex(face) : null };
    });
    DV.desc(doc, 'charPr').forEach(function (cp) {
      const fmt = {};
      const h = parseInt(attr(cp, 'height') || '1000', 10);
      fmt.size = Math.round(h / 50) / 2;
      const tc = DV.colorHex(attr(cp, 'textColor'));
      if (tc && tc !== '000000') fmt.color = tc;
      const sc = attr(cp, 'shadeColor');
      if (sc && sc !== 'none') { const s = DV.colorHex(sc); if (s && s !== 'FFFFFF') fmt.bg = s; }
      if (kid(cp, 'bold')) fmt.b = true;
      if (kid(cp, 'italic')) fmt.i = true;
      const ul = kid(cp, 'underline');
      if (ul && attr(ul, 'type') && attr(ul, 'type') !== 'NONE') fmt.u = true;
      const so = kid(cp, 'strikeout');
      if (so && attr(so, 'shape') && attr(so, 'shape') !== 'NONE' && attr(so, 'shape') !== '3D') fmt.s = true;
      if (kid(cp, 'supscript')) fmt.sup = true;
      if (kid(cp, 'subscript')) fmt.sub = true;
      res.chars[attr(cp, 'id')] = fmt;
    });
    DV.desc(doc, 'paraPr').forEach(function (pp) {
      const al = kid(pp, 'align');
      const h = al ? attr(al, 'horizontal') : 'JUSTIFY';
      const left = DV.desc(pp, 'left')[0];
      const intent = DV.desc(pp, 'intent')[0];
      res.paras[attr(pp, 'id')] = {
        align: { LEFT: 'left', CENTER: 'center', RIGHT: 'right', JUSTIFY: 'justify', DISTRIBUTE: 'justify', DISTRIBUTE_SPACE: 'justify' }[h] || null,
        leftPt: left ? parseInt(attr(left, 'value') || '0', 10) / 100 : 0,
        indentPt: intent ? parseInt(attr(intent, 'value') || '0', 10) / 100 : 0
      };
    });
    // 기본 글자 크기: 바탕글 스타일의 글자 모양
    const st0 = DV.desc(doc, 'style').filter(function (s) { return attr(s, 'id') === '0'; })[0];
    const cp0 = st0 ? res.chars[attr(st0, 'charPrIDRef')] : res.chars['0'];
    if (cp0 && cp0.size) res.basePt = cp0.size;
    DV.desc(doc, 'style').forEach(function (s) {
      res.styles[attr(s, 'id')] = { name: attr(s, 'name') || '', engName: attr(s, 'engName') || '' };
    });
    return res;
  }

  function paraStyle(pp) {
    const css = [];
    if (!pp) return '';
    if (pp.align && pp.align !== 'justify') css.push('text-align:' + pp.align);
    if (pp.align === 'justify') css.push('text-align:justify');
    if (pp.leftPt > 0) css.push('margin-left:' + Math.round(pp.leftPt) + 'pt');
    if (pp.indentPt > 0) css.push('text-indent:' + Math.round(pp.indentPt) + 'pt');
    return css.join(';');
  }

  function paragraphsToHtml(ps, ctx) {
    let out = '';
    ps.forEach(function (p) {
      if (attr(p, 'pageBreak') === '1' && out) out += '<div class="page-break" contenteditable="false"></div>';
      out += paragraphToHtml(p, ctx);
    });
    return out;
  }

  function charFmt(ctx, id) {
    const f = Object.assign({}, ctx.head.chars[id] || {});
    if (f.size && Math.abs(f.size - ctx.head.basePt) < 0.01) delete f.size;
    return f;
  }

  function paragraphToHtml(p, ctx) {
    const pp = ctx.head.paras[attr(p, 'paraPrIDRef')];
    const st = ctx.head.styles[attr(p, 'styleIDRef')];
    let tag = 'p';
    if (st) {
      const m = /^Outline (\d+)$/.exec(st.engName) || /^개요 (\d+)$/.exec(st.name);
      void m; // 개요 스타일은 일반 문단으로 유지 (번호가 글자로 들어가는 경우가 많음)
    }
    const pw = new ParaWriter(tag, paraStyle(pp));
    const rb = new RunBuffer(ctx.head.basePt);
    let produced = false;
    let hadText = false;

    function flushPara(force) {
      const h = rb.take();
      if (h || force) { pw.para(h, true); produced = true; }
    }

    kids(p, 'run').forEach(function (run) {
      const fmt = charFmt(ctx, attr(run, 'charPrIDRef'));
      kids(run).forEach(function (c) {
        switch (c.localName) {
          case 't':
            for (let n = c.firstChild; n; n = n.nextSibling) {
              if (n.nodeType === 3) { if (n.nodeValue) { rb.add(fmt, esc(n.nodeValue)); hadText = true; } }
              else if (n.nodeType === 1) {
                const ln = n.localName;
                if (ln === 'tab') rb.add(fmt, '\t');
                else if (ln === 'lineBreak') rb.add(fmt, '<br>');
                else if (ln === 'nbSpace' || ln === 'fwSpace') rb.add(fmt, ' ');
                else if (ln === 'hyphen') rb.add(fmt, '-');
                else if (n.textContent) rb.add(fmt, esc(n.textContent));
              }
            }
            break;
          case 'tbl': {
            flushPara(false);
            pw.out += tableToHtml(c, ctx);
            produced = true;
            break;
          }
          case 'pic': {
            const img = DV.desc(c, 'img')[0];
            const ref = img ? attr(img, 'binaryItemIDRef') : null;
            const src = ref ? ctx.images[ref] : null;
            if (src) {
              const sz = kid(c, 'curSz') && +attr(kid(c, 'curSz'), 'width') ? kid(c, 'curSz') : kid(c, 'sz');
              const w = sz ? Math.round(parseInt(attr(sz, 'width'), 10) / HWPUNIT_PER_PX) : 0;
              const h = sz ? Math.round(parseInt(attr(sz, 'height'), 10) / HWPUNIT_PER_PX) : 0;
              rb.raw('<img src="' + src + '" alt=""' + (w && h ? ' style="width:' + w + 'px;height:' + h + 'px"' : '') + '>');
              hadText = true;
            }
            break;
          }
          case 'rect': case 'ellipse': case 'polygon': case 'container': case 'curve': case 'arc': {
            // 글상자: 안의 문단을 꺼내 표시
            const lists = DV.desc(c, 'drawText');
            if (lists.length) {
              flushPara(false);
              lists.forEach(function (dt) {
                const sl = kid(dt, 'subList');
                if (sl) pw.out += '<div class="dv-textbox-block">' + paragraphsToHtml(kids(sl, 'p'), ctx) + '</div>';
              });
              produced = true;
            }
            break;
          }
          case 'equation': {
            const sc = kid(c, 'script');
            if (sc && sc.textContent) rb.add(Object.assign({}, fmt, { i: true }), esc(sc.textContent));
            break;
          }
          default: break;
        }
      });
    });
    const rest = rb.take();
    if (rest || !produced) pw.para(rest, true);
    void hadText;
    return pw.out;
  }

  function tableToHtml(tbl, ctx) {
    let h = '<table class="dv-table"><tbody>';
    kids(tbl, 'tr').forEach(function (tr) {
      h += '<tr>';
      kids(tr, 'tc').forEach(function (tc) {
        const span = kid(tc, 'cellSpan');
        const cs = span ? parseInt(attr(span, 'colSpan') || '1', 10) : 1;
        const rs = span ? parseInt(attr(span, 'rowSpan') || '1', 10) : 1;
        const sz = kid(tc, 'cellSz');
        const css = [];
        const w = sz ? Math.round(parseInt(attr(sz, 'width') || '0', 10) / HWPUNIT_PER_PX) : 0;
        if (w) css.push('width:' + w + 'px');
        const bf = ctx.head.borderFills[attr(tc, 'borderFillIDRef')];
        if (bf && bf.bg && bf.bg !== 'FFFFFF') css.push('background-color:#' + bf.bg);
        const sl = kid(tc, 'subList');
        if (sl) {
          const va = attr(sl, 'vertAlign');
          if (va === 'CENTER') css.push('vertical-align:middle');
          else if (va === 'BOTTOM') css.push('vertical-align:bottom');
        }
        const inner = sl ? paragraphsToHtml(kids(sl, 'p'), ctx) : '<p><br></p>';
        h += '<td' + (cs > 1 ? ' colspan="' + cs + '"' : '') + (rs > 1 ? ' rowspan="' + rs + '"' : '') + (css.length ? ' style="' + css.join(';') + '"' : '') + '>' + inner + '</td>';
      });
      h += '</tr>';
    });
    return h + '</tbody></table>';
  }

  /* ---------------- 모델 → HWPX ---------------- */
  const T = function () { return DV.HWPX_TEMPLATE; };
  const TEXT_WIDTH = 42520; // A4 - 좌우 여백 (HWPUNIT)

  async function blocksToHwpx(blocks, opts) {
    opts = opts || {};
    const tpl = T();
    let header = tpl['Contents/header.xml'];
    const charTpl = /<hh:charPr id="0"[\s\S]*?<\/hh:charPr>/.exec(header)[0];
    const paraTpl = /<hh:paraPr id="0"[\s\S]*?<\/hh:paraPr>/.exec(header)[0];
    const charCount = +/<hh:charProperties itemCnt="(\d+)"/.exec(header)[1];
    const paraCount = +/<hh:paraProperties itemCnt="(\d+)"/.exec(header)[1];
    const bfCount = +/<hh:borderFills itemCnt="(\d+)"/.exec(header)[1];

    const chars = [], charMap = {};
    const paras = [], paraMap = {};
    const fills = [], fillMap = {};
    const bins = [];
    const HEADING_PT = { 1: 20, 2: 16, 3: 14, 4: 12, 5: 11, 6: 10 };

    // 표 테두리용 borderFill (실선)
    function borderFillId(bg, border) {
      const key = (bg || '') + '|' + (border ? 1 : 0);
      if (fillMap[key] != null) return fillMap[key];
      const id = bfCount + fills.length + 1;
      const bt = border ? 'SOLID' : 'NONE';
      const bw = border ? '0.12 mm' : '0.1 mm';
      let x = '<hh:borderFill id="' + id + '" threeD="0" shadow="0" centerLine="NONE" breakCellSeparateLine="0"><hh:slash type="NONE" Crooked="0" isCounter="0"/><hh:backSlash type="NONE" Crooked="0" isCounter="0"/>' +
        ['left', 'right', 'top', 'bottom'].map(function (s) { return '<hh:' + s + 'Border type="' + bt + '" width="' + bw + '" color="#000000"/>'; }).join('') +
        '<hh:diagonal type="SOLID" width="0.1 mm" color="#000000"/>';
      if (bg) x += '<hc:fillBrush><hc:winBrush faceColor="#' + bg + '" hatchColor="#999999" alpha="0"/></hc:fillBrush>';
      x += '</hh:borderFill>';
      fills.push(x);
      fillMap[key] = id;
      return id;
    }

    function charId(r, heading) {
      const f = {
        b: !!(r.b || heading), i: !!r.i, u: !!r.u, s: !!r.s, sup: !!r.sup, sub: !!r.sub,
        color: r.color || null, bg: r.bg || null,
        size: r.size || (heading ? HEADING_PT[heading] : null), font: fontRefFor(r.font)
      };
      if (!f.b && !f.i && !f.u && !f.s && !f.sup && !f.sub && !f.color && !f.bg && (!f.size || f.size === 10) && f.font == null) return 0;
      const key = JSON.stringify(f);
      if (charMap[key] != null) return charMap[key];
      const id = charCount + chars.length;
      let x = charTpl.replace('id="0"', 'id="' + id + '"')
        .replace(/height="\d+"/, 'height="' + Math.round((f.size || 10) * 100) + '"')
        .replace(/textColor="[^"]*"/, 'textColor="#' + (f.color || '000000') + '"')
        .replace(/shadeColor="[^"]*"/, 'shadeColor="' + (f.bg ? '#' + f.bg : 'none') + '"');
      if (f.font != null) x = x.replace(/<hh:fontRef [^>]*\/>/, '<hh:fontRef hangul="' + f.font + '" latin="' + f.font + '" hanja="' + f.font + '" japanese="' + f.font + '" other="' + f.font + '" symbol="' + f.font + '" user="' + f.font + '"/>');
      const deco = (f.i ? '<hh:italic/>' : '') + (f.b ? '<hh:bold/>' : '');
      x = x.replace(/<hh:underline [^>]*\/>/, deco + '<hh:underline type="' + (f.u ? 'BOTTOM' : 'NONE') + '" shape="SOLID" color="' + (f.u && f.color ? '#' + f.color : '#000000') + '"/>');
      x = x.replace(/<hh:strikeout [^>]*\/>/, '<hh:strikeout shape="' + (f.s ? 'SOLID' : 'NONE') + '" color="#000000"/>');
      if (f.sup || f.sub) x = x.replace('</hh:charPr>', (f.sup ? '<hh:supscript/>' : '<hh:subscript/>') + '</hh:charPr>');
      chars.push(x);
      charMap[key] = id;
      return id;
    }

    function fontRefFor(font) {
      if (!font) return null;
      // 템플릿 글꼴: 0 = 함초롬돋움, 1 = 함초롬바탕
      if (/바탕|명조|batang|myeongjo|times|serif|궁서|georgia/i.test(font) && !/sans/i.test(font)) return 1;
      return 0;
    }

    function paraId(p) {
      const align = p.align ? ({ left: 'LEFT', center: 'CENTER', right: 'RIGHT', justify: 'JUSTIFY' })[p.align] : null;
      const left = Math.round((p.indentPt || 0) * 100);
      const intent = Math.round((p.firstLinePt || 0) * 100);
      if (!align && !left && !intent) return 0;
      const key = (align || 'JUSTIFY') + '|' + left + '|' + intent;
      if (paraMap[key] != null) return paraMap[key];
      const id = paraCount + paras.length;
      let x = paraTpl.replace('id="0"', 'id="' + id + '"').replace(/horizontal="[A-Z_]+"/, 'horizontal="' + (align || 'JUSTIFY') + '"');
      // case(실제 값) / default(2배 값)
      x = x.replace(/(<hp:case[\s\S]*?<\/hp:case>)/, function (m) {
        return m.replace(/<hc:left value="-?\d+"/, '<hc:left value="' + left + '"').replace(/<hc:intent value="-?\d+"/, '<hc:intent value="' + intent + '"');
      }).replace(/(<hp:default>[\s\S]*?<\/hp:default>)/, function (m) {
        return m.replace(/<hc:left value="-?\d+"/, '<hc:left value="' + left * 2 + '"').replace(/<hc:intent value="-?\d+"/, '<hc:intent value="' + intent * 2 + '"');
      });
      paras.push(x);
      paraMap[key] = id;
      return id;
    }

    async function imageXml(img) {
      let src = img.src;
      let info = /^data:/.test(src) ? DV.dataUrlInfo(src) : null;
      if (info && !/^image\/(png|jpeg|gif|bmp)$/.test(info.mime)) {
        const png = await DV.ensurePng(src);
        info = png ? DV.dataUrlInfo(png) : null;
      }
      if (!info) return '';
      const n = bins.length + 1;
      const ext = info.ext === 'jpeg' ? 'jpg' : info.ext;
      const id = 'image' + n;
      bins.push({ id: id, href: 'BinData/' + id + '.' + ext, mime: info.mime, bytes: info.bytes });
      const w = Math.round(img.w * HWPUNIT_PER_PX), h = Math.round(img.h * HWPUNIT_PER_PX);
      const inst = 1000000 + n;
      return '<hp:pic id="' + inst + '" zOrder="' + n + '" numberingType="PICTURE" textWrap="TOP_AND_BOTTOM" textFlow="BOTH_SIDES" lock="0" dropcapstyle="None" href="" groupLevel="0" instid="' + inst + '" reverse="0">' +
        '<hp:offset x="0" y="0"/><hp:orgSz width="' + w + '" height="' + h + '"/><hp:curSz width="' + w + '" height="' + h + '"/><hp:flip horizontal="0" vertical="0"/>' +
        '<hp:rotationInfo angle="0" centerX="' + Math.round(w / 2) + '" centerY="' + Math.round(h / 2) + '" rotateimage="1"/>' +
        '<hp:renderingInfo><hc:transMatrix e1="1" e2="0" e3="0" e4="0" e5="1" e6="0"/><hc:scaMatrix e1="1" e2="0" e3="0" e4="0" e5="1" e6="0"/><hc:rotMatrix e1="1" e2="0" e3="0" e4="0" e5="1" e6="0"/></hp:renderingInfo>' +
        '<hc:img binaryItemIDRef="' + id + '" bright="0" contrast="0" effect="REAL_PIC" alpha="0"/>' +
        '<hp:imgRect><hc:pt0 x="0" y="0"/><hc:pt1 x="' + w + '" y="0"/><hc:pt2 x="' + w + '" y="' + h + '"/><hc:pt3 x="0" y="' + h + '"/></hp:imgRect>' +
        '<hp:imgClip left="0" right="' + w + '" top="0" bottom="' + h + '"/><hp:inMargin left="0" right="0" top="0" bottom="0"/><hp:imgDim dimwidth="' + w + '" dimheight="' + h + '"/><hp:effects/>' +
        '<hp:sz width="' + w + '" widthRelTo="ABSOLUTE" height="' + h + '" heightRelTo="ABSOLUTE" protect="0"/>' +
        '<hp:pos treatAsChar="1" affectLSpacing="0" flowWithText="1" allowOverlap="0" holdAnchorAndSO="0" vertRelTo="PARA" horzRelTo="COLUMN" vertAlign="TOP" horzAlign="LEFT" vertOffset="0" horzOffset="0"/>' +
        '<hp:outMargin left="0" right="0" top="0" bottom="0"/></hp:pic>';
    }

    let tblSeq = 0;
    let first = true;

    function pOpen(paraPr, pageBreak) {
      return '<hp:p id="0" paraPrIDRef="' + paraPr + '" styleIDRef="0" pageBreak="' + (pageBreak ? 1 : 0) + '" columnBreak="0" merged="0">';
    }
    function secRun() {
      if (!first) return '';
      first = false;
      return '<hp:run charPrIDRef="0">' + tpl.secPr + '</hp:run>';
    }

    async function paraXml(p) {
      const ppr = paraId(p);
      let runs = '';
      const prefix = p.list ? DV.model.listLabel(p.list) : '';
      const listIndent = p.list ? (p.list.level + 1) * 20 : 0;
      const realPpr = p.list ? paraId(Object.assign({}, p, { indentPt: listIndent, firstLinePt: -10 })) : ppr;
      let curId = null, curText = '';
      function flushRun() {
        if (curId !== null) runs += '<hp:run charPrIDRef="' + curId + '"><hp:t>' + curText + '</hp:t></hp:run>';
        curId = null; curText = '';
      }
      function addText(id, t) {
        if (curId !== null && curId !== id) flushRun();
        curId = id;
        curText += t;
      }
      if (prefix) addText(charId({}, p.heading), DV.xmlEsc(prefix));
      if (p.hr) addText(0, '―'.repeat(30));
      for (const r of p.runs) {
        if (r.img) {
          flushRun();
          const x = await imageXml(r.img);
          if (x) runs += '<hp:run charPrIDRef="0">' + x + '<hp:t/></hp:run>';
          continue;
        }
        if (r.br) { addText(curId == null ? 0 : curId, '<hp:lineBreak/>'); continue; }
        if (r.tab) { addText(charId(r.fmt || {}, p.heading), '<hp:tab width="4000" leader="0" type="1"/>'); continue; }
        if (r.text) addText(charId(r, p.heading), DV.xmlEsc(r.text));
      }
      flushRun();
      if (!runs) runs = '<hp:run charPrIDRef="' + (p.heading ? charId({}, p.heading) : 0) + '"><hp:t/></hp:run>';
      return pOpen(realPpr, p.pageBreakBefore) + secRun() + runs + '</hp:p>';
    }

    async function tableXml(t, availWidth) {
      const totalPx = t.colWidthsPx.reduce(function (a, b) { return a + b; }, 0) || 1;
      let k = HWPUNIT_PER_PX;
      if (totalPx * k > availWidth) k = availWidth / totalPx;
      const colW = t.colWidthsPx.map(function (w) { return Math.max(1000, Math.round(w * k)); });
      const totalW = colW.reduce(function (a, b) { return a + b; }, 0);
      const nrows = t.rows.length;
      const ncols = colW.length;
      // 그리드 배치 (가려진 칸은 HWPX에서 생략)
      const grid = [];
      const placed = [];
      t.rows.forEach(function (row, r) {
        grid[r] = grid[r] || [];
        let c = 0;
        row.forEach(function (cell) {
          while (grid[r][c]) c++;
          const rs = Math.min(cell.rowspan, nrows - r), cs = Math.min(cell.colspan, Math.max(1, ncols - c));
          for (let i = 0; i < rs; i++) { grid[r + i] = grid[r + i] || []; for (let j = 0; j < cs; j++) grid[r + i][c + j] = true; }
          placed.push({ cell: cell, r: r, c: c, rs: rs, cs: cs });
          c += cs;
        });
      });
      const ROW_H = 282 + 282;
      const bfTable = borderFillId(null, t.border);
      const tid = 1900000000 + (++tblSeq);
      let x = '<hp:tbl id="' + tid + '" zOrder="0" numberingType="TABLE" textWrap="TOP_AND_BOTTOM" textFlow="BOTH_SIDES" lock="0" dropcapstyle="None" pageBreak="CELL" repeatHeader="1" rowCnt="' + nrows + '" colCnt="' + ncols + '" cellSpacing="0" borderFillIDRef="' + bfTable + '" noAdjust="0">' +
        '<hp:sz width="' + totalW + '" widthRelTo="ABSOLUTE" height="' + ROW_H * nrows + '" heightRelTo="ABSOLUTE" protect="0"/>' +
        // 행이 많은 표는 쪽에서 나뉘도록 글자처럼 취급하지 않습니다
        (nrows >= 8
          ? '<hp:pos treatAsChar="0" affectLSpacing="0" flowWithText="1" allowOverlap="0" holdAnchorAndSO="0" vertRelTo="PARA" horzRelTo="PARA" vertAlign="TOP" horzAlign="LEFT" vertOffset="0" horzOffset="0"/>'
          : '<hp:pos treatAsChar="1" affectLSpacing="0" flowWithText="1" allowOverlap="0" holdAnchorAndSO="0" vertRelTo="PARA" horzRelTo="COLUMN" vertAlign="TOP" horzAlign="LEFT" vertOffset="0" horzOffset="0"/>') +
        '<hp:outMargin left="283" right="283" top="283" bottom="283"/><hp:inMargin left="510" right="510" top="141" bottom="141"/>';
      for (let r = 0; r < nrows; r++) {
        x += '<hp:tr>';
        for (const pc of placed) {
          if (pc.r !== r) continue;
          let w = 0;
          for (let j = 0; j < pc.cs; j++) w += colW[pc.c + j] || 0;
          const bf = borderFillId(pc.cell.bg, t.border);
          const va = pc.cell.valign === 'bottom' ? 'BOTTOM' : pc.cell.valign === 'top' ? 'TOP' : 'CENTER';
          let inner = '';
          const savedFirst = first; first = false;
          for (const b of pc.cell.blocks) inner += b.type === 'table' ? await tableParaXml(b, w - 1020) : await paraXml(b);
          first = savedFirst;
          if (!inner) inner = pOpen(0) + '<hp:run charPrIDRef="0"><hp:t/></hp:run></hp:p>';
          x += '<hp:tc name="" header="0" hasMargin="0" protect="0" editable="0" dirty="0" borderFillIDRef="' + bf + '">' +
            '<hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="' + va + '" linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">' + inner + '</hp:subList>' +
            '<hp:cellAddr colAddr="' + pc.c + '" rowAddr="' + pc.r + '"/><hp:cellSpan colSpan="' + pc.cs + '" rowSpan="' + pc.rs + '"/>' +
            '<hp:cellSz width="' + w + '" height="' + ROW_H * pc.rs + '"/><hp:cellMargin left="510" right="510" top="141" bottom="141"/></hp:tc>';
        }
        x += '</hp:tr>';
      }
      return x + '</hp:tbl>';
    }

    async function tableParaXml(t, availWidth) {
      const head = pOpen(0) + secRun();
      return head + '<hp:run charPrIDRef="0">' + await tableXml(t, availWidth) + '<hp:t/></hp:run></hp:p>';
    }

    let body = '';
    for (const b of blocks) {
      if (b.type === 'table') body += await tableParaXml(b, TEXT_WIDTH);
      else body += await paraXml(b);
    }
    if (!body) body = pOpen(0) + secRun() + '<hp:run charPrIDRef="0"><hp:t/></hp:run></hp:p>';

    // header.xml 갱신
    if (chars.length) header = header.replace(/<hh:charProperties itemCnt="\d+">/, '<hh:charProperties itemCnt="' + (charCount + chars.length) + '">').replace('</hh:charProperties>', chars.join('') + '</hh:charProperties>');
    if (paras.length) header = header.replace(/<hh:paraProperties itemCnt="\d+">/, '<hh:paraProperties itemCnt="' + (paraCount + paras.length) + '">').replace('</hh:paraProperties>', paras.join('') + '</hh:paraProperties>');
    if (fills.length) header = header.replace(/<hh:borderFills itemCnt="\d+">/, '<hh:borderFills itemCnt="' + (bfCount + fills.length) + '">').replace('</hh:borderFills>', fills.join('') + '</hh:borderFills>');

    const section = '<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>' + tpl.secOpen + body + '</hs:sec>';
    let hpf = '<?xml version="1.0" encoding="UTF-8" standalone="yes" ?><opf:package xmlns:ha="http://www.hancom.co.kr/hwpml/2011/app" xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph" xmlns:hp10="http://www.hancom.co.kr/hwpml/2016/paragraph" xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section" xmlns:hc="http://www.hancom.co.kr/hwpml/2011/core" xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head" xmlns:hhs="http://www.hancom.co.kr/hwpml/2011/history" xmlns:hm="http://www.hancom.co.kr/hwpml/2011/master-page" xmlns:hpf="http://www.hancom.co.kr/schema/2011/hpf" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:opf="http://www.idpf.org/2007/opf/" xmlns:ooxmlchart="http://www.hancom.co.kr/hwpml/2016/ooxmlchart" xmlns:hwpunitchar="http://www.hancom.co.kr/hwpml/2016/HwpUnitChar" xmlns:epub="http://www.idpf.org/2007/ops" xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0" version="" unique-identifier="" id="">';
    const now = new Date();
    const iso = now.toISOString().replace(/\.\d+Z$/, 'Z');
    hpf += '<opf:metadata><opf:title>' + DV.xmlEsc(opts.title || '') + '</opf:title><opf:language>ko</opf:language><opf:meta name="creator" content="text">문서 편집기</opf:meta><opf:meta name="subject" content="text"/><opf:meta name="description" content="text"/><opf:meta name="lastsaveby" content="text">문서 편집기</opf:meta><opf:meta name="CreatedDate" content="text">' + iso + '</opf:meta><opf:meta name="ModifiedDate" content="text">' + iso + '</opf:meta><opf:meta name="keyword" content="text"/></opf:metadata>';
    hpf += '<opf:manifest><opf:item id="header" href="Contents/header.xml" media-type="application/xml"/>';
    bins.forEach(function (b) { hpf += '<opf:item id="' + b.id + '" href="' + b.href + '" media-type="' + b.mime + '" isEmbeded="1"/>'; });
    hpf += '<opf:item id="section0" href="Contents/section0.xml" media-type="application/xml"/><opf:item id="settings" href="settings.xml" media-type="application/xml"/></opf:manifest>' +
      '<opf:spine><opf:itemref idref="header" linear="yes"/><opf:itemref idref="section0" linear="yes"/></opf:spine></opf:package>';

    const zip = new JSZip();
    const zfile = zip.file.bind(zip);
    zip.file = function (n, d, o) { return zfile(n, d, Object.assign({ createFolders: false }, o)); };
    zip.file('mimetype', 'application/hwp+zip', { compression: 'STORE' });
    zip.file('version.xml', tpl['version.xml']);
    zip.file('Contents/header.xml', header);
    zip.file('Contents/section0.xml', section);
    bins.forEach(function (b) { zip.file(b.href, b.bytes); });
    zip.file('Preview/PrvText.txt', previewText(blocks));
    zip.file('settings.xml', tpl['settings.xml']);
    zip.file('META-INF/container.rdf', tpl['META-INF/container.rdf']);
    zip.file('Contents/content.hpf', hpf);
    zip.file('META-INF/container.xml', tpl['META-INF/container.xml']);
    zip.file('META-INF/manifest.xml', tpl['META-INF/manifest.xml']);
    return zip.generateAsync({ type: 'blob', mimeType: 'application/hwp+zip', compression: 'DEFLATE' });
  }

  function previewText(blocks) {
    const lines = [];
    (function walk(bs) {
      bs.forEach(function (b) {
        if (b.type === 'p') lines.push(b.runs.map(function (r) { return r.text || (r.tab ? '\t' : ''); }).join(''));
        else b.rows.forEach(function (row) { row.forEach(function (c) { walk(c.blocks); }); });
      });
    })(blocks);
    return lines.join('\r\n').slice(0, 1024);
  }

  DV.hwpx = {
    hwpxToHtml: hwpxToHtml,
    htmlToHwpx: function (root, opts) { return blocksToHwpx(DV.model.fromDOM(root), opts); },
    blocksToHwpx: blocksToHwpx
  };
})(window.DV);
