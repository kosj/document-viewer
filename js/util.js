/* 공용 유틸리티 */
window.DV = window.DV || {};
(function (DV) {
  'use strict';

  DV.esc = function (s) {
    return String(s == null ? '' : s)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  };

  /* XML 1.0에서 허용되지 않는 제어문자를 제거하고 이스케이프 */
  DV.xmlEsc = function (s) {
    return DV.esc(String(s == null ? '' : s).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]/g, ''));
  };

  DV.ext = function (name) {
    const m = /\.([^.]+)$/.exec(name || '');
    return m ? m[1].toLowerCase() : '';
  };

  DV.baseName = function (name) {
    return (name || '문서').replace(/\.[^.]+$/, '');
  };

  DV.download = function (blob, filename) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(function () { URL.revokeObjectURL(url); }, 4000);
  };

  DV.readFile = function (file) {
    if (file.arrayBuffer) return file.arrayBuffer();
    return new Promise(function (resolve, reject) {
      const fr = new FileReader();
      fr.onload = function () { resolve(fr.result); };
      fr.onerror = function () { reject(fr.error); };
      fr.readAsArrayBuffer(file);
    });
  };

  DV.toast = function (msg, type, ms) {
    const box = document.getElementById('toasts');
    if (!box) return;
    const el = document.createElement('div');
    el.className = 'toast ' + (type || 'info');
    el.textContent = msg;
    box.appendChild(el);
    setTimeout(function () { el.classList.add('hide'); }, ms || 3800);
    setTimeout(function () { el.remove(); }, (ms || 3800) + 400);
  };

  /* ---------- XML 도우미 (네임스페이스 접두사와 무관하게 localName으로 탐색) ---------- */
  DV.parseXml = function (str) {
    const doc = new DOMParser().parseFromString(str, 'application/xml');
    const err = doc.getElementsByTagName('parsererror')[0];
    if (err) throw new Error('XML 구문 오류: ' + err.textContent.slice(0, 120));
    return doc;
  };
  DV.kids = function (el, name) {
    const out = [];
    if (!el) return out;
    for (let n = el.firstElementChild; n; n = n.nextElementSibling) {
      if (!name || n.localName === name) out.push(n);
    }
    return out;
  };
  DV.kid = function (el, name) {
    if (!el) return null;
    for (let n = el.firstElementChild; n; n = n.nextElementSibling) {
      if (n.localName === name) return n;
    }
    return null;
  };
  /* 경로 탐색: DV.path(el, 'pPr', 'jc') */
  DV.path = function (el) {
    let cur = el;
    for (let i = 1; i < arguments.length && cur; i++) cur = DV.kid(cur, arguments[i]);
    return cur;
  };
  DV.desc = function (el, name) {
    if (!el) return [];
    return Array.prototype.slice.call(el.getElementsByTagNameNS('*', name));
  };
  DV.attr = function (el, name) {
    if (!el || !el.attributes) return null;
    for (let i = 0; i < el.attributes.length; i++) {
      const a = el.attributes[i];
      if (a.localName === name) return a.value;
    }
    return null;
  };

  /* ---------- 색상/크기 ---------- */
  const NAMED = {
    black: '000000', white: 'FFFFFF', red: 'FF0000', green: '008000', blue: '0000FF',
    yellow: 'FFFF00', gray: '808080', grey: '808080', orange: 'FFA500', purple: '800080',
    navy: '000080', teal: '008080', maroon: '800000', olive: '808000', lime: '00FF00',
    aqua: '00FFFF', cyan: '00FFFF', fuchsia: 'FF00FF', magenta: 'FF00FF', silver: 'C0C0C0'
  };
  /* CSS 색상 → 'RRGGBB' (투명/없음이면 null) */
  DV.colorHex = function (c) {
    if (!c) return null;
    c = String(c).trim().toLowerCase();
    if (!c || c === 'transparent' || c === 'inherit' || c === 'initial' || c === 'none' || c === 'auto') return null;
    if (NAMED[c]) return NAMED[c];
    let m = /^#([0-9a-f]{3})$/.exec(c);
    if (m) return m[1].split('').map(function (x) { return x + x; }).join('').toUpperCase();
    m = /^#([0-9a-f]{6})([0-9a-f]{2})?$/.exec(c);
    if (m) return m[1].toUpperCase();
    m = /^rgba?\(\s*(\d+)[\s,]+(\d+)[\s,]+(\d+)(?:[\s,/]+([\d.]+%?))?\s*\)$/.exec(c);
    if (m) {
      if (m[4] !== undefined) {
        const a = m[4].endsWith('%') ? parseFloat(m[4]) / 100 : parseFloat(m[4]);
        if (a === 0) return null;
      }
      return [m[1], m[2], m[3]].map(function (x) {
        return ('0' + Math.min(255, +x).toString(16)).slice(-2);
      }).join('').toUpperCase();
    }
    return null;
  };

  /* CSS 길이 → pt */
  DV.toPt = function (v, basePt) {
    if (v == null || v === '') return null;
    const s = String(v).trim().toLowerCase();
    const m = /^(-?[\d.]+)\s*(pt|px|em|rem|%|cm|mm|in)?$/.exec(s);
    if (!m) {
      const kw = { 'xx-small': 7, 'x-small': 7.5, small: 10, medium: 12, large: 13.5, 'x-large': 18, 'xx-large': 24, 'xxx-large': 36 };
      return kw[s] || null;
    }
    const n = parseFloat(m[1]);
    switch (m[2]) {
      case 'px': return n * 0.75;
      case 'em': case 'rem': return n * (basePt || 10);
      case '%': return n / 100 * (basePt || 10);
      case 'cm': return n * 28.3465;
      case 'mm': return n * 2.83465;
      case 'in': return n * 72;
      default: return n;
    }
  };

  DV.dataUrlInfo = function (url) {
    const m = /^data:([^;,]+)?(;base64)?,(.*)$/s.exec(url || '');
    if (!m) return null;
    const mime = (m[1] || 'application/octet-stream').toLowerCase();
    let bytes;
    if (m[2]) {
      const bin = atob(m[3]);
      bytes = new Uint8Array(bin.length);
      for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
    } else {
      bytes = new TextEncoder().encode(decodeURIComponent(m[3]));
    }
    const extMap = { 'image/png': 'png', 'image/jpeg': 'jpeg', 'image/jpg': 'jpeg', 'image/gif': 'gif', 'image/bmp': 'bmp', 'image/webp': 'png', 'image/svg+xml': 'svg' };
    return { mime: mime, bytes: bytes, ext: extMap[mime] || 'png' };
  };

  DV.bytesToDataUrl = function (bytes, mime) {
    let bin = '';
    const chunk = 0x8000;
    for (let i = 0; i < bytes.length; i += chunk) {
      bin += String.fromCharCode.apply(null, bytes.subarray(i, i + chunk));
    }
    return 'data:' + mime + ';base64,' + btoa(bin);
  };

  DV.mimeFromExt = function (ext) {
    ext = (ext || '').toLowerCase();
    return ({
      png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp',
      svg: 'image/svg+xml', webp: 'image/webp', tif: 'image/tiff', tiff: 'image/tiff',
      emf: 'image/x-emf', wmf: 'image/x-wmf'
    })[ext] || 'application/octet-stream';
  };

  /* 브라우저에서 표시할 수 있는 이미지인지 */
  DV.isWebImage = function (mime) {
    return /^image\/(png|jpeg|gif|bmp|webp|svg\+xml)$/.test(mime);
  };

  /* WEBP 등 docx/hwpx에서 지원하지 않는 형식을 PNG로 변환 */
  DV.ensurePng = function (dataUrl) {
    return new Promise(function (resolve) {
      const img = new Image();
      img.onload = function () {
        const cv = document.createElement('canvas');
        cv.width = img.naturalWidth || 1;
        cv.height = img.naturalHeight || 1;
        cv.getContext('2d').drawImage(img, 0, 0);
        resolve(cv.toDataURL('image/png'));
      };
      img.onerror = function () { resolve(null); };
      img.src = dataUrl;
    });
  };

  DV.inflateRaw = async function (bytes) {
    if (typeof DecompressionStream === 'undefined') {
      throw new Error('이 브라우저는 압축 해제를 지원하지 않습니다. 최신 Chrome, Edge, Safari, Firefox를 사용해 주세요.');
    }
    // 압축 데이터 뒤에 남는 바이트가 있으면 DecompressionStream이 오류를 내므로,
    // 그 전까지 풀린 데이터를 모아 사용합니다.
    const ds = new DecompressionStream('deflate-raw');
    const writer = ds.writable.getWriter();
    writer.write(bytes).catch(function () { /* 아래에서 처리 */ });
    writer.close().catch(function () { /* 아래에서 처리 */ });
    const reader = ds.readable.getReader();
    const chunks = [];
    let total = 0;
    try {
      for (;;) {
        const r = await reader.read();
        if (r.done) break;
        chunks.push(r.value);
        total += r.value.length;
      }
    } catch (err) {
      if (!total) throw new Error('압축 해제 실패: ' + err.message);
    }
    const out = new Uint8Array(total);
    let off = 0;
    chunks.forEach(function (c) { out.set(c, off); off += c.length; });
    return out;
  };

  /* 허용된 태그/속성만 남기는 간단한 HTML 정리기 (붙여넣기·가져오기용) */
  const ALLOWED = {
    P: 1, DIV: 1, SPAN: 1, BR: 1, B: 1, STRONG: 1, I: 1, EM: 1, U: 1, S: 1, STRIKE: 1, DEL: 1, INS: 1,
    SUB: 1, SUP: 1, A: 1, H1: 1, H2: 1, H3: 1, H4: 1, H5: 1, H6: 1, UL: 1, OL: 1, LI: 1,
    TABLE: 1, THEAD: 1, TBODY: 1, TFOOT: 1, TR: 1, TD: 1, TH: 1, COLGROUP: 1, COL: 1, CAPTION: 1,
    IMG: 1, BLOCKQUOTE: 1, PRE: 1, CODE: 1, HR: 1, FONT: 1, MARK: 1, SMALL: 1, BIG: 1, CENTER: 1
  };
  const DROP = { SCRIPT: 1, STYLE: 1, IFRAME: 1, OBJECT: 1, EMBED: 1, LINK: 1, META: 1, TITLE: 1, HEAD: 1, NOSCRIPT: 1, TEMPLATE: 1, SVG: 1, MATH: 1, FORM: 1, INPUT: 1, BUTTON: 1, SELECT: 1, TEXTAREA: 1 };
  const STYLE_OK = /^(color|background-color|background|font-size|font-family|font-weight|font-style|text-decoration|text-decoration-line|text-align|vertical-align|margin-left|padding-left|text-indent|width|height|border|border-collapse)$/;

  DV.sanitizeHTML = function (html) {
    const doc = new DOMParser().parseFromString('<body>' + html + '</body>', 'text/html');
    function clean(node) {
      let child = node.firstChild;
      while (child) {
        const next = child.nextSibling;
        if (child.nodeType === 1) {
          const tag = child.tagName.toUpperCase();
          if (DROP[tag]) { child.remove(); }
          else if (!ALLOWED[tag]) {
            clean(child);
            while (child.firstChild) node.insertBefore(child.firstChild, child);
            child.remove();
          } else {
            const attrs = Array.prototype.slice.call(child.attributes);
            attrs.forEach(function (a) {
              const n = a.name.toLowerCase();
              if (n === 'style') {
                const kept = [];
                String(a.value).split(';').forEach(function (decl) {
                  const i = decl.indexOf(':');
                  if (i < 0) return;
                  const p = decl.slice(0, i).trim().toLowerCase();
                  const v = decl.slice(i + 1).trim();
                  if (STYLE_OK.test(p) && !/url\s*\(|expression|javascript:/i.test(v)) kept.push(p + ':' + v);
                });
                if (kept.length) child.setAttribute('style', kept.join(';')); else child.removeAttribute('style');
              } else if (n === 'href') {
                if (!/^(https?:|mailto:|#)/i.test(a.value.trim())) child.removeAttribute('href');
              } else if (n === 'src') {
                if (!/^(data:image\/(png|jpe?g|gif|bmp|webp);|https?:|blob:)/i.test(a.value.trim())) child.removeAttribute('src');
              } else if (['colspan', 'rowspan', 'width', 'height', 'align', 'color', 'face', 'size', 'class', 'alt', 'contenteditable', 'data-dv'].indexOf(n) < 0) {
                child.removeAttribute(a.name);
              } else if (n === 'class' && !/^(page-break|dv-[\w-]+)( dv-[\w-]+)*$/.test(a.value)) {
                child.removeAttribute('class');
              }
            });
            clean(child);
          }
        } else if (child.nodeType !== 3) {
          child.remove();
        }
        child = next;
      }
    }
    clean(doc.body);
    return doc.body.innerHTML;
  };

  DV.debounce = function (fn, ms) {
    let t;
    return function () {
      const args = arguments, self = this;
      clearTimeout(t);
      t = setTimeout(function () { fn.apply(self, args); }, ms);
    };
  };

  /* A4 본문 너비(px, 96dpi): 21cm - 좌우 여백 2.5cm×2 = 16cm */
  DV.PAGE_TEXT_WIDTH_PX = 605;

  DV.icon = function (name) {
    const P = {
      undo: '<path d="M9 14 4 9l5-5"/><path d="M4 9h10.5a5.5 5.5 0 0 1 0 11H11"/>',
      redo: '<path d="m15 14 5-5-5-5"/><path d="M20 9H9.5a5.5 5.5 0 0 0 0 11H13"/>',
      left: '<path d="M4 6h16M4 10h10M4 14h16M4 18h10"/>',
      center: '<path d="M4 6h16M7 10h10M4 14h16M7 18h10"/>',
      right: '<path d="M4 6h16M10 10h10M4 14h16M10 18h10"/>',
      justify: '<path d="M4 6h16M4 10h16M4 14h16M4 18h16"/>',
      ul: '<circle cx="5" cy="7" r="1.3" fill="currentColor"/><circle cx="5" cy="12" r="1.3" fill="currentColor"/><circle cx="5" cy="17" r="1.3" fill="currentColor"/><path d="M9 7h11M9 12h11M9 17h11"/>',
      ol: '<path d="M10 7h10M10 12h10M10 17h10"/><path d="M4 5h1.5v4M4 9h3M4 14.5a1.5 1.5 0 1 1 2.6 1L4 19h3" stroke-width="1.4"/>',
      indent: '<path d="M4 6h16M10 10h10M10 14h10M4 18h16M4 9.5l3 2.5-3 2.5"/>',
      outdent: '<path d="M4 6h16M10 10h10M10 14h10M4 18h16M7 9.5 4 12l3 2.5"/>',
      table: '<rect x="3" y="4" width="18" height="16" rx="1.5"/><path d="M3 9.5h18M3 15h18M9 4v16M15 4v16"/>',
      image: '<rect x="3" y="4" width="18" height="16" rx="2"/><circle cx="8.5" cy="9.5" r="1.8"/><path d="m21 16-5-5-9 9"/>',
      link: '<path d="M10 14a4 4 0 0 0 5.7 0l3-3a4 4 0 0 0-5.7-5.7l-1 1"/><path d="M14 10a4 4 0 0 0-5.7 0l-3 3A4 4 0 0 0 11 18.7l1-1"/>',
      clear: '<path d="M6 5h12M10 5 8 19M14 5l-1 6M4 20 20 4"/>',
      search: '<circle cx="11" cy="11" r="6.5"/><path d="m20 20-4.2-4.2"/>',
      pagebreak: '<path d="M6 3v5h12V3M6 21v-5h12v5"/><path d="M3 12h3M9 12h2M13 12h2M18 12h3"/>',
      preview: '<path d="M2 12s3.5-7 10-7 10 7 10 7-3.5 7-10 7S2 12 2 12z"/><circle cx="12" cy="12" r="3"/>',
      wrap: '<path d="M4 6h16M4 12h13a3 3 0 0 1 0 6h-4m2-2-2 2 2 2M4 18h5"/>',
      merge: '<rect x="3" y="5" width="18" height="14" rx="1.5"/><path d="M8 12h8M8 12l2-2M8 12l2 2M16 12l-2-2M16 12l-2 2"/>',
      sum: '<path d="M18 5H6l6 7-6 7h12"/>',
      fill: '<path d="m19 11-8-8-8.5 8.5a2 2 0 0 0 0 2.8l5.2 5.2a2 2 0 0 0 2.8 0z"/><path d="M5 2l5 5M20 15s2 2.4 2 3.5a2 2 0 0 1-4 0c0-1.1 2-3.5 2-3.5z"/>',
      sortasc: '<path d="M7 4v16M3 16l4 4 4-4M14 6h6M14 11h4.5M14 16h3"/>',
      sortdesc: '<path d="M7 4v16M3 16l4 4 4-4M14 6h3M14 11h4.5M14 16h6"/>',
      rowins: '<rect x="3" y="4" width="18" height="6" rx="1"/><rect x="3" y="14" width="18" height="6" rx="1"/><path d="M12 10v4"/>',
      close: '<path d="M6 6l12 12M18 6 6 18"/>'
    };
    return '<svg viewBox="0 0 24 24" aria-hidden="true">' + (P[name] || '') + '</svg>';
  };
})(window.DV);
