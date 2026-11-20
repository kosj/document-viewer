/* 문서(Word·한글) 편집기 화면 */
window.DV = window.DV || {};
(function (DV) {
  'use strict';

  const FONTS = [
    ['', '글꼴'], ["'맑은 고딕','Malgun Gothic',sans-serif", '맑은 고딕'], ["'나눔고딕','NanumGothic',sans-serif", '나눔고딕'],
    ["'바탕','Batang',serif", '바탕'], ["'굴림','Gulim',sans-serif", '굴림'], ["'돋움','Dotum',sans-serif", '돋움'],
    ["'궁서','Gungsuh',serif", '궁서'], ['Arial,sans-serif', 'Arial'], ["'Times New Roman',serif", 'Times New Roman'],
    ['Consolas,monospace', 'Consolas']
  ];
  const SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 48];
  const COLORS = ['000000', '444444', '888888', 'CCCCCC', 'FFFFFF', 'C00000', 'FF0000', 'FFC000', 'FFFF00', '92D050',
    '00B050', '00B0F0', '0070C0', '002060', '7030A0', 'F4B183', 'FFE699', 'C5E0B4', 'BDD7EE', 'D9D2E9'];

  function btn(cmd, title, content, extra) {
    return '<button type="button" class="tb" data-cmd="' + cmd + '" title="' + title + '"' + (extra || '') + '>' + content + '</button>';
  }

  function createDocEditor(container, opts) {
    opts = opts || {};
    const el = document.createElement('div');
    el.className = 'doc-view';
    el.innerHTML =
      '<div class="toolbar" role="toolbar">' +
        '<div class="tb-group">' + btn('undo', '실행 취소 (Ctrl+Z)', DV.icon('undo')) + btn('redo', '다시 실행 (Ctrl+Y)', DV.icon('redo')) + '</div>' +
        '<div class="tb-group">' +
          '<select class="tb-select" data-cmd="block" title="문단 스타일"><option value="p">본문</option><option value="h1">제목 1</option><option value="h2">제목 2</option><option value="h3">제목 3</option><option value="h4">제목 4</option><option value="blockquote">인용</option></select>' +
          '<select class="tb-select font" data-cmd="font" title="글꼴">' + FONTS.map(function (f) { return '<option value="' + DV.esc(f[0]) + '">' + f[1] + '</option>'; }).join('') + '</select>' +
          '<select class="tb-select size" data-cmd="size" title="글자 크기"><option value="">크기</option>' + SIZES.map(function (s) { return '<option value="' + s + '">' + s + '</option>'; }).join('') + '</select>' +
        '</div>' +
        '<div class="tb-group">' +
          btn('bold', '굵게 (Ctrl+B)', '<b>가</b>') + btn('italic', '기울임 (Ctrl+I)', '<i style="font-family:serif">가</i>') +
          btn('underline', '밑줄 (Ctrl+U)', '<u>가</u>') + btn('strikeThrough', '취소선', '<s>가</s>') +
          btn('superscript', '위 첨자', 'x<sup>2</sup>') + btn('subscript', '아래 첨자', 'x<sub>2</sub>') +
          '<span class="color-btn" data-kind="foreColor" title="글자 색"><button type="button" class="tb" data-cmd="foreColor"><span class="glyph">가</span><span class="bar" style="background:#C00000"></span></button><button type="button" class="tb caret-sm" data-cmd="palette-fore">▾</button></span>' +
          '<span class="color-btn" data-kind="hiliteColor" title="형광펜"><button type="button" class="tb" data-cmd="hiliteColor"><span class="glyph hl">가</span><span class="bar" style="background:#FFFF00"></span></button><button type="button" class="tb caret-sm" data-cmd="palette-back">▾</button></span>' +
        '</div>' +
        '<div class="tb-group">' +
          btn('justifyLeft', '왼쪽 정렬', DV.icon('left')) + btn('justifyCenter', '가운데 정렬', DV.icon('center')) +
          btn('justifyRight', '오른쪽 정렬', DV.icon('right')) + btn('justifyFull', '양쪽 정렬', DV.icon('justify')) +
        '</div>' +
        '<div class="tb-group">' +
          btn('insertUnorderedList', '글머리 기호', DV.icon('ul')) + btn('insertOrderedList', '번호 매기기', DV.icon('ol')) +
          btn('outdent', '내어쓰기', DV.icon('outdent')) + btn('indent', '들여쓰기', DV.icon('indent')) +
        '</div>' +
        '<div class="tb-group">' +
          btn('table', '표 넣기', DV.icon('table')) + btn('image', '그림 넣기', DV.icon('image')) + btn('link', '링크', DV.icon('link')) +
          btn('pagebreak', '쪽 나누기', DV.icon('pagebreak')) + btn('removeFormat', '서식 지우기', DV.icon('clear')) +
        '</div>' +
        '<div class="tb-group">' + btn('find', '찾기/바꾸기 (Ctrl+F)', DV.icon('search')) + btn('preview', '페이지 미리보기', DV.icon('preview')) + '</div>' +
        '<div class="tb-group zoom"><button type="button" class="tb" data-cmd="zoomOut" title="축소">−</button><span class="zoom-val">100%</span><button type="button" class="tb" data-cmd="zoomIn" title="확대">+</button></div>' +
      '</div>' +
      '<div class="findbar" hidden><input type="text" class="find-q" placeholder="찾을 내용"><input type="text" class="find-r" placeholder="바꿀 내용">' +
        '<button type="button" class="btn sm" data-find="next">다음 찾기</button><button type="button" class="btn sm" data-find="replace">바꾸기</button><button type="button" class="btn sm" data-find="all">모두 바꾸기</button>' +
        '<span class="find-info"></span><button type="button" class="tb" data-find="close" title="닫기">' + DV.icon('close') + '</button></div>' +
      '<div class="doc-scroll"><div class="page-wrap"><div class="page" contenteditable="true" spellcheck="false"></div></div></div>' +
      '<div class="preview-pane" hidden><div class="preview-head"><b>페이지 미리보기</b><span>Word로 저장했을 때의 쪽 나눔 모양입니다 (읽기 전용)</span><button type="button" class="btn sm" data-preview="close">닫기</button></div><div class="preview-body"></div></div>' +
      '<div class="statusbar"><span class="st-kind"></span><span class="st-count"></span></div>' +
      '<div class="palette" hidden></div>' +
      '<div class="ctxmenu" hidden></div>' +
      '<input type="file" class="img-input" accept="image/*" hidden>';
    container.appendChild(el);

    const page = el.querySelector('.page');
    const scroll = el.querySelector('.doc-scroll');
    const toolbar = el.querySelector('.toolbar');
    const palette = el.querySelector('.palette');
    const ctxmenu = el.querySelector('.ctxmenu');
    const findbar = el.querySelector('.findbar');
    const statusCount = el.querySelector('.st-count');
    el.querySelector('.st-kind').textContent = opts.kindLabel || '';
    const listeners = [];
    let zoom = 1;
    let savedRange = null;
    let foreColor = 'C00000', backColor = 'FFFF00';

    if (opts.defaultPt) page.style.fontSize = opts.defaultPt + 'pt';
    page.innerHTML = DV.sanitizeHTML(opts.html || '<p><br></p>');
    normalize();
    try { document.execCommand('defaultParagraphSeparator', false, 'p'); } catch (e) { /* 무시 */ }

    function normalize() {
      if (!page.firstChild) page.innerHTML = '<p><br></p>';
      page.querySelectorAll('.page-break').forEach(function (pb) { pb.setAttribute('contenteditable', 'false'); });
      page.querySelectorAll('table').forEach(function (t) { t.classList.add('dv-table'); });
    }

    function changed() { listeners.forEach(function (f) { f(); }); updateCount(); }
    const updateCount = DV.debounce(function () {
      const text = page.innerText || '';
      const chars = text.replace(/\s/g, '').length;
      const words = (text.trim().match(/\S+/g) || []).length;
      statusCount.textContent = '글자 ' + chars.toLocaleString() + ' · 단어 ' + words.toLocaleString();
    }, 250);
    updateCount();

    page.addEventListener('input', changed);

    function saveSel() {
      const sel = window.getSelection();
      if (sel.rangeCount && page.contains(sel.getRangeAt(0).commonAncestorContainer)) savedRange = sel.getRangeAt(0).cloneRange();
    }
    function restoreSel() {
      if (!savedRange) { page.focus(); return; }
      page.focus();
      const sel = window.getSelection();
      sel.removeAllRanges();
      sel.addRange(savedRange);
    }
    document.addEventListener('selectionchange', function () {
      if (!el.isConnected || el.offsetParent === null) return;
      saveSel();
      updateState();
    });

    function exec(cmd, val) {
      restoreSel();
      try { document.execCommand('styleWithCSS', false, cmd === 'hiliteColor' || cmd === 'foreColor'); } catch (e) { /* 무시 */ }
      document.execCommand(cmd, false, val == null ? null : val);
      saveSel();
      changed();
      updateState();
    }

    /* 선택 영역을 span으로 감싸 스타일 적용 (글꼴·크기) */
    function applySpanStyle(prop, value) {
      restoreSel();
      const sel = window.getSelection();
      if (!sel.rangeCount) return;
      if (sel.isCollapsed) {
        // 선택이 없으면 이후 입력할 글자에 적용될 빈 span을 만듭니다
        const span = document.createElement('span');
        span.style[prop] = value;
        span.textContent = '​';
        const r = sel.getRangeAt(0);
        r.insertNode(span);
        const nr = document.createRange();
        nr.setStart(span.firstChild, 1);
        nr.collapse(true);
        sel.removeAllRanges();
        sel.addRange(nr);
        saveSel();
        return;
      }
      const marker = 'dv-tmp-' + Date.now();
      try { document.execCommand('styleWithCSS', false, false); } catch (e) { /* 무시 */ }
      document.execCommand('fontName', false, marker);
      page.querySelectorAll('font[face="' + marker + '"]').forEach(function (f) {
        const span = document.createElement('span');
        span.style[prop] = value;
        while (f.firstChild) span.appendChild(f.firstChild);
        f.replaceWith(span);
        // 안쪽의 같은 속성 제거
        span.querySelectorAll('span').forEach(function (s) { s.style[prop] = ''; if (!s.getAttribute('style')) { s.replaceWith.apply(s, Array.from(s.childNodes)); } });
      });
      page.querySelectorAll('span[style*="' + marker + '"]').forEach(function (s) { s.style.fontFamily = ''; s.style[prop] = value; });
      saveSel();
      changed();
    }

    function currentBlock() {
      const sel = window.getSelection();
      if (!sel.rangeCount) return null;
      let n = sel.anchorNode;
      while (n && n !== page) {
        if (n.nodeType === 1 && /^(P|H[1-6]|LI|BLOCKQUOTE|DIV|PRE|TD|TH)$/.test(n.tagName)) return n;
        n = n.parentNode;
      }
      return null;
    }

    function updateState() {
      ['bold', 'italic', 'underline', 'strikeThrough', 'superscript', 'subscript', 'justifyLeft', 'justifyCenter', 'justifyRight', 'justifyFull', 'insertUnorderedList', 'insertOrderedList'].forEach(function (c) {
        const b = toolbar.querySelector('[data-cmd="' + c + '"]');
        let on = false;
        try { on = document.queryCommandState(c); } catch (e) { /* 무시 */ }
        if (b) b.classList.toggle('on', !!on);
      });
      const blk = currentBlock();
      const sel = toolbar.querySelector('[data-cmd="block"]');
      if (blk && sel) {
        let n = blk;
        while (n && n !== page && !/^(H[1-6]|BLOCKQUOTE)$/.test(n.tagName)) n = n.parentNode;
        sel.value = n && n !== page ? n.tagName.toLowerCase() : 'p';
      }
      const sizeSel = toolbar.querySelector('[data-cmd="size"]');
      const s = window.getSelection();
      if (s.rangeCount && sizeSel) {
        let node = s.anchorNode;
        if (node && node.nodeType === 3) node = node.parentNode;
        if (node && page.contains(node)) {
          const pt = Math.round(parseFloat(getComputedStyle(node).fontSize) * 0.75 * 2) / 2;
          sizeSel.value = SIZES.indexOf(pt) >= 0 ? String(pt) : '';
        }
      }
    }

    function showPalette(anchor, kind) {
      palette.innerHTML = '<div class="pal-grid">' + COLORS.map(function (c) {
        return '<button type="button" style="background:#' + c + '" data-color="' + c + '" title="#' + c + '"></button>';
      }).join('') + '</div>' + (kind === 'back' ? '<button type="button" class="pal-none" data-color="none">색 없음</button>' : '<button type="button" class="pal-none" data-color="000000">자동(검정)</button>');
      palette.dataset.kind = kind;
      const r = anchor.getBoundingClientRect(), er = el.getBoundingClientRect();
      palette.style.left = Math.min(r.left - er.left, er.width - 220) + 'px';
      palette.style.top = (r.bottom - er.top + 4) + 'px';
      palette.hidden = false;
    }

    function insertHTML(html) {
      restoreSel();
      document.execCommand('insertHTML', false, html);
      saveSel();
      normalize();
      changed();
    }

    function insertTable() {
      const spec = prompt('표 크기를 입력하세요 (열 x 행)', '3 x 3');
      if (!spec) return;
      const m = /(\d+)\s*[x×*,\s]\s*(\d+)/i.exec(spec);
      const cols = Math.min(30, Math.max(1, m ? +m[1] : 3)), rows = Math.min(200, Math.max(1, m ? +m[2] : 3));
      let h = '<table class="dv-table" style="width:100%"><tbody>';
      for (let r = 0; r < rows; r++) {
        h += '<tr>';
        for (let c = 0; c < cols; c++) h += '<td><p><br></p></td>';
        h += '</tr>';
      }
      h += '</tbody></table>';
      const tmp = document.createElement('div');
      tmp.innerHTML = h;
      const table = tmp.firstChild;
      insertBlock(table);
      const first = table.querySelector('td');
      const r = document.createRange();
      r.selectNodeContents(first.firstChild);
      r.collapse(true);
      const s = window.getSelection(); s.removeAllRanges(); s.addRange(r);
      saveSel();
      changed();
    }

    /* 현재 문단 뒤에 블록(표·쪽 나누기)을 넣습니다. 빈 문단이면 그 자리를 대신합니다. */
    function topBlock() {
      if (!savedRange) return null;
      let n = savedRange.startContainer;
      while (n && n.parentNode !== page) n = n.parentNode;
      return n && n.parentNode === page ? n : null;
    }
    function insertBlock(node) {
      restoreSel();
      const blk = topBlock();
      const isEmpty = blk && blk.nodeType === 1 && /^(P|DIV|H[1-6])$/.test(blk.tagName) && !blk.textContent.replace(/\u200b/g, '') && !blk.querySelector('img,table');
      if (!blk) page.appendChild(node);
      else if (isEmpty) blk.replaceWith(node);
      else blk.after(node);
      if (!node.nextElementSibling || node.nextElementSibling.classList.contains('page-break') || node.nextElementSibling.tagName === 'TABLE') {
        const p = document.createElement('p');
        p.innerHTML = '<br>';
        node.after(p);
      }
      normalize();
    }

    function insertImage(file) {
      const fr = new FileReader();
      fr.onload = function () {
        const img = new Image();
        img.onload = function () {
          let w = img.naturalWidth, h = img.naturalHeight;
          if (w > DV.PAGE_TEXT_WIDTH_PX) { h = Math.round(h * DV.PAGE_TEXT_WIDTH_PX / w); w = DV.PAGE_TEXT_WIDTH_PX; }
          insertHTML('<img src="' + fr.result + '" style="width:' + w + 'px;height:' + h + 'px" alt="">');
        };
        img.src = fr.result;
      };
      fr.readAsDataURL(file);
    }

    toolbar.addEventListener('mousedown', function (e) {
      // 버튼 클릭 시 편집 영역의 선택이 풀리지 않도록
      if (e.target.closest('button')) e.preventDefault();
    });
    toolbar.addEventListener('click', function (e) {
      const b = e.target.closest('button[data-cmd]');
      if (!b) return;
      const cmd = b.dataset.cmd;
      palette.hidden = true;
      switch (cmd) {
        case 'undo': case 'redo': exec(cmd); break;
        case 'foreColor': exec('foreColor', '#' + foreColor); break;
        case 'hiliteColor': exec('hiliteColor', '#' + backColor); break;
        case 'palette-fore': showPalette(b, 'fore'); break;
        case 'palette-back': showPalette(b, 'back'); break;
        case 'table': insertTable(); break;
        case 'image': el.querySelector('.img-input').click(); break;
        case 'link': {
          const url = prompt('링크 주소(URL)를 입력하세요', 'https://');
          if (url && /^(https?:|mailto:)/i.test(url.trim())) exec('createLink', url.trim());
          break;
        }
        case 'pagebreak': {
          const pb = document.createElement('div');
          pb.className = 'page-break';
          pb.setAttribute('contenteditable', 'false');
          insertBlock(pb);
          const nx = pb.nextElementSibling;
          const r = document.createRange();
          r.selectNodeContents(nx); r.collapse(true);
          const s = window.getSelection(); s.removeAllRanges(); s.addRange(r);
          saveSel(); changed();
          break;
        }
        case 'find': toggleFind(true); break;
        case 'preview': showPreview(); break;
        case 'zoomIn': setZoom(zoom + 0.1); break;
        case 'zoomOut': setZoom(zoom - 0.1); break;
        default: exec(cmd);
      }
    });
    toolbar.addEventListener('change', function (e) {
      const s = e.target.closest('select');
      if (!s) return;
      const v = s.value;
      if (s.dataset.cmd === 'block') exec('formatBlock', '<' + v + '>');
      else if (s.dataset.cmd === 'font' && v) applySpanStyle('fontFamily', v);
      else if (s.dataset.cmd === 'size' && v) applySpanStyle('fontSize', v + 'pt');
      if (s.dataset.cmd === 'font') s.value = '';
      restoreSel();
    });
    palette.addEventListener('mousedown', function (e) { e.preventDefault(); });
    palette.addEventListener('click', function (e) {
      const b = e.target.closest('[data-color]');
      if (!b) return;
      const c = b.dataset.color;
      if (palette.dataset.kind === 'fore') {
        foreColor = c;
        toolbar.querySelector('[data-kind="foreColor"] .bar').style.background = '#' + c;
        exec('foreColor', '#' + c);
      } else {
        if (c === 'none') { exec('hiliteColor', 'transparent'); }
        else {
          backColor = c;
          toolbar.querySelector('[data-kind="hiliteColor"] .bar').style.background = '#' + c;
          exec('hiliteColor', '#' + c);
        }
      }
      palette.hidden = true;
    });
    document.addEventListener('mousedown', function (e) {
      if (!palette.hidden && !palette.contains(e.target) && !e.target.closest('.color-btn')) palette.hidden = true;
      if (!ctxmenu.hidden && !ctxmenu.contains(e.target)) ctxmenu.hidden = true;
    });
    el.querySelector('.img-input').addEventListener('change', function (e) {
      const f = e.target.files[0];
      if (f) insertImage(f);
      e.target.value = '';
    });

    // 붙여넣기: 안전한 HTML만 남김
    page.addEventListener('paste', function (e) {
      const cd = e.clipboardData;
      if (!cd) return;
      const files = Array.from(cd.files || []).filter(function (f) { return /^image\//.test(f.type); });
      if (files.length) { e.preventDefault(); saveSel(); files.forEach(insertImage); return; }
      const html = cd.getData('text/html');
      if (html) {
        e.preventDefault();
        let clean = DV.sanitizeHTML(html.replace(/<!--[\s\S]*?-->/g, ''));
        // 서식용 줄바꿈 제거 (본문은 공백 보존 모드)
        clean = clean.replace(/>\s*\n\s*</g, '><').replace(/\n/g, ' ');
        document.execCommand('insertHTML', false, clean);
        normalize();
        changed();
      }
    });
    page.addEventListener('drop', function (e) {
      const files = Array.from((e.dataTransfer && e.dataTransfer.files) || []).filter(function (f) { return /^image\//.test(f.type); });
      if (files.length) {
        e.preventDefault();
        e.stopPropagation();
        if (document.caretRangeFromPoint) {
          const r = document.caretRangeFromPoint(e.clientX, e.clientY);
          if (r) { const s = window.getSelection(); s.removeAllRanges(); s.addRange(r); saveSel(); }
        }
        files.forEach(insertImage);
      }
    });

    page.addEventListener('keydown', function (e) {
      const mod = e.ctrlKey || e.metaKey;
      if (e.key === 'Tab') {
        const td = closestCell();
        if (td) { e.preventDefault(); moveCell(td, e.shiftKey ? -1 : 1); return; }
        e.preventDefault();
        if (inList()) exec(e.shiftKey ? 'outdent' : 'indent'); else exec('insertText', '\t');
        return;
      }
    });
    el.addEventListener('keydown', function (e) {
      const mod = e.ctrlKey || e.metaKey;
      const k = e.key.toLowerCase();
      if (mod && (k === 'f' || k === 'h')) { e.preventDefault(); toggleFind(true); }
    });

    function inList() {
      const b = currentBlock();
      return b && b.closest('li');
    }
    function closestCell() {
      const sel = window.getSelection();
      if (!sel.rangeCount) return null;
      let n = sel.anchorNode;
      if (n && n.nodeType === 3) n = n.parentNode;
      const td = n && n.closest ? n.closest('td,th') : null;
      return td && page.contains(td) ? td : null;
    }
    function moveCell(td, dir) {
      const table = td.closest('table');
      const cells = Array.from(table.querySelectorAll('td,th')).filter(function (c) { return c.closest('table') === table; });
      let i = cells.indexOf(td) + dir;
      if (i >= cells.length) { tableOp('rowBelow', cells[cells.length - 1]); return moveCell(cells[cells.length - 1], 1); }
      if (i < 0) i = 0;
      const target = cells[i];
      const r = document.createRange();
      r.selectNodeContents(target);
      r.collapse(true);
      const s = window.getSelection(); s.removeAllRanges(); s.addRange(r);
      saveSel();
    }

    /* ----- 표 편집 (오른쪽 클릭 메뉴) ----- */
    function tableOp(op, td) {
      const tr = td.parentElement;
      const table = td.closest('table');
      if (!table || !tr) return;
      const idx = Array.prototype.indexOf.call(tr.children, td);
      const rows = Array.from(table.rows);
      function newCell() { const c = document.createElement('td'); c.innerHTML = '<p><br></p>'; return c; }
      switch (op) {
        case 'rowAbove': case 'rowBelow': {
          const nr = document.createElement('tr');
          let n = 0;
          Array.from(tr.children).forEach(function (c) { n += c.colSpan || 1; });
          for (let i = 0; i < n; i++) nr.appendChild(newCell());
          if (op === 'rowAbove') tr.before(nr); else tr.after(nr);
          break;
        }
        case 'colLeft': case 'colRight': {
          rows.forEach(function (row) {
            const ref = row.children[Math.min(idx, row.children.length - 1)];
            const c = document.createElement(ref && ref.tagName === 'TH' ? 'th' : 'td');
            c.innerHTML = '<p><br></p>';
            if (!ref) row.appendChild(c);
            else if (op === 'colLeft') ref.before(c); else ref.after(c);
          });
          table.querySelectorAll('colgroup').forEach(function (cg) { cg.remove(); });
          break;
        }
        case 'delRow':
          if (rows.length <= 1) table.remove(); else tr.remove();
          break;
        case 'delCol':
          rows.forEach(function (row) { const c = row.children[Math.min(idx, row.children.length - 1)]; if (c) { if (c.colSpan > 1) c.colSpan--; else c.remove(); } });
          table.querySelectorAll('colgroup').forEach(function (cg) { cg.remove(); });
          if (!table.querySelector('td,th')) table.remove();
          break;
        case 'delTable': table.remove(); break;
        case 'cellColor': {
          const c = prompt('셀 배경색 (예: #FFF2CC, 비우면 없음)', td.style.backgroundColor ? '#' + DV.colorHex(td.style.backgroundColor) : '#FFF2CC');
          if (c === null) return;
          td.style.backgroundColor = DV.colorHex(c) ? '#' + DV.colorHex(c) : '';
          break;
        }
        case 'border': table.classList.toggle('dv-noborder'); break;
      }
      changed();
    }

    page.addEventListener('contextmenu', function (e) {
      const td = e.target.closest('td,th');
      const img = e.target.closest('img');
      if (!td && !img) return;
      e.preventDefault();
      let items;
      if (img) {
        items = [['img25', '그림 크기 25%'], ['img50', '그림 크기 50%'], ['img75', '그림 크기 75%'], ['img100', '그림 크기 100%'], ['imgDel', '그림 삭제']];
      } else {
        items = [['rowAbove', '위에 행 삽입'], ['rowBelow', '아래에 행 삽입'], ['colLeft', '왼쪽에 열 삽입'], ['colRight', '오른쪽에 열 삽입'], ['-'],
          ['delRow', '행 삭제'], ['delCol', '열 삭제'], ['delTable', '표 삭제'], ['-'], ['cellColor', '셀 배경색…'], ['border', '테두리 표시/숨기기']];
      }
      ctxmenu.innerHTML = items.map(function (it) { return it[0] === '-' ? '<hr>' : '<button type="button" data-op="' + it[0] + '">' + it[1] + '</button>'; }).join('');
      const er = el.getBoundingClientRect();
      ctxmenu.style.left = Math.min(e.clientX - er.left, er.width - 190) + 'px';
      ctxmenu.style.top = Math.min(e.clientY - er.top, er.height - items.length * 32 - 10) + 'px';
      ctxmenu.hidden = false;
      ctxmenu._target = img || td;
    });
    ctxmenu.addEventListener('click', function (e) {
      const b = e.target.closest('[data-op]');
      if (!b) return;
      const t = ctxmenu._target;
      ctxmenu.hidden = true;
      const op = b.dataset.op;
      if (/^img/.test(op)) {
        if (op === 'imgDel') t.remove();
        else {
          const pct = +op.slice(3) / 100;
          const nw = t.naturalWidth || t.width, nh = t.naturalHeight || t.height;
          let w = Math.round(Math.min(nw, DV.PAGE_TEXT_WIDTH_PX) * pct);
          t.style.width = w + 'px';
          t.style.height = Math.round(w * nh / nw) + 'px';
        }
        changed();
        return;
      }
      tableOp(op, t);
    });

    /* ----- 찾기/바꾸기 ----- */
    const fq = findbar.querySelector('.find-q'), fr = findbar.querySelector('.find-r'), finfo = findbar.querySelector('.find-info');
    function toggleFind(show) {
      findbar.hidden = !show;
      if (show) {
        const s = window.getSelection().toString();
        if (s && s.length < 100) fq.value = s;
        fq.focus(); fq.select();
      }
    }
    function textNodes() {
      const out = [];
      const w = document.createTreeWalker(page, NodeFilter.SHOW_TEXT);
      let n;
      while ((n = w.nextNode())) out.push(n);
      return out;
    }
    let lastFound = null;
    function findNext() {
      const q = fq.value;
      if (!q) return false;
      const nodes = textNodes();
      let startNode = null, startOff = 0;
      if (lastFound && page.contains(lastFound.node)) { startNode = lastFound.node; startOff = lastFound.off + 1; }
      let started = !startNode;
      for (let pass = 0; pass < 2; pass++) {
        for (const n of nodes) {
          let from = 0;
          if (!started) { if (n !== startNode) continue; started = true; from = startOff; }
          const i = n.nodeValue.toLowerCase().indexOf(q.toLowerCase(), from);
          if (i >= 0) {
            const r = document.createRange();
            r.setStart(n, i); r.setEnd(n, i + q.length);
            const s = window.getSelection(); s.removeAllRanges(); s.addRange(r);
            savedRange = r.cloneRange();
            lastFound = { node: n, off: i };
            const rect = r.getBoundingClientRect(), sr = scroll.getBoundingClientRect();
            if (rect.top < sr.top + 40 || rect.bottom > sr.bottom - 40) scroll.scrollTop += rect.top - sr.top - sr.height / 3;
            finfo.textContent = '';
            return true;
          }
        }
        started = true; startNode = null;
      }
      finfo.textContent = '찾을 수 없음';
      lastFound = null;
      return false;
    }
    function replaceOne() {
      const q = fq.value;
      const s = window.getSelection();
      if (q && s.rangeCount && s.toString().toLowerCase() === q.toLowerCase() && page.contains(s.anchorNode)) {
        document.execCommand('insertText', false, fr.value);
        changed();
      }
      findNext();
    }
    function replaceAll() {
      const q = fq.value;
      if (!q) return;
      let count = 0;
      const re = new RegExp(q.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'gi');
      textNodes().forEach(function (n) {
        const v = n.nodeValue;
        const nv = v.replace(re, function () { count++; return fr.value; });
        if (nv !== v) n.nodeValue = nv;
      });
      finfo.textContent = count + '개 바꿈';
      if (count) changed();
    }
    findbar.addEventListener('click', function (e) {
      const b = e.target.closest('[data-find]');
      if (!b) return;
      const a = b.dataset.find;
      if (a === 'next') findNext();
      else if (a === 'replace') replaceOne();
      else if (a === 'all') replaceAll();
      else toggleFind(false);
    });
    findbar.addEventListener('keydown', function (e) {
      if (e.key === 'Enter') { e.preventDefault(); if (e.target === fr) replaceOne(); else findNext(); }
      if (e.key === 'Escape') { toggleFind(false); page.focus(); }
    });

    function setZoom(z) {
      zoom = Math.max(0.5, Math.min(2, Math.round(z * 10) / 10));
      el.querySelector('.page-wrap').style.zoom = zoom;
      el.querySelector('.zoom-val').textContent = Math.round(zoom * 100) + '%';
    }
    // 화면이 좁으면 자동으로 축소
    function fitWidth() {
      const avail = scroll.clientWidth - 24;
      const pw = 794;
      if (avail > 0 && avail < pw) setZoom(Math.max(0.5, Math.floor(avail / pw * 10) / 10));
    }
    setTimeout(fitWidth, 0);

    async function showPreview() {
      const pane = el.querySelector('.preview-pane');
      const body = pane.querySelector('.preview-body');
      pane.hidden = false;
      body.innerHTML = '<div class="loading">미리보기를 만드는 중…</div>';
      try {
        const blob = await DV.word.htmlToDocx(page, {});
        body.innerHTML = '';
        await window.docx.renderAsync(blob, body, null, { inWrapper: true, ignoreWidth: false, ignoreHeight: false, breakPages: true, experimental: false });
      } catch (err) {
        body.innerHTML = '<div class="loading">미리보기를 만들 수 없습니다: ' + DV.esc(err.message) + '</div>';
      }
    }
    el.querySelector('[data-preview="close"]').addEventListener('click', function () { el.querySelector('.preview-pane').hidden = true; });

    return {
      el: el,
      page: page,
      getHTML: function () { return page.innerHTML; },
      onChange: function (f) { listeners.push(f); },
      focus: function () { page.focus(); },
      destroy: function () { el.remove(); }
    };
  }

  DV.createDocEditor = createDocEditor;
})(window.DV);
