/* 앱: 파일 열기·새로 만들기·탭·저장 */
(function (DV) {
  'use strict';

  const $ = function (s) { return document.querySelector(s); };
  const workspace = $('#workspace');
  const home = $('#home');
  const tabsEl = $('#tabs');
  const fileInput = $('#file-input');
  const docs = [];
  let active = null;
  let seq = 0;
  let lastPointerTouch = false;
  window.addEventListener('pointerdown', function (e) { lastPointerTouch = e.pointerType === 'touch'; }, true);

  const KIND = {
    docx: { label: 'Word 문서', icon: 'W', cls: 'fi-word', type: 'doc' },
    hwp: { label: '한글 문서 (HWP)', icon: '한', cls: 'fi-hwp', type: 'doc' },
    hwpx: { label: '한글 문서 (HWPX)', icon: '한', cls: 'fi-hwp', type: 'doc' },
    xlsx: { label: 'Excel 통합 문서', icon: 'X', cls: 'fi-excel', type: 'sheet' },
    xlsm: { label: 'Excel 매크로 통합 문서', icon: 'X', cls: 'fi-excel', type: 'sheet' },
    xls: { label: 'Excel 97-2003 통합 문서', icon: 'X', cls: 'fi-excel', type: 'sheet' },
    ods: { label: 'OpenDocument 스프레드시트', icon: 'X', cls: 'fi-excel', type: 'sheet' },
    csv: { label: 'CSV (쉼표로 구분)', icon: 'X', cls: 'fi-excel', type: 'sheet' }
  };

  function setDirty(doc, v) {
    doc.dirty = v;
    if (doc.tabEl) doc.tabEl.classList.toggle('dirty', v);
  }

  function renderTabs() {
    tabsEl.hidden = docs.length === 0;
    tabsEl.innerHTML = '';
    docs.forEach(function (d) {
      const k = KIND[d.format] || KIND.docx;
      const t = document.createElement('div');
      t.className = 'tab' + (d === active ? ' on' : '') + (d.dirty ? ' dirty' : '');
      t.setAttribute('role', 'tab');
      t.title = d.name;
      t.innerHTML = '<i class="fi ' + k.cls + '">' + k.icon + '</i><span class="tab-name"></span><span class="dot" title="저장하지 않은 변경 사항"></span><button type="button" class="tab-close" title="닫기">' + DV.icon('close') + '</button>';
      t.querySelector('.tab-name').textContent = d.name;
      t.addEventListener('click', function (e) {
        if (e.target.closest('.tab-close')) { closeDoc(d); return; }
        activate(d);
      });
      t.addEventListener('auxclick', function (e) { if (e.button === 1) closeDoc(d); });
      d.tabEl = t;
      tabsEl.appendChild(t);
    });
    const on = !!active;
    $('#btn-save').disabled = !on;
    $('#btn-save-as').disabled = !on;
    $('#btn-print').disabled = !on;
    home.hidden = on;
    document.title = active ? active.name + ' — 문서 편집기' : '문서 뷰어 · 편집기 — Word · Excel · 한글';
  }

  function activate(d) {
    if (active && active.view && active.view.commitEdit) active.view.commitEdit();
    active = d;
    docs.forEach(function (x) { x.host.hidden = x !== d; });
    renderTabs();
    if (d && d.view.refresh) d.view.refresh();
    if (d) setTimeout(function () { d.view.focus && d.view.focus(); }, 0);
  }

  function closeDoc(d) {
    if (d.dirty && !confirm('"' + d.name + '"의 변경 사항을 저장하지 않았습니다. 닫을까요?')) return;
    const i = docs.indexOf(d);
    docs.splice(i, 1);
    d.view.destroy();
    d.host.remove();
    if (active === d) active = docs[Math.min(i, docs.length - 1)] || null;
    if (active) activate(active); else renderTabs();
  }

  function addDoc(opts) {
    const host = document.createElement('section');
    host.className = 'doc-host';
    workspace.appendChild(host);
    const k = KIND[opts.format] || KIND.docx;
    const d = { id: ++seq, name: opts.name, format: opts.format, host: host, dirty: false, type: k.type, isNew: !!opts.isNew };
    if (k.type === 'sheet') {
      d.view = DV.createSheetEditor(host, opts.book, { kindLabel: k.label, isTouch: function () { return lastPointerTouch; } });
    } else {
      d.view = DV.createDocEditor(host, { html: opts.html, defaultPt: opts.defaultPt, kindLabel: k.label });
    }
    d.view.onChange(function () { if (!d.dirty) setDirty(d, true); });
    docs.push(d);
    activate(d);
    return d;
  }

  /* ---------- 열기 ---------- */
  async function openFile(file) {
    const ext = DV.ext(file.name);
    const loading = DV.toast ? showLoading(file.name) : null;
    try {
      const buf = await DV.readFile(file);
      let res;
      if (ext === 'docx' || ext === 'docm' || ext === 'dotx') {
        res = await DV.word.docxToHtml(buf);
        addDoc({ name: file.name, format: 'docx', html: res.html, defaultPt: res.defaultPt });
      } else if (ext === 'hwpx') {
        res = await DV.hwpx.hwpxToHtml(buf);
        addDoc({ name: file.name, format: 'hwpx', html: res.html, defaultPt: res.defaultPt });
      } else if (ext === 'hwp') {
        const b = new Uint8Array(buf);
        if (b[0] === 0x50 && b[1] === 0x4B) res = await DV.hwpx.hwpxToHtml(buf);
        else res = await DV.hwp.hwpToHtml(buf);
        addDoc({ name: file.name, format: 'hwp', html: res.html, defaultPt: res.defaultPt });
      } else if (KIND[ext] && KIND[ext].type === 'sheet' || ext === 'tsv' || ext === 'txt') {
        const book = await DV.xlsxIO.readWorkbook(buf, ext);
        addDoc({ name: file.name, format: KIND[ext] ? ext : 'csv', book: book });
      } else if (ext === 'doc') {
        throw new Error('예전 Word 형식(.doc)은 지원하지 않습니다. Word에서 .docx로 저장한 뒤 열어 주세요.');
      } else if (ext === 'ppt' || ext === 'pptx') {
        throw new Error('PowerPoint 파일은 지원하지 않습니다.');
      } else {
        throw new Error('지원하지 않는 파일 형식입니다: .' + (ext || '?'));
      }
      if (res && res.warnings && res.warnings.length) res.warnings.forEach(function (w) { DV.toast(w, 'warn', 6000); });
      if (ext === 'hwp') DV.toast('HWP 파일은 편집 후 .hwpx(한글 2014 이상) 또는 .docx로 저장됩니다.', 'info', 6000);
    } catch (err) {
      console.error(err);
      DV.toast('"' + file.name + '"을(를) 열 수 없습니다. ' + (err && err.message ? err.message : ''), 'error', 8000);
    } finally {
      if (loading) loading.remove();
    }
  }

  function showLoading(name) {
    const el = document.createElement('div');
    el.className = 'loading-overlay';
    el.innerHTML = '<div class="spinner"></div><div></div>';
    el.lastChild.textContent = name + ' 여는 중…';
    document.body.appendChild(el);
    return el;
  }

  function openFiles(list) {
    Array.from(list || []).reduce(function (p, f) { return p.then(function () { return openFile(f); }); }, Promise.resolve());
  }

  /* ---------- 새로 만들기 ---------- */
  function newName(base, ext) {
    let n = 1;
    let name;
    do { name = base + (n > 1 ? ' ' + n : '') + '.' + ext; n++; } while (docs.some(function (d) { return d.name === name; }));
    return name;
  }
  function createNew(kind) {
    if (kind === 'xlsx') addDoc({ name: newName('새 통합 문서', 'xlsx'), format: 'xlsx', book: { sheets: [DV.xlsxIO.newSheet('Sheet1')] }, isNew: true });
    else if (kind === 'hwpx') addDoc({ name: newName('새 한글 문서', 'hwpx'), format: 'hwpx', html: '<p><br></p>', defaultPt: 10, isNew: true });
    else addDoc({ name: newName('새 문서', 'docx'), format: 'docx', html: '<p><br></p>', defaultPt: 10, isNew: true });
  }

  /* ---------- 저장 ---------- */
  function saveOptions(d) {
    if (d.type === 'sheet') return [['xlsx', 'Excel 통합 문서 (.xlsx)'], ['csv', 'CSV UTF-8 — 현재 시트 (.csv)']];
    return [['docx', 'Word 문서 (.docx)'], ['hwpx', '한글 문서 (.hwpx)'], ['html', '웹 페이지 (.html)'], ['txt', '텍스트 (.txt)'], ['pdf', 'PDF (인쇄 → PDF로 저장)']];
  }
  function defaultFormat(d) {
    if (d.type === 'sheet') return d.format === 'csv' ? 'csv' : 'xlsx';
    return d.format === 'docx' ? 'docx' : 'hwpx';
  }

  async function save(d, fmt) {
    if (!d) return;
    fmt = fmt || defaultFormat(d);
    if (fmt === 'pdf') { print(); return; }
    const base = DV.baseName(d.name);
    try {
      let blob;
      if (d.type === 'sheet') {
        blob = fmt === 'csv' ? d.view.exportCSV() : await d.view.exportXlsx();
      } else {
        const page = d.view.page;
        if (fmt === 'docx') blob = await DV.word.htmlToDocx(page, { title: base });
        else if (fmt === 'hwpx') blob = await DV.hwpx.htmlToHwpx(page, { title: base });
        else if (fmt === 'html') {
          const css = 'body{font-family:"맑은 고딕","Malgun Gothic",sans-serif;font-size:10pt;line-height:1.6;max-width:16cm;margin:2cm auto;white-space:pre-wrap;tab-size:4}table{border-collapse:collapse}td,th{border:1px solid #888;padding:3px 6px;vertical-align:top}table.dv-noborder td{border:0}p{margin:0}.page-break{page-break-after:always}img{max-width:100%}';
          blob = new Blob(['<!DOCTYPE html>\n<html lang="ko"><head><meta charset="utf-8"><title>' + DV.esc(base) + '</title><style>' + css + '</style></head><body>' + page.innerHTML + '</body></html>'], { type: 'text/html;charset=utf-8' });
        } else if (fmt === 'txt') {
          blob = new Blob([page.innerText.replace(/\n/g, '\r\n')], { type: 'text/plain;charset=utf-8' });
        }
      }
      const ext = fmt;
      let fname = base + '.' + ext;
      if (fmt === 'csv' && d.type === 'sheet' && d.view.getBook().sheets.length > 1) fname = base + ' - ' + d.view.sheetName() + '.csv';
      DV.download(blob, fname);
      if (fmt === defaultFormat(d) || (d.type === 'doc' && (fmt === 'docx' || fmt === 'hwpx')) || (d.type === 'sheet' && fmt === 'xlsx')) {
        setDirty(d, false);
        if (DV.ext(d.name) !== ext && d.format !== 'hwp' && fmt !== 'csv') { d.name = fname; d.format = ext; renderTabs(); }
      }
      DV.toast('"' + fname + '"(으)로 저장했습니다.', 'ok');
      if (d.format === 'hwp' && fmt === 'hwpx') DV.toast('HWP 원본 형식으로는 저장할 수 없어 .hwpx로 저장했습니다. 한글 2014 이상에서 열 수 있습니다.', 'info', 7000);
    } catch (err) {
      console.error(err);
      DV.toast('저장하지 못했습니다: ' + err.message, 'error', 8000);
    }
  }

  function print() {
    if (!active) return;
    const area = $('#print-area');
    if (active.type === 'sheet') {
      area.innerHTML = active.view.printHTML();
    } else {
      area.innerHTML = '<div class="print-page"></div>';
      area.firstChild.innerHTML = active.view.page.innerHTML;
      area.firstChild.style.fontSize = active.view.page.style.fontSize;
    }
    document.body.classList.add('printing');
    setTimeout(function () {
      window.print();
      setTimeout(function () { document.body.classList.remove('printing'); area.innerHTML = ''; }, 300);
    }, 50);
  }

  /* ---------- 메뉴 ---------- */
  function closeMenus() { document.querySelectorAll('.menu').forEach(function (m) { m.hidden = true; }); }
  function toggleMenu(m) { const was = m.hidden; closeMenus(); m.hidden = !was; }

  $('#btn-open').addEventListener('click', function () { fileInput.click(); });
  fileInput.addEventListener('change', function () { openFiles(fileInput.files); fileInput.value = ''; });
  $('#btn-new').addEventListener('click', function (e) { e.stopPropagation(); toggleMenu($('#menu-new')); });
  document.querySelectorAll('[data-new]').forEach(function (b) {
    b.addEventListener('click', function () { closeMenus(); createNew(b.dataset.new); });
  });
  $('#btn-save').addEventListener('click', function () { save(active); });
  $('#btn-save-as').addEventListener('click', function (e) {
    e.stopPropagation();
    if (!active) return;
    const m = $('#menu-save');
    m.innerHTML = '<div class="menu-title">다른 형식으로 저장</div>' + saveOptions(active).map(function (o) { return '<button data-fmt="' + o[0] + '">' + o[1] + '</button>'; }).join('');
    toggleMenu(m);
  });
  $('#menu-save').addEventListener('click', function (e) {
    const b = e.target.closest('[data-fmt]');
    if (!b) return;
    closeMenus();
    save(active, b.dataset.fmt);
  });
  $('#btn-print').addEventListener('click', print);
  $('#brand').addEventListener('click', function () {
    if (active && active.view.commitEdit) active.view.commitEdit();
    active = null;
    docs.forEach(function (x) { x.host.hidden = true; });
    renderTabs();
  });
  document.addEventListener('click', function (e) { if (!e.target.closest('.menu-wrap')) closeMenus(); });

  const dz = $('#dropzone');
  dz.addEventListener('click', function () { fileInput.click(); });
  dz.addEventListener('keydown', function (e) { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); fileInput.click(); } });

  // 끌어다 놓기
  const overlay = $('#drop-overlay');
  let dragDepth = 0;
  function hasFiles(e) { return e.dataTransfer && Array.from(e.dataTransfer.types || []).indexOf('Files') >= 0; }
  window.addEventListener('dragenter', function (e) {
    if (!hasFiles(e) || (e.target.closest && e.target.closest('.page'))) return;
    dragDepth++; overlay.hidden = false;
  });
  window.addEventListener('dragleave', function (e) { if (!hasFiles(e)) return; dragDepth = Math.max(0, dragDepth - 1); if (!dragDepth) overlay.hidden = true; });
  window.addEventListener('dragover', function (e) { if (hasFiles(e)) e.preventDefault(); });
  window.addEventListener('drop', function (e) {
    dragDepth = 0; overlay.hidden = true;
    if (!hasFiles(e)) return;
    if (e.target.closest && e.target.closest('.page') && Array.from(e.dataTransfer.files).every(function (f) { return /^image\//.test(f.type); })) return;
    e.preventDefault();
    openFiles(e.dataTransfer.files);
  });

  // 단축키
  window.addEventListener('keydown', function (e) {
    const mod = e.ctrlKey || e.metaKey;
    if (!mod) return;
    const k = e.key.toLowerCase();
    if (k === 's') { e.preventDefault(); if (active) save(active, e.shiftKey ? null : null); }
    else if (k === 'o') { e.preventDefault(); fileInput.click(); }
    else if (k === 'p') { if (active) { e.preventDefault(); print(); } }
  });

  window.addEventListener('beforeunload', function (e) {
    if (docs.some(function (d) { return d.dirty; })) { e.preventDefault(); e.returnValue = ''; }
  });

  // 다른 앱에서 "이 앱으로 열기" (PWA 파일 처리) 지원
  if ('launchQueue' in window) {
    window.launchQueue.setConsumer(async function (params) {
      for (const h of params.files || []) openFile(await h.getFile());
    });
  }

  DV.app = { openFile: openFile, createNew: createNew, save: save, docs: docs, get active() { return active; } };
  renderTabs();
})(window.DV);
