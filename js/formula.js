/* 스프레드시트 수식 계산기 (Excel 호환 함수 일부) */
window.DV = window.DV || {};
(function (DV) {
  'use strict';

  function Err(code) { this.e = code; }
  Err.prototype.toString = function () { return this.e; };
  const E = {
    DIV0: new Err('#DIV/0!'), NA: new Err('#N/A'), NAME: new Err('#NAME?'), NUM: new Err('#NUM!'),
    REF: new Err('#REF!'), VALUE: new Err('#VALUE!'), NULL: new Err('#NULL!'), CIRC: new Err('#CIRC!')
  };
  const ERR_BY_CODE = {};
  Object.keys(E).forEach(function (k) { ERR_BY_CODE[E[k].e] = E[k]; });
  function isErr(v) { return v instanceof Err; }

  /* ---------- 주소 도우미 ---------- */
  function colName(c) {
    let s = '';
    c += 1;
    while (c > 0) { const m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - 1) / 26); }
    return s;
  }
  function colIndex(name) {
    let n = 0;
    name = name.toUpperCase();
    for (let i = 0; i < name.length; i++) n = n * 26 + (name.charCodeAt(i) - 64);
    return n - 1;
  }
  function addr(r, c) { return colName(c) + (r + 1); }
  function parseAddr(a) {
    const m = /^\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(a);
    return m ? { r: parseInt(m[2], 10) - 1, c: colIndex(m[1]) } : null;
  }

  /* ---------- 토큰 분리 ---------- */
  const SHEET_RE = /^(?:'((?:[^']|'')+)'|([A-Za-z_\u00A0-\uFFFF][\w.\u00A0-\uFFFF]*))!/;
  const CELL = '(\\$?)([A-Za-z]{1,3})(\\$?)(\\d+)';
  const CELL_RANGE_RE = new RegExp('^' + CELL + '(?::' + CELL + ')?');
  const COL_RANGE_RE = /^(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})(?![\w(])/;
  const ROW_RANGE_RE = /^(\$?)(\d+):(\$?)(\d+)/;

  function tokenize(src) {
    const toks = [];
    let i = 0;
    const n = src.length;
    while (i < n) {
      const ch = src[i];
      const rest = src.slice(i);
      let m;
      if (/\s/.test(ch)) { m = /^\s+/.exec(rest); toks.push({ t: 'ws', text: m[0] }); i += m[0].length; continue; }
      if (ch === '"') {
        let j = i + 1, s = '';
        while (j < n) {
          if (src[j] === '"') { if (src[j + 1] === '"') { s += '"'; j += 2; continue; } break; }
          s += src[j++];
        }
        toks.push({ t: 'str', v: s, text: src.slice(i, j + 1) });
        i = j + 1;
        continue;
      }
      if (ch === '#') {
        m = /^#(DIV\/0!|N\/A|NAME\?|NULL!|NUM!|REF!|VALUE!|SPILL!|CALC!|CIRC!)/i.exec(rest);
        if (m) { toks.push({ t: 'err', v: ERR_BY_CODE[m[0].toUpperCase()] || new Err(m[0].toUpperCase()), text: m[0] }); i += m[0].length; continue; }
      }
      // 시트 접두사가 붙은 참조
      let sheet = null, sheetText = '';
      m = SHEET_RE.exec(rest);
      if (m) { sheet = m[1] != null ? m[1].replace(/''/g, "'") : m[2]; sheetText = m[0]; }
      const after = rest.slice(sheetText.length);
      if (sheet !== null || /[$A-Za-z0-9]/.test(ch)) {
        let rm = CELL_RANGE_RE.exec(after);
        if (rm && !/^[\w(.]/.test(after.slice(rm[0].length))) {
          const a = { cAbs: !!rm[1], c: colIndex(rm[2]), rAbs: !!rm[3], r: parseInt(rm[4], 10) - 1 };
          const b = rm[5] !== undefined ? { cAbs: !!rm[5], c: colIndex(rm[6]), rAbs: !!rm[7], r: parseInt(rm[8], 10) - 1 } : null;
          toks.push({ t: 'ref', kind: 'cell', sheet: sheet, sheetText: sheetText, a: a, b: b, text: sheetText + rm[0] });
          i += sheetText.length + rm[0].length;
          continue;
        }
        rm = COL_RANGE_RE.exec(after);
        if (rm) {
          toks.push({ t: 'ref', kind: 'col', sheet: sheet, sheetText: sheetText, a: { cAbs: !!rm[1], c: colIndex(rm[2]), r: 0 }, b: { cAbs: !!rm[3], c: colIndex(rm[4]), r: 1048575 }, text: sheetText + rm[0] });
          i += sheetText.length + rm[0].length;
          continue;
        }
        rm = ROW_RANGE_RE.exec(after);
        if (rm && (sheet !== null || !/^\d+(\.\d+)?$/.test(after))) {
          toks.push({ t: 'ref', kind: 'row', sheet: sheet, sheetText: sheetText, a: { rAbs: !!rm[1], r: parseInt(rm[2], 10) - 1, c: 0 }, b: { rAbs: !!rm[3], r: parseInt(rm[4], 10) - 1, c: 16383 }, text: sheetText + rm[0] });
          i += sheetText.length + rm[0].length;
          continue;
        }
        if (sheet !== null) { toks.push({ t: 'bad', text: rest }); break; }
      }
      m = /^(\d+\.?\d*(?:[eE][+-]?\d+)?|\.\d+(?:[eE][+-]?\d+)?)/.exec(rest);
      if (m) { toks.push({ t: 'num', v: parseFloat(m[0]), text: m[0] }); i += m[0].length; continue; }
      m = /^[A-Za-z_\u00A0-\uFFFF][\w.\u00A0-\uFFFF]*/.exec(rest);
      if (m) {
        const word = m[0];
        const next = src.slice(i + word.length).replace(/^\s+/, '');
        if (next[0] === '(') toks.push({ t: 'func', v: word.toUpperCase().replace(/^_XLFN\./, ''), text: word });
        else if (/^(TRUE|FALSE)$/i.test(word)) toks.push({ t: 'bool', v: word.toUpperCase() === 'TRUE', text: word });
        else toks.push({ t: 'name', v: word, text: word });
        i += word.length;
        continue;
      }
      m = /^(<=|>=|<>|[-+*/^&=<>%(),;{}:])/.exec(rest);
      if (m) {
        const op = m[0] === ';' ? ',' : m[0];
        toks.push({ t: op === '(' || op === ')' || op === ',' ? op : 'op', v: op, text: m[0] });
        i += m[0].length;
        continue;
      }
      toks.push({ t: 'bad', text: ch });
      i++;
    }
    return toks;
  }

  /* ---------- 구문 분석 ---------- */
  function parse(src) {
    const toks = tokenize(src).filter(function (t) { return t.t !== 'ws'; });
    let p = 0;
    function peek() { return toks[p]; }
    function next() { return toks[p++]; }
    function isOp(v) { const t = peek(); return t && t.t === 'op' && t.v === v; }
    function fail() { throw new SyntaxError('수식 오류'); }

    function compare() {
      let l = concat();
      while (peek() && peek().t === 'op' && /^(=|<>|<|>|<=|>=)$/.test(peek().v)) { const op = next().v; l = { k: 'bin', op: op, l: l, r: concat() }; }
      return l;
    }
    function concat() {
      let l = additive();
      while (isOp('&')) { next(); l = { k: 'bin', op: '&', l: l, r: additive() }; }
      return l;
    }
    function additive() {
      let l = mult();
      while (isOp('+') || isOp('-')) { const op = next().v; l = { k: 'bin', op: op, l: l, r: mult() }; }
      return l;
    }
    function mult() {
      let l = power();
      while (isOp('*') || isOp('/')) { const op = next().v; l = { k: 'bin', op: op, l: l, r: power() }; }
      return l;
    }
    function power() {
      let l = unary();
      while (isOp('^')) { next(); l = { k: 'bin', op: '^', l: l, r: unary() }; }
      return l;
    }
    function unary() {
      if (isOp('-')) { next(); return { k: 'neg', x: unary() }; }
      if (isOp('+')) { next(); return unary(); }
      return postfix();
    }
    function postfix() {
      let x = primary();
      while (isOp('%')) { next(); x = { k: 'pct', x: x }; }
      return x;
    }
    function primary() {
      const t = next();
      if (!t) fail();
      switch (t.t) {
        case 'num': return { k: 'num', v: t.v };
        case 'str': return { k: 'str', v: t.v };
        case 'bool': return { k: 'bool', v: t.v };
        case 'err': return { k: 'err', v: t.v };
        case 'ref': return { k: 'ref', sheet: t.sheet, kind: t.kind, a: t.a, b: t.b };
        case 'name': return { k: 'name', v: t.v };
        case 'func': {
          if (!peek() || peek().t !== '(') fail();
          next();
          const args = [];
          if (peek() && peek().t === ')') { next(); return { k: 'fn', name: t.v, args: args }; }
          for (;;) {
            if (peek() && (peek().t === ',' || peek().t === ')')) args.push({ k: 'empty' });
            else args.push(compare());
            const s = next();
            if (!s) fail();
            if (s.t === ')') break;
            if (s.t !== ',') fail();
          }
          return { k: 'fn', name: t.v, args: args };
        }
        case '(': {
          const x = compare();
          const c = next();
          if (!c || c.t !== ')') fail();
          return x;
        }
        default: fail();
      }
    }
    const ast = compare();
    if (p < toks.length) fail();
    return ast;
  }

  const cache = new Map();
  function compile(src) {
    let a = cache.get(src);
    if (a === undefined) {
      try { a = parse(src); } catch (e) { a = { k: 'err', v: E.NAME, syntax: true }; }
      if (cache.size > 5000) cache.clear();
      cache.set(src, a);
    }
    return a;
  }

  /* ---------- 값 변환 ---------- */
  function Range(ctx, sheet, r1, c1, r2, c2) {
    this.ctx = ctx; this.sheet = sheet;
    this.r1 = Math.min(r1, r2); this.c1 = Math.min(c1, c2); this.r2 = Math.max(r1, r2); this.c2 = Math.max(c1, c2);
  }
  Range.prototype.clip = function () {
    const ext = this.ctx.extent(this.sheet);
    return { r2: Math.min(this.r2, ext.rows - 1), c2: Math.min(this.c2, ext.cols - 1) };
  };
  Range.prototype.rows = function () { return this.r2 - this.r1 + 1; };
  Range.prototype.cols = function () { return this.c2 - this.c1 + 1; };
  Range.prototype.get = function (i, j) { return this.ctx.cell(this.sheet, this.r1 + i, this.c1 + j); };
  Range.prototype.each = function (fn) {
    const cl = this.clip();
    for (let r = this.r1; r <= cl.r2; r++) for (let c = this.c1; c <= cl.c2; c++) fn(this.ctx.cell(this.sheet, r, c), r - this.r1, c - this.c1);
  };
  Range.prototype.list = function () { const a = []; this.each(function (v) { a.push(v); }); return a; };

  function numFromStr(s) {
    const t = String(s).trim().replace(/,/g, '');
    if (t === '') return null;
    let m = /^([-+]?\d*\.?\d+(?:[eE][-+]?\d+)?)(%?)$/.exec(t);
    if (m) return parseFloat(m[1]) / (m[2] ? 100 : 1);
    m = /^(\d{4})[-./](\d{1,2})[-./](\d{1,2})$/.exec(t);
    if (m) return dateSerial(+m[1], +m[2], +m[3]);
    return null;
  }
  function toNum(v) {
    if (isErr(v)) return v;
    if (typeof v === 'number') return v;
    if (typeof v === 'boolean') return v ? 1 : 0;
    if (v == null || v === '') return 0;
    if (v instanceof Range) return v.rows() === 1 && v.cols() === 1 ? toNum(v.get(0, 0)) : E.VALUE;
    const n = numFromStr(v);
    return n == null ? E.VALUE : n;
  }
  function fmtNum(n) {
    if (!isFinite(n)) return '#NUM!';
    if (Number.isInteger(n)) return String(n);
    let s = String(parseFloat(n.toPrecision(15)));
    if (/e/.test(s)) s = n.toExponential(10).replace(/\.?0+e/, 'E');
    return s;
  }
  function toStr(v) {
    if (isErr(v)) return v;
    if (v == null) return '';
    if (typeof v === 'number') return fmtNum(v);
    if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
    if (v instanceof Range) return v.rows() === 1 && v.cols() === 1 ? toStr(v.get(0, 0)) : E.VALUE;
    return String(v);
  }
  function toBool(v) {
    if (isErr(v)) return v;
    if (typeof v === 'boolean') return v;
    if (typeof v === 'number') return v !== 0;
    if (v == null || v === '') return false;
    if (v instanceof Range) return v.rows() === 1 && v.cols() === 1 ? toBool(v.get(0, 0)) : E.VALUE;
    const u = String(v).toUpperCase();
    if (u === 'TRUE') return true;
    if (u === 'FALSE') return false;
    return E.VALUE;
  }
  function scalar(v) {
    if (v instanceof Range) return v.rows() === 1 && v.cols() === 1 ? v.get(0, 0) : E.VALUE;
    return v;
  }
  function typeRank(v) { return typeof v === 'number' ? 1 : typeof v === 'string' ? 2 : typeof v === 'boolean' ? 3 : 0; }
  function compareVals(a, b) {
    if (a == null) a = typeof b === 'string' ? '' : typeof b === 'boolean' ? false : 0;
    if (b == null) b = typeof a === 'string' ? '' : typeof a === 'boolean' ? false : 0;
    const ta = typeRank(a), tb = typeRank(b);
    if (ta !== tb) return ta - tb;
    if (typeof a === 'string') { const x = a.toLowerCase(), y = b.toLowerCase(); return x < y ? -1 : x > y ? 1 : 0; }
    if (typeof a === 'boolean') return (a ? 1 : 0) - (b ? 1 : 0);
    return a < b ? -1 : a > b ? 1 : 0;
  }

  /* ---------- 날짜 ---------- */
  const EPOCH = Date.UTC(1899, 11, 30);
  function dateSerial(y, m, d) { return Math.round((Date.UTC(y, m - 1, d) - EPOCH) / 86400000); }
  function serialToDate(s) { return new Date(EPOCH + Math.floor(s) * 86400000); }

  /* ---------- 평가 ---------- */
  function evaluate(ast, ctx) {
    switch (ast.k) {
      case 'num': case 'str': case 'bool': return ast.v;
      case 'err': return ast.v;
      case 'empty': return null;
      case 'name': return E.NAME;
      case 'ref': {
        const rg = refRange(ast, ctx);
        if (isErr(rg)) return rg;
        if (rg.r1 === rg.r2 && rg.c1 === rg.c2) return rg.get(0, 0);
        return rg;
      }
      case 'neg': { const v = toNum(scalar(evaluate(ast.x, ctx))); return isErr(v) ? v : -v; }
      case 'pct': { const v = toNum(scalar(evaluate(ast.x, ctx))); return isErr(v) ? v : v / 100; }
      case 'bin': return binop(ast.op, scalar(evaluate(ast.l, ctx)), scalar(evaluate(ast.r, ctx)));
      case 'fn': return callFn(ast, ctx);
    }
    return E.VALUE;
  }

  function refRange(ast, ctx) {
    const sheet = ast.sheet;
    if (sheet != null && !ctx.hasSheet(sheet)) return E.REF;
    const b = ast.b || ast.a;
    return new Range(ctx, sheet, ast.a.r, ast.a.c, b.r, b.c);
  }

  function binop(op, a, b) {
    if (isErr(a)) return a;
    if (isErr(b)) return b;
    switch (op) {
      case '&': { const x = toStr(a), y = toStr(b); return isErr(x) ? x : isErr(y) ? y : x + y; }
      case '=': return compareVals(a, b) === 0;
      case '<>': return compareVals(a, b) !== 0;
      case '<': return compareVals(a, b) < 0;
      case '>': return compareVals(a, b) > 0;
      case '<=': return compareVals(a, b) <= 0;
      case '>=': return compareVals(a, b) >= 0;
    }
    const x = toNum(a), y = toNum(b);
    if (isErr(x)) return x;
    if (isErr(y)) return y;
    switch (op) {
      case '+': return x + y;
      case '-': return x - y;
      case '*': return x * y;
      case '/': return y === 0 ? E.DIV0 : x / y;
      case '^': { const r = Math.pow(x, y); return isFinite(r) ? r : E.NUM; }
    }
    return E.VALUE;
  }

  function argVal(ast, ctx) {
    if (ast.k === 'ref') { const rg = refRange(ast, ctx); return rg; }
    return evaluate(ast, ctx);
  }

  /* 집계 함수용: 범위 안의 숫자만, 직접 입력값은 변환 */
  function nums(args, ctx, opt) {
    opt = opt || {};
    const out = [];
    for (const a of args) {
      const v = argVal(a, ctx);
      if (v instanceof Range) {
        let err = null;
        v.each(function (x) {
          if (err) return;
          if (isErr(x)) err = x;
          else if (typeof x === 'number') out.push(x);
          else if (opt.countText && x != null && x !== '') out.push(opt.textAs || 0);
        });
        if (err) return err;
      } else if (isErr(v)) return v;
      else if (v == null) { if (a.k !== 'empty') out.push(0); }
      else {
        const n = toNum(v);
        if (isErr(n)) return n;
        out.push(n);
      }
    }
    return out;
  }
  function flat(args, ctx) {
    const out = [];
    for (const a of args) {
      const v = argVal(a, ctx);
      if (v instanceof Range) v.each(function (x) { out.push(x); }); else out.push(v);
    }
    return out;
  }

  function makeCrit(c) {
    if (typeof c === 'number' || typeof c === 'boolean') return function (v) { return compareVals(v, c) === 0 && v != null; };
    const s = c == null ? '' : String(c);
    const m = /^(<=|>=|<>|<|>|=)?(.*)$/s.exec(s);
    const op = m[1] || '=';
    const raw = m[2];
    const n = numFromStr(raw);
    let target = n != null ? n : raw;
    if (/^(TRUE|FALSE)$/i.test(raw)) target = raw.toUpperCase() === 'TRUE';
    const wild = typeof target === 'string' && /[*?]/.test(target) && (op === '=' || op === '<>');
    let re = null;
    if (wild) re = new RegExp('^' + target.replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/~\*/g, '\u0001').replace(/~\?/g, '\u0002').replace(/\*/g, '.*').replace(/\?/g, '.').replace(/\u0001/g, '\\*').replace(/\u0002/g, '\\?') + '$', 'i');
    return function (v) {
      if (isErr(v)) return false;
      if (re) { const ok = v != null && re.test(toStr(v)); return op === '=' ? ok : !ok; }
      if (raw === '' && op === '=') return v == null || v === '';
      if (raw === '' && op === '<>') return !(v == null || v === '');
      if (typeof target === 'number' && typeof v !== 'number') {
        if (op === '<>') return true;
        return false;
      }
      if (typeof target === 'string' && typeof v !== 'string') return op === '<>';
      const cmp = compareVals(v, target);
      switch (op) {
        case '=': return cmp === 0;
        case '<>': return cmp !== 0;
        case '<': return cmp < 0;
        case '>': return cmp > 0;
        case '<=': return cmp <= 0;
        case '>=': return cmp >= 0;
      }
      return false;
    };
  }

  function round(x, d, mode) {
    const f = Math.pow(10, d);
    const v = x * f;
    let r;
    const eps = 1e-9 * Math.max(1, Math.abs(v));
    if (mode === 'up') r = v >= 0 ? Math.ceil(v - eps) : -Math.ceil(-v - eps);
    else if (mode === 'down') r = v >= 0 ? Math.floor(v + eps) : -Math.floor(-v + eps);
    else r = v >= 0 ? Math.floor(v + 0.5 + eps) : -Math.floor(-v + 0.5 + eps);
    return r / f;
  }

  function lookupRange(v) { return v instanceof Range ? v : null; }

  function matchIndex(needle, list, type) {
    if (type === 0) {
      const crit = typeof needle === 'string' && /[*?]/.test(needle) ? makeCrit(needle) : null;
      for (let i = 0; i < list.length; i++) {
        if (crit ? crit(list[i]) : (list[i] != null && compareVals(list[i], needle) === 0)) return i;
      }
      return -1;
    }
    let best = -1;
    for (let i = 0; i < list.length; i++) {
      if (list[i] == null) continue;
      const c = compareVals(list[i], needle);
      if (type === 1) { if (c <= 0) best = i; else break; }
      else { if (c >= 0) best = i; else break; }
    }
    return best;
  }

  const FN = {};
  function agg(name, fn, opt) {
    FN[name] = function (args, ctx) { const a = nums(args, ctx, opt); return isErr(a) ? a : fn(a); };
  }
  agg('SUM', function (a) { return a.reduce(function (s, x) { return s + x; }, 0); });
  agg('PRODUCT', function (a) { return a.length ? a.reduce(function (s, x) { return s * x; }, 1) : 0; });
  agg('AVERAGE', function (a) { return a.length ? a.reduce(function (s, x) { return s + x; }, 0) / a.length : E.DIV0; });
  agg('MIN', function (a) { return a.length ? Math.min.apply(null, a) : 0; });
  agg('MAX', function (a) { return a.length ? Math.max.apply(null, a) : 0; });
  agg('COUNT', function (a) { return a.length; });
  agg('MEDIAN', function (a) {
    if (!a.length) return E.NUM;
    a.sort(function (x, y) { return x - y; });
    const m = a.length >> 1;
    return a.length % 2 ? a[m] : (a[m - 1] + a[m]) / 2;
  });
  function variance(a, sample) {
    if (a.length < (sample ? 2 : 1)) return E.DIV0;
    const mean = a.reduce(function (s, x) { return s + x; }, 0) / a.length;
    return a.reduce(function (s, x) { return s + (x - mean) * (x - mean); }, 0) / (a.length - (sample ? 1 : 0));
  }
  agg('VAR', function (a) { return variance(a, true); });
  FN['VAR.S'] = FN.VAR;
  agg('VAR.P', function (a) { return variance(a, false); });
  FN.VARP = FN['VAR.P'];
  agg('STDEV', function (a) { const v = variance(a, true); return isErr(v) ? v : Math.sqrt(v); });
  FN['STDEV.S'] = FN.STDEV;
  agg('STDEV.P', function (a) { const v = variance(a, false); return isErr(v) ? v : Math.sqrt(v); });
  FN.STDEVP = FN['STDEV.P'];

  FN.COUNTA = function (args, ctx) {
    let n = 0;
    flat(args, ctx).forEach(function (v) { if (v != null && v !== '') n++; });
    return n;
  };
  FN.COUNTBLANK = function (args, ctx) {
    let n = 0;
    args.forEach(function (a) { const v = argVal(a, ctx); if (v instanceof Range) { for (let i = 0; i < v.rows(); i++) for (let j = 0; j < v.cols(); j++) { const x = v.get(i, j); if (x == null || x === '') n++; } } });
    return n;
  };
  function largeSmall(args, ctx, large) {
    const a = nums([args[0]], ctx);
    if (isErr(a)) return a;
    const k = toNum(scalar(evaluate(args[1], ctx)));
    if (isErr(k)) return k;
    if (k < 1 || k > a.length) return E.NUM;
    a.sort(function (x, y) { return large ? y - x : x - y; });
    return a[Math.floor(k) - 1];
  }
  FN.LARGE = function (args, ctx) { return largeSmall(args, ctx, true); };
  FN.SMALL = function (args, ctx) { return largeSmall(args, ctx, false); };
  FN.RANK = function (args, ctx) {
    const x = toNum(scalar(evaluate(args[0], ctx)));
    if (isErr(x)) return x;
    const a = nums([args[1]], ctx);
    if (isErr(a)) return a;
    const asc = args[2] ? toBool(scalar(evaluate(args[2], ctx))) === true : false;
    if (a.indexOf(x) < 0) return E.NA;
    return 1 + a.filter(function (v) { return asc ? v < x : v > x; }).length;
  };
  FN['RANK.EQ'] = FN.RANK;

  function ifAgg(args, ctx, kind) {
    const rg = argVal(args[0], ctx);
    if (!(rg instanceof Range)) return E.VALUE;
    const crit = makeCrit(scalar(evaluate(args[1], ctx)));
    const sr = args[2] ? argVal(args[2], ctx) : rg;
    if (!(sr instanceof Range)) return E.VALUE;
    let sum = 0, cnt = 0;
    const cl = rg.clip();
    for (let r = 0; r <= cl.r2 - rg.r1; r++) {
      for (let c = 0; c <= cl.c2 - rg.c1; c++) {
        if (crit(rg.get(r, c))) {
          cnt++;
          if (kind !== 'count') { const v = sr.get(r, c); if (typeof v === 'number') sum += v; }
        }
      }
    }
    if (kind === 'count') return cnt;
    if (kind === 'avg') return cnt ? sum / cnt : E.DIV0;
    return sum;
  }
  FN.SUMIF = function (a, c) { return ifAgg(a, c, 'sum'); };
  FN.COUNTIF = function (a, c) { return ifAgg(a, c, 'count'); };
  FN.AVERAGEIF = function (a, c) { return ifAgg(a, c, 'avg'); };
  function ifsAgg(args, ctx, kind) {
    let start = kind === 'count' ? 0 : 1;
    const sr = kind === 'count' ? null : argVal(args[0], ctx);
    const pairs = [];
    for (let i = start; i + 1 < args.length; i += 2) {
      const rg = argVal(args[i], ctx);
      if (!(rg instanceof Range)) return E.VALUE;
      pairs.push([rg, makeCrit(scalar(evaluate(args[i + 1], ctx)))]);
    }
    if (!pairs.length) return E.VALUE;
    const base = pairs[0][0];
    const cl = base.clip();
    let sum = 0, cnt = 0;
    for (let r = 0; r <= cl.r2 - base.r1; r++) {
      for (let c = 0; c <= cl.c2 - base.c1; c++) {
        if (pairs.every(function (p) { return p[1](p[0].get(r, c)); })) {
          cnt++;
          if (sr) { const v = sr.get(r, c); if (typeof v === 'number') sum += v; }
        }
      }
    }
    if (kind === 'count') return cnt;
    if (kind === 'avg') return cnt ? sum / cnt : E.DIV0;
    return sum;
  }
  FN.SUMIFS = function (a, c) { return ifsAgg(a, c, 'sum'); };
  FN.COUNTIFS = function (a, c) { return ifsAgg(a, c, 'count'); };
  FN.AVERAGEIFS = function (a, c) { return ifsAgg(a, c, 'avg'); };
  FN.SUMPRODUCT = function (args, ctx) {
    const rs = args.map(function (a) { return argVal(a, ctx); });
    if (!rs.length || !rs.every(function (r) { return r instanceof Range; })) return E.VALUE;
    const R = rs[0].rows(), C = rs[0].cols();
    if (!rs.every(function (r) { return r.rows() === R && r.cols() === C; })) return E.VALUE;
    let s = 0;
    for (let i = 0; i < R; i++) for (let j = 0; j < C; j++) {
      let p = 1;
      for (const r of rs) { const v = r.get(i, j); p *= typeof v === 'number' ? v : 0; }
      s += p;
    }
    return s;
  };

  // 논리
  FN.IF = function (args, ctx) {
    const c = toBool(scalar(evaluate(args[0], ctx)));
    if (isErr(c)) return c;
    if (c) return args.length > 1 ? scalar(evaluate(args[1], ctx)) : true;
    return args.length > 2 ? scalar(evaluate(args[2], ctx)) : false;
  };
  FN.IFS = function (args, ctx) {
    for (let i = 0; i + 1 < args.length; i += 2) {
      const c = toBool(scalar(evaluate(args[i], ctx)));
      if (isErr(c)) return c;
      if (c) return scalar(evaluate(args[i + 1], ctx));
    }
    return E.NA;
  };
  FN.IFERROR = function (args, ctx) { const v = scalar(evaluate(args[0], ctx)); return isErr(v) ? scalar(evaluate(args[1], ctx)) : v; };
  FN.IFNA = function (args, ctx) { const v = scalar(evaluate(args[0], ctx)); return v === E.NA || (isErr(v) && v.e === '#N/A') ? scalar(evaluate(args[1], ctx)) : v; };
  function logic(args, ctx, kind) {
    const vals = flat(args, ctx);
    let res = kind === 'and';
    let seen = false;
    let x = false;
    for (const v of vals) {
      if (isErr(v)) return v;
      if (v == null || typeof v === 'string') continue;
      const b = toBool(v);
      seen = true;
      if (kind === 'and') res = res && b;
      else if (kind === 'or') res = res || b;
      else if (b) x = !x;
    }
    if (!seen) return E.VALUE;
    return kind === 'xor' ? x : res;
  }
  FN.AND = function (a, c) { return logic(a, c, 'and'); };
  FN.OR = function (a, c) { return logic(a, c, 'or'); };
  FN.XOR = function (a, c) { return logic(a, c, 'xor'); };
  FN.NOT = function (args, ctx) { const b = toBool(scalar(evaluate(args[0], ctx))); return isErr(b) ? b : !b; };
  FN.TRUE = function () { return true; };
  FN.FALSE = function () { return false; };
  FN.CHOOSE = function (args, ctx) {
    const i = toNum(scalar(evaluate(args[0], ctx)));
    if (isErr(i)) return i;
    if (i < 1 || i >= args.length) return E.VALUE;
    return scalar(evaluate(args[Math.floor(i)], ctx));
  };

  // 수학
  function num1(fn) {
    return function (args, ctx) {
      const v = [];
      for (const a of args) { const x = toNum(scalar(evaluate(a, ctx))); if (isErr(x)) return x; v.push(x); }
      const r = fn.apply(null, v);
      return typeof r === 'number' && !isFinite(r) ? E.NUM : r;
    };
  }
  FN.ROUND = num1(function (x, d) { return round(x, Math.trunc(d || 0)); });
  FN.ROUNDUP = num1(function (x, d) { return round(x, Math.trunc(d || 0), 'up'); });
  FN.ROUNDDOWN = num1(function (x, d) { return round(x, Math.trunc(d || 0), 'down'); });
  FN.TRUNC = num1(function (x, d) { return round(x, Math.trunc(d || 0), 'down'); });
  FN.INT = num1(Math.floor);
  FN.ABS = num1(Math.abs);
  FN.SQRT = num1(function (x) { return x < 0 ? E.NUM : Math.sqrt(x); });
  FN.POWER = num1(Math.pow);
  FN.MOD = num1(function (a, b) { if (b === 0) return E.DIV0; const r = a - b * Math.floor(a / b); return r; });
  FN.EXP = num1(Math.exp);
  FN.LN = num1(function (x) { return x <= 0 ? E.NUM : Math.log(x); });
  FN.LOG10 = num1(function (x) { return x <= 0 ? E.NUM : Math.log10(x); });
  FN.LOG = num1(function (x, b) { if (x <= 0) return E.NUM; return Math.log(x) / Math.log(b == null ? 10 : b); });
  FN.PI = function () { return Math.PI; };
  FN.SIGN = num1(Math.sign);
  FN.RAND = function () { return Math.random(); };
  FN.RANDBETWEEN = num1(function (a, b) { return Math.floor(Math.random() * (Math.floor(b) - Math.ceil(a) + 1)) + Math.ceil(a); });
  FN.CEILING = num1(function (x, s) { s = s == null ? 1 : s; return s === 0 ? 0 : Math.ceil(x / s - 1e-12) * s; });
  FN.FLOOR = num1(function (x, s) { s = s == null ? 1 : s; return s === 0 ? E.DIV0 : Math.floor(x / s + 1e-12) * s; });
  FN['CEILING.MATH'] = FN.CEILING;
  FN['FLOOR.MATH'] = FN.FLOOR;
  FN.QUOTIENT = num1(function (a, b) { return b === 0 ? E.DIV0 : Math.trunc(a / b); });

  // 문자열
  function str1(fn) {
    return function (args, ctx) {
      const v = [];
      for (let i = 0; i < args.length; i++) {
        const raw = scalar(evaluate(args[i], ctx));
        if (isErr(raw)) return raw;
        v.push(raw);
      }
      return fn.apply(null, v);
    };
  }
  FN.LEN = str1(function (s) { return toStr(s).length; });
  FN.LEFT = str1(function (s, n) { n = n == null ? 1 : toNum(n); return isErr(n) ? n : n < 0 ? E.VALUE : toStr(s).slice(0, n); });
  FN.RIGHT = str1(function (s, n) { n = n == null ? 1 : toNum(n); if (isErr(n)) return n; if (n < 0) return E.VALUE; const t = toStr(s); return n === 0 ? '' : t.slice(-n); });
  FN.MID = str1(function (s, a, n) { a = toNum(a); n = toNum(n); if (isErr(a)) return a; if (isErr(n)) return n; if (a < 1 || n < 0) return E.VALUE; return toStr(s).substr(a - 1, n); });
  FN.UPPER = str1(function (s) { return toStr(s).toUpperCase(); });
  FN.LOWER = str1(function (s) { return toStr(s).toLowerCase(); });
  FN.PROPER = str1(function (s) { return toStr(s).toLowerCase().replace(/(^|[^a-z])([a-z])/g, function (m, a, b) { return a + b.toUpperCase(); }); });
  FN.TRIM = str1(function (s) { return toStr(s).trim().replace(/ {2,}/g, ' '); });
  FN.REPT = str1(function (s, n) { n = toNum(n); return isErr(n) ? n : toStr(s).repeat(Math.max(0, Math.floor(n))); });
  FN.EXACT = str1(function (a, b) { return toStr(a) === toStr(b); });
  FN.CONCATENATE = function (args, ctx) {
    let s = '';
    for (const a of args) { const v = scalar(evaluate(a, ctx)); if (isErr(v)) return v; s += toStr(v); }
    return s;
  };
  FN.CONCAT = function (args, ctx) {
    let s = '';
    for (const v of flat(args, ctx)) { if (isErr(v)) return v; s += toStr(v); }
    return s;
  };
  FN.TEXTJOIN = function (args, ctx) {
    const d = toStr(scalar(evaluate(args[0], ctx)));
    const skip = toBool(scalar(evaluate(args[1], ctx)));
    const parts = [];
    for (const v of flat(args.slice(2), ctx)) { if (isErr(v)) return v; if (skip && (v == null || v === '')) continue; parts.push(toStr(v)); }
    return parts.join(d);
  };
  FN.SUBSTITUTE = str1(function (s, o, n, k) {
    s = toStr(s); o = toStr(o); n = toStr(n);
    if (!o) return s;
    if (k == null) return s.split(o).join(n);
    let idx = -1, count = 0;
    k = toNum(k);
    while ((idx = s.indexOf(o, idx + 1)) >= 0) { if (++count === k) return s.slice(0, idx) + n + s.slice(idx + o.length); }
    return s;
  });
  FN.REPLACE = str1(function (s, a, n, t) { s = toStr(s); a = toNum(a); n = toNum(n); return s.slice(0, a - 1) + toStr(t) + s.slice(a - 1 + n); });
  FN.FIND = str1(function (f, s, st) { st = st == null ? 1 : toNum(st); const i = toStr(s).indexOf(toStr(f), st - 1); return i < 0 ? E.VALUE : i + 1; });
  FN.SEARCH = str1(function (f, s, st) {
    st = st == null ? 1 : toNum(st);
    const re = new RegExp(toStr(f).replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/\*/g, '.*').replace(/\?/g, '.'), 'i');
    const sub = toStr(s).slice(st - 1);
    const m = re.exec(sub);
    return m ? m.index + st : E.VALUE;
  });
  FN.VALUE = str1(function (s) { if (typeof s === 'number') return s; const n = numFromStr(toStr(s)); return n == null ? E.VALUE : n; });
  FN.TEXT = str1(function (v, f) {
    try { return window.XLSX ? XLSX.SSF.format(toStr(f), typeof v === 'string' && numFromStr(v) != null ? numFromStr(v) : v) : toStr(v); } catch (e) { return E.VALUE; }
  });
  FN.CHAR = num1(function (n) { return String.fromCharCode(n); });
  FN.CODE = str1(function (s) { s = toStr(s); return s ? s.charCodeAt(0) : E.VALUE; });
  FN.N = str1(function (v) { return typeof v === 'number' ? v : typeof v === 'boolean' ? (v ? 1 : 0) : 0; });
  FN.T = str1(function (v) { return typeof v === 'string' ? v : ''; });

  // 정보
  FN.ISBLANK = function (args, ctx) { const v = scalar(argVal(args[0], ctx)); return v == null; };
  FN.ISNUMBER = function (args, ctx) { return typeof scalar(argVal(args[0], ctx)) === 'number'; };
  FN.ISTEXT = function (args, ctx) { return typeof scalar(argVal(args[0], ctx)) === 'string'; };
  FN.ISLOGICAL = function (args, ctx) { return typeof scalar(argVal(args[0], ctx)) === 'boolean'; };
  FN.ISERROR = function (args, ctx) { return isErr(scalar(evaluate(args[0], ctx))); };
  FN.ISERR = function (args, ctx) { const v = scalar(evaluate(args[0], ctx)); return isErr(v) && v.e !== '#N/A'; };
  FN.ISNA = function (args, ctx) { const v = scalar(evaluate(args[0], ctx)); return isErr(v) && v.e === '#N/A'; };
  FN.ISEVEN = num1(function (n) { return Math.floor(Math.abs(n)) % 2 === 0; });
  FN.ISODD = num1(function (n) { return Math.floor(Math.abs(n)) % 2 === 1; });
  FN.NA = function () { return E.NA; };

  // 날짜
  function today() { const d = new Date(); return dateSerial(d.getFullYear(), d.getMonth() + 1, d.getDate()); }
  FN.TODAY = function () { return today(); };
  FN.NOW = function () { const d = new Date(); return today() + (d.getHours() * 3600 + d.getMinutes() * 60 + d.getSeconds()) / 86400; };
  FN.DATE = num1(function (y, m, d) { if (y < 100) y += 1900; return Math.round((Date.UTC(y, m - 1, d) - EPOCH) / 86400000); });
  FN.TIME = num1(function (h, m, s) { return ((h * 3600 + m * 60 + s) % 86400) / 86400; });
  FN.YEAR = num1(function (s) { return serialToDate(s).getUTCFullYear(); });
  FN.MONTH = num1(function (s) { return serialToDate(s).getUTCMonth() + 1; });
  FN.DAY = num1(function (s) { return serialToDate(s).getUTCDate(); });
  FN.HOUR = num1(function (s) { return Math.floor(((s % 1) * 24) + 1e-9); });
  FN.MINUTE = num1(function (s) { return Math.floor(((s * 1440) % 60) + 1e-9); });
  FN.SECOND = num1(function (s) { return Math.round((s * 86400) % 60) % 60; });
  FN.WEEKDAY = num1(function (s, t) {
    const d = serialToDate(s).getUTCDay();
    t = t || 1;
    if (t === 2) return d === 0 ? 7 : d;
    if (t === 3) return d === 0 ? 6 : d - 1;
    return d + 1;
  });
  FN.EDATE = num1(function (s, m) { const d = serialToDate(s); return Math.round((Date.UTC(d.getUTCFullYear(), d.getUTCMonth() + m, d.getUTCDate()) - EPOCH) / 86400000); });
  FN.EOMONTH = num1(function (s, m) { const d = serialToDate(s); return Math.round((Date.UTC(d.getUTCFullYear(), d.getUTCMonth() + m + 1, 0) - EPOCH) / 86400000); });
  FN.DATEDIF = str1(function (a, b, u) {
    a = toNum(a); b = toNum(b);
    if (isErr(a)) return a; if (isErr(b)) return b;
    if (a > b) return E.NUM;
    const d1 = serialToDate(a), d2 = serialToDate(b);
    u = toStr(u).toUpperCase();
    let months = (d2.getUTCFullYear() - d1.getUTCFullYear()) * 12 + d2.getUTCMonth() - d1.getUTCMonth();
    if (d2.getUTCDate() < d1.getUTCDate()) months--;
    if (u === 'D') return Math.floor(b) - Math.floor(a);
    if (u === 'M') return months;
    if (u === 'Y') return Math.floor(months / 12);
    if (u === 'YM') return months % 12;
    return E.NUM;
  });
  FN.DAYS = num1(function (e, s) { return Math.floor(e) - Math.floor(s); });

  // 찾기/참조
  FN.VLOOKUP = function (args, ctx) { return vhLookup(args, ctx, true); };
  FN.HLOOKUP = function (args, ctx) { return vhLookup(args, ctx, false); };
  function vhLookup(args, ctx, vertical) {
    const needle = scalar(evaluate(args[0], ctx));
    if (isErr(needle)) return needle;
    const rg = lookupRange(argVal(args[1], ctx));
    if (!rg) return E.VALUE;
    const idx = toNum(scalar(evaluate(args[2], ctx)));
    if (isErr(idx)) return idx;
    const approx = args.length > 3 ? toBool(scalar(evaluate(args[3], ctx))) : true;
    if (idx < 1 || idx > (vertical ? rg.cols() : rg.rows())) return E.REF;
    const cl = rg.clip();
    const n = vertical ? cl.r2 - rg.r1 + 1 : cl.c2 - rg.c1 + 1;
    const list = [];
    for (let i = 0; i < n; i++) list.push(vertical ? rg.get(i, 0) : rg.get(0, i));
    const pos = matchIndex(needle, list, approx === true ? 1 : 0);
    if (pos < 0) return E.NA;
    return vertical ? rg.get(pos, Math.floor(idx) - 1) : rg.get(Math.floor(idx) - 1, pos);
  }
  FN.MATCH = function (args, ctx) {
    const needle = scalar(evaluate(args[0], ctx));
    if (isErr(needle)) return needle;
    const rg = lookupRange(argVal(args[1], ctx));
    if (!rg) return E.NA;
    const type = args.length > 2 ? toNum(scalar(evaluate(args[2], ctx))) : 1;
    const list = rg.list();
    const pos = matchIndex(needle, list, type > 0 ? 1 : type < 0 ? -1 : 0);
    return pos < 0 ? E.NA : pos + 1;
  };
  FN.INDEX = function (args, ctx) {
    const rg = lookupRange(argVal(args[0], ctx));
    if (!rg) return E.VALUE;
    let r = args.length > 1 ? toNum(scalar(evaluate(args[1], ctx))) : 1;
    let c = args.length > 2 ? toNum(scalar(evaluate(args[2], ctx))) : 1;
    if (isErr(r)) return r; if (isErr(c)) return c;
    if (rg.rows() === 1 && args.length === 2) { c = r; r = 1; }
    if (r < 1 || c < 1 || r > rg.rows() || c > rg.cols()) return E.REF;
    return rg.get(Math.floor(r) - 1, Math.floor(c) - 1);
  };
  FN.XLOOKUP = function (args, ctx) {
    const needle = scalar(evaluate(args[0], ctx));
    const lr = lookupRange(argVal(args[1], ctx)), rr = lookupRange(argVal(args[2], ctx));
    if (!lr || !rr) return E.VALUE;
    const list = lr.list();
    const pos = matchIndex(needle, list, 0);
    if (pos < 0) return args.length > 3 ? scalar(evaluate(args[3], ctx)) : E.NA;
    return lr.rows() === 1 ? rr.get(0, pos) : rr.get(pos, 0);
  };
  FN.ROW = function (args, ctx) { if (!args.length || args[0].k === 'empty') return ctx.curRow + 1; const r = argVal(args[0], ctx); return r instanceof Range ? r.r1 + 1 : E.VALUE; };
  FN.COLUMN = function (args, ctx) { if (!args.length || args[0].k === 'empty') return ctx.curCol + 1; const r = argVal(args[0], ctx); return r instanceof Range ? r.c1 + 1 : E.VALUE; };
  FN.ROWS = function (args, ctx) { const r = argVal(args[0], ctx); return r instanceof Range ? r.rows() : 1; };
  FN.COLUMNS = function (args, ctx) { const r = argVal(args[0], ctx); return r instanceof Range ? r.cols() : 1; };

  function callFn(ast, ctx) {
    const f = FN[ast.name];
    if (!f) return E.NAME;
    try {
      const v = f(ast.args, ctx);
      return v === undefined ? E.VALUE : v;
    } catch (e) {
      if (e && e.dvCircular) throw e;
      return E.VALUE;
    }
  }

  /* ---------- 참조 이동/조정 ---------- */
  function partText(p, kind) {
    if (kind === 'col') return (p.cAbs ? '$' : '') + colName(p.c);
    if (kind === 'row') return (p.rAbs ? '$' : '') + (p.r + 1);
    return (p.cAbs ? '$' : '') + colName(p.c) + (p.rAbs ? '$' : '') + (p.r + 1);
  }
  function refText(t) {
    return t.sheetText + partText(t.a, t.kind) + (t.b ? ':' + partText(t.b, t.kind) : '');
  }
  function rebuild(toks) { return toks.map(function (t) { return t.t === 'ref' ? (t.bad ? t.sheetText + '#REF!' : refText(t)) : t.text; }).join(''); }

  /* 복사·붙여넣기 시 상대 참조 이동 */
  function shift(formula, dr, dc) {
    const toks = tokenize(formula);
    toks.forEach(function (t) {
      if (t.t !== 'ref') return;
      [t.a, t.b].forEach(function (p) {
        if (!p) return;
        if (t.kind !== 'col' && !p.rAbs) p.r += dr;
        if (t.kind !== 'row' && !p.cAbs) p.c += dc;
        if (p.r < 0 || p.c < 0) t.bad = true;
      });
    });
    return rebuild(toks);
  }

  /* 행/열 삽입·삭제 시 참조 조정. axis: 'row'|'col', count > 0 삽입, < 0 삭제 */
  function adjust(formula, ownSheet, targetSheet, axis, at, count) {
    const toks = tokenize(formula);
    let changed = false;
    const key = axis === 'row' ? 'r' : 'c';
    toks.forEach(function (t) {
      if (t.t !== 'ref') return;
      const sheet = t.sheet != null ? t.sheet : ownSheet;
      if (sheet.toLowerCase() !== targetSheet.toLowerCase()) return;
      if ((axis === 'row' && t.kind === 'col') || (axis === 'col' && t.kind === 'row')) return;
      if (count > 0) {
        [t.a, t.b].forEach(function (p) { if (p && p[key] >= at) { p[key] += count; changed = true; } });
      } else {
        const del = -count, end = at + del - 1;
        if (!t.b) {
          if (t.a[key] >= at && t.a[key] <= end) { t.bad = true; changed = true; }
          else if (t.a[key] > end) { t.a[key] -= del; changed = true; }
        } else {
          let lo = Math.min(t.a[key], t.b[key]), hi = Math.max(t.a[key], t.b[key]);
          if (lo >= at && hi <= end) { t.bad = true; changed = true; return; }
          const nlo = lo > end ? lo - del : lo >= at ? at : lo;
          const nhi = hi > end ? hi - del : hi >= at ? at - 1 : hi;
          if (nlo !== lo || nhi !== hi) changed = true;
          if (t.a[key] <= t.b[key]) { t.a[key] = nlo; t.b[key] = nhi; } else { t.b[key] = nlo; t.a[key] = nhi; }
        }
      }
    });
    return changed ? rebuild(toks) : formula;
  }

  /* 시트 이름 변경 시 참조 갱신 */
  function renameSheet(formula, oldName, newName) {
    const toks = tokenize(formula);
    let changed = false;
    toks.forEach(function (t) {
      if (t.t === 'ref' && t.sheet != null && t.sheet.toLowerCase() === oldName.toLowerCase()) {
        t.sheetText = quoteSheet(newName) + '!';
        changed = true;
      }
    });
    return changed ? rebuild(toks) : formula;
  }
  function quoteSheet(name) { return /^[A-Za-z_\u00A0-\uFFFF][\w.\u00A0-\uFFFF]*$/.test(name) && !/^[A-Za-z]{1,3}\d+$/.test(name) ? name : "'" + name.replace(/'/g, "''") + "'"; }

  DV.formula = {
    Err: Err, E: E, isErr: isErr, Range: Range,
    tokenize: tokenize, compile: compile, evaluate: evaluate,
    colName: colName, colIndex: colIndex, addr: addr, parseAddr: parseAddr,
    shift: shift, adjust: adjust, renameSheet: renameSheet, quoteSheet: quoteSheet,
    toStr: toStr, fmtNum: fmtNum, numFromStr: numFromStr, dateSerial: dateSerial,
    functions: Object.keys(FN).sort()
  };
})(window.DV);
