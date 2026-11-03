/* Launchpad Grid - the formula engine.
   Reads formulas like =SUM(A1:A5)*2 into a tree and works them out.
   No eval: everything is parsed by hand, so a cell can only ever do maths. */
(function (root) {
  function Err(code) { this.err = code; }
  Err.prototype.toString = function () { return this.err; };
  var E = {
    div: new Err("#DIV/0!"), val: new Err("#VALUE!"), ref: new Err("#REF!"), name: new Err("#NAME?"),
    na: new Err("#N/A"), num: new Err("#NUM!"), circ: new Err("#CIRC!")
  };
  function isErr(v) { return v instanceof Err; }

  /* ---------- cell names ---------- */
  function colName(c) { var s = ""; c++; while (c > 0) { var m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - 1) / 26); } return s; }
  function colNum(s) { var n = 0; for (var i = 0; i < s.length; i++) n = n * 26 + (s.charCodeAt(i) - 64); return n - 1; }
  function parseRef(s) {
    var m = /^(\$?)([A-Z]{1,3})(\$?)(\d+)$/.exec(String(s).toUpperCase()); if (!m) return null;
    return { c: colNum(m[2]), r: +m[4] - 1, ac: !!m[1], ar: !!m[3] };
  }
  function refName(r, c) { return colName(c) + (r + 1); }

  /* ---------- reading a formula into tokens ---------- */
  function tokenize(src) {
    var t = [], i = 0, s = src;
    while (i < s.length) {
      var ch = s[i];
      if (/\s/.test(ch)) { i++; continue; }
      if (ch === '"') {
        var j = i + 1, str = "";
        for (;;) {
          if (j >= s.length) throw E.val;
          if (s[j] === '"') { if (s[j + 1] === '"') { str += '"'; j += 2; continue; } break; }
          str += s[j++];
        }
        t.push({ t: "str", v: str }); i = j + 1; continue;
      }
      var m = /^(\d+\.?\d*|\.\d+)(e[+-]?\d+)?/i.exec(s.slice(i));
      if (m) { t.push({ t: "num", v: parseFloat(m[0]) }); i += m[0].length; continue; }
      m = /^\$?[A-Za-z]{1,3}\$?\d+(?![A-Za-z0-9_(])/.exec(s.slice(i));
      if (m) { t.push({ t: "ref", v: m[0].toUpperCase() }); i += m[0].length; continue; }
      m = /^[A-Za-z_][A-Za-z0-9_.]*/.exec(s.slice(i));
      if (m) {
        var w = m[0].toUpperCase(); i += m[0].length;
        if (w === "TRUE" || w === "FALSE") t.push({ t: "bool", v: w === "TRUE" });
        else t.push({ t: "name", v: w });
        continue;
      }
      var two = s.substr(i, 2);
      if (two === "<=" || two === ">=" || two === "<>") { t.push({ t: "op", v: two }); i += 2; continue; }
      if ("+-*/^&=<>%".indexOf(ch) >= 0) { t.push({ t: "op", v: ch }); i++; continue; }
      if (ch === "(" || ch === ")" || ch === "," || ch === ";" || ch === ":") { t.push({ t: ch === ";" ? "," : ch }); i++; continue; }
      throw E.val;
    }
    return t;
  }

  /* ---------- tokens into a tree ---------- */
  function parse(src) {
    var toks = tokenize(src), p = 0;
    function peek() { return toks[p]; }
    function isOp(v) { var k = toks[p]; return k && k.t === "op" && k.v === v; }
    function next() { return toks[p++]; }
    function expect(t) { var k = next(); if (!k || k.t !== t) throw E.val; return k; }
    function cmp() {
      var a = cat();
      while (peek() && peek().t === "op" && /^(=|<>|<|>|<=|>=)$/.test(peek().v)) { var o = next().v; a = { k: "bin", o: o, a: a, b: cat() }; }
      return a;
    }
    function cat() { var a = add(); while (isOp("&")) { next(); a = { k: "bin", o: "&", a: a, b: add() }; } return a; }
    function add() { var a = mul(); while (isOp("+") || isOp("-")) { var o = next().v; a = { k: "bin", o: o, a: a, b: mul() }; } return a; }
    function mul() { var a = pow(); while (isOp("*") || isOp("/")) { var o = next().v; a = { k: "bin", o: o, a: a, b: pow() }; } return a; }
    function pow() { var a = unary(); while (isOp("^")) { next(); a = { k: "bin", o: "^", a: a, b: unary() }; } return a; }
    function unary() {
      if (isOp("-")) { next(); return { k: "neg", a: unary() }; }
      if (isOp("+")) { next(); return unary(); }
      var a = prim();
      while (isOp("%")) { next(); a = { k: "pct", a: a }; }
      return a;
    }
    function prim() {
      var k = next(); if (!k) throw E.val;
      if (k.t === "num") return { k: "num", v: k.v };
      if (k.t === "str") return { k: "str", v: k.v };
      if (k.t === "bool") return { k: "bool", v: k.v };
      if (k.t === "ref") {
        var a = parseRef(k.v); if (!a) throw E.ref;
        if (peek() && peek().t === ":") { next(); var b = parseRef(expect("ref").v); if (!b) throw E.ref; return { k: "range", a: a, b: b }; }
        return { k: "ref", a: a };
      }
      if (k.t === "name") {
        if (!peek() || peek().t !== "(") { if (k.v === "PI") return { k: "fn", n: "PI", args: [] }; throw E.name; }
        next(); var args = [];
        if (peek() && peek().t === ")") { next(); return { k: "fn", n: k.v, args: args }; }
        for (;;) {
          if (peek() && (peek().t === "," || peek().t === ")")) args.push({ k: "blank" }); else args.push(cmp());
          var d = next(); if (!d) throw E.val;
          if (d.t === ")") break; if (d.t !== ",") throw E.val;
        }
        return { k: "fn", n: k.v, args: args };
      }
      if (k.t === "(") { var e = cmp(); expect(")"); return e; }
      throw E.val;
    }
    var tree = cmp(); if (p < toks.length) throw E.val;
    return tree;
  }

  /* ---------- working it out ---------- */
  function num(v) {
    if (isErr(v)) throw v;
    if (v === null || v === undefined || v === "") return 0;
    if (typeof v === "number") return v;
    if (typeof v === "boolean") return v ? 1 : 0;
    var s = String(v).trim().replace(/^£/, "").replace(/,/g, "");
    if (/%$/.test(s) && isFinite(parseFloat(s))) return parseFloat(s) / 100;
    if (s !== "" && isFinite(+s)) return +s;
    throw E.val;
  }
  function str(v) {
    if (isErr(v)) throw v;
    if (v === null || v === undefined) return "";
    if (typeof v === "boolean") return v ? "TRUE" : "FALSE";
    if (typeof v === "number") return fmtNum(v);
    return String(v);
  }
  function bool(v) {
    if (isErr(v)) throw v;
    if (typeof v === "boolean") return v;
    if (v === null || v === "") return false;
    if (typeof v === "number") return v !== 0;
    var s = String(v).toUpperCase(); if (s === "TRUE") return true; if (s === "FALSE") return false;
    throw E.val;
  }
  function fmtNum(n) {
    if (!isFinite(n)) return "#NUM!";
    if (Number.isInteger(n)) return String(n);
    var s = String(+n.toPrecision(12)); return s;
  }
  function chk(n) { if (!isFinite(n)) throw E.num; return n; }

  function evaluate(node, ctx) {
    switch (node.k) {
      case "num": case "str": case "bool": return node.v;
      case "blank": return null;
      case "ref": return ctx.cell(node.a.r, node.a.c);
      case "range": return { range: true, a: node.a, b: node.b, ctx: ctx };
      case "neg": return -num(scalar(evaluate(node.a, ctx)));
      case "pct": return num(scalar(evaluate(node.a, ctx))) / 100;
      case "bin": {
        var a = scalar(evaluate(node.a, ctx)), b = scalar(evaluate(node.b, ctx));
        if (isErr(a)) throw a; if (isErr(b)) throw b;
        switch (node.o) {
          case "+": return chk(num(a) + num(b));
          case "-": return chk(num(a) - num(b));
          case "*": return chk(num(a) * num(b));
          case "/": { var d = num(b); if (d === 0) throw E.div; return num(a) / d; }
          case "^": return chk(Math.pow(num(a), num(b)));
          case "&": return str(a) + str(b);
          default: return compare(a, b, node.o);
        }
      }
      case "fn": {
        var f = FN[node.n]; if (!f) throw E.name;
        if (f.lazy) return f(node.args, ctx);
        return f(node.args.map(function (x) { return evaluate(x, ctx); }), ctx);
      }
    }
    throw E.val;
  }
  function compare(a, b, o) {
    var x = a, y = b;
    if (x === null) x = typeof y === "string" ? "" : 0;
    if (y === null) y = typeof x === "string" ? "" : 0;
    var c;
    if (typeof x === "number" && typeof y === "number") c = x - y;
    else if (typeof x === "string" && typeof y === "string") c = x.toLowerCase() < y.toLowerCase() ? -1 : x.toLowerCase() > y.toLowerCase() ? 1 : 0;
    else if (typeof x === "boolean" && typeof y === "boolean") c = (x ? 1 : 0) - (y ? 1 : 0);
    else { var rank = function (v) { return typeof v === "number" ? 0 : typeof v === "string" ? 1 : 2; }; c = rank(x) - rank(y); }
    switch (o) { case "=": return c === 0; case "<>": return c !== 0; case "<": return c < 0; case ">": return c > 0; case "<=": return c <= 0; case ">=": return c >= 0; }
  }
  function scalar(v) {
    if (v && v.range) { // a range used as one value: its top-left cell
      return v.ctx.cell(Math.min(v.a.r, v.b.r), Math.min(v.a.c, v.b.c));
    }
    return v;
  }
  function grid(v) { // a range as rows of values
    if (!(v && v.range)) return [[v]];
    var r1 = Math.min(v.a.r, v.b.r), r2 = Math.max(v.a.r, v.b.r), c1 = Math.min(v.a.c, v.b.c), c2 = Math.max(v.a.c, v.b.c);
    if ((r2 - r1 + 1) * (c2 - c1 + 1) > 200000) throw E.num;
    var out = [];
    for (var r = r1; r <= r2; r++) { var row = []; for (var c = c1; c <= c2; c++) row.push(v.ctx.cell(r, c)); out.push(row); }
    return out;
  }
  /* every value in the arguments; numbers only from ranges (like Excel) */
  function nums(args, strict) {
    var out = [];
    args.forEach(function (a) {
      if (a && a.range) grid(a).forEach(function (row) { row.forEach(function (v) { if (isErr(v)) throw v; if (typeof v === "number") out.push(v); }); });
      else if (a !== null) { out.push(num(a)); }
    });
    return out;
  }
  function flat(args) {
    var out = [];
    args.forEach(function (a) { if (a && a.range) grid(a).forEach(function (row) { out.push.apply(out, row); }); else out.push(a); });
    return out;
  }
  function criteria(c) {
    if (typeof c === "number" || typeof c === "boolean") return function (v) { return v === c || (typeof v === "string" && v !== "" && +v === c); };
    var s = str(c), m = /^(<=|>=|<>|<|>|=)?(.*)$/.exec(s), op = m[1] || "=", rhs = m[2];
    var n = rhs !== "" && isFinite(+rhs) ? +rhs : null;
    if (n !== null) return function (v) { if (typeof v !== "number") { if (op === "<>") return true; return false; } return compare(v, n, op); };
    var re = new RegExp("^" + rhs.replace(/[.+^${}()|[\]\\]/g, "\\$&").replace(/\*/g, ".*").replace(/\?/g, ".") + "$", "i");
    return function (v) { var t = v === null ? "" : str(v); var hit = re.test(t); return op === "<>" ? !hit : op === "=" ? hit : compare(t, rhs, op); };
  }
  function round(n, d, mode) {
    var k = Math.pow(10, d), x = n * k;
    x = mode === "up" ? (x < 0 ? -Math.ceil(-x - 1e-9) : Math.ceil(x - 1e-9)) : mode === "down" ? (x < 0 ? -Math.floor(-x + 1e-9) : Math.floor(x + 1e-9)) : Math.sign(x) * Math.round(Math.abs(x) + 1e-9);
    return x / k;
  }

  var FN = {
    SUM: function (a) { return nums(a).reduce(function (s, x) { return s + x; }, 0); },
    PRODUCT: function (a) { return nums(a).reduce(function (s, x) { return s * x; }, 1); },
    AVERAGE: function (a) { var n = nums(a); if (!n.length) throw E.div; return n.reduce(function (s, x) { return s + x; }, 0) / n.length; },
    MIN: function (a) { var n = nums(a); return n.length ? Math.min.apply(null, n) : 0; },
    MAX: function (a) { var n = nums(a); return n.length ? Math.max.apply(null, n) : 0; },
    MEDIAN: function (a) { var n = nums(a).sort(function (x, y) { return x - y; }); if (!n.length) throw E.num; var h = n.length >> 1; return n.length % 2 ? n[h] : (n[h - 1] + n[h]) / 2; },
    MODE: function (a) { var n = nums(a), best = null, bc = 1, cnt = {}; n.forEach(function (x) { cnt[x] = (cnt[x] || 0) + 1; if (cnt[x] > bc) { bc = cnt[x]; best = x; } }); if (best === null) throw E.na; return best; },
    STDEV: function (a) { var n = nums(a); if (n.length < 2) throw E.div; var m = n.reduce(function (s, x) { return s + x; }, 0) / n.length; return Math.sqrt(n.reduce(function (s, x) { return s + (x - m) * (x - m); }, 0) / (n.length - 1)); },
    COUNT: function (a) { var c = 0; flat(a).forEach(function (v) { if (typeof v === "number") c++; }); return c; },
    COUNTA: function (a) { var c = 0; flat(a).forEach(function (v) { if (v !== null && v !== "") c++; }); return c; },
    COUNTBLANK: function (a) { var c = 0; flat(a).forEach(function (v) { if (v === null || v === "") c++; }); return c; },
    ROUND: function (a) { return round(num(scalar(a[0])), num(scalar(a[1]))); },
    ROUNDUP: function (a) { return round(num(scalar(a[0])), num(scalar(a[1])), "up"); },
    ROUNDDOWN: function (a) { return round(num(scalar(a[0])), num(scalar(a[1])), "down"); },
    INT: function (a) { return Math.floor(num(scalar(a[0]))); },
    ABS: function (a) { return Math.abs(num(scalar(a[0]))); },
    SQRT: function (a) { var n = num(scalar(a[0])); if (n < 0) throw E.num; return Math.sqrt(n); },
    POWER: function (a) { return chk(Math.pow(num(scalar(a[0])), num(scalar(a[1])))); },
    MOD: function (a) { var d = num(scalar(a[1])); if (d === 0) throw E.div; var n = num(scalar(a[0])); return n - d * Math.floor(n / d); },
    PI: function () { return Math.PI; },
    RAND: function () { return Math.random(); },
    RANDBETWEEN: function (a) { var lo = Math.ceil(num(scalar(a[0]))), hi = Math.floor(num(scalar(a[1]))); if (hi < lo) throw E.num; return lo + Math.floor(Math.random() * (hi - lo + 1)); },
    IF: Object.assign(function (args, ctx) {
      var c = bool(scalar(evaluate(args[0], ctx)));
      var pick = c ? args[1] : args[2];
      if (!pick) return c;
      return scalar(evaluate(pick, ctx));
    }, { lazy: true }),
    IFERROR: Object.assign(function (args, ctx) {
      try { var v = scalar(evaluate(args[0], ctx)); if (isErr(v)) throw v; return v; }
      catch (e) { if (isErr(e)) return args[1] ? scalar(evaluate(args[1], ctx)) : ""; throw e; }
    }, { lazy: true }),
    AND: function (a) { return flat(a).every(function (v) { return v === null || bool(v); }); },
    OR: function (a) { return flat(a).some(function (v) { return v !== null && bool(v); }); },
    NOT: function (a) { return !bool(scalar(a[0])); },
    CONCAT: function (a) { return flat(a).map(str).join(""); },
    CONCATENATE: function (a) { return flat(a).map(str).join(""); },
    LEN: function (a) { return str(scalar(a[0])).length; },
    UPPER: function (a) { return str(scalar(a[0])).toUpperCase(); },
    LOWER: function (a) { return str(scalar(a[0])).toLowerCase(); },
    PROPER: function (a) { return str(scalar(a[0])).toLowerCase().replace(/(^|[^a-z])([a-z])/g, function (m, p, c) { return p + c.toUpperCase(); }); },
    TRIM: function (a) { return str(scalar(a[0])).trim().replace(/\s+/g, " "); },
    LEFT: function (a) { return str(scalar(a[0])).slice(0, a.length > 1 ? num(scalar(a[1])) : 1); },
    RIGHT: function (a) { var s = str(scalar(a[0])), n = a.length > 1 ? num(scalar(a[1])) : 1; return n ? s.slice(-n) : ""; },
    MID: function (a) { var s = str(scalar(a[0])), st = num(scalar(a[1])), n = num(scalar(a[2])); if (st < 1) throw E.val; return s.substr(st - 1, n); },
    TODAY: function () { var d = new Date(); return ("0" + d.getDate()).slice(-2) + "/" + ("0" + (d.getMonth() + 1)).slice(-2) + "/" + d.getFullYear(); },
    NOW: function () { var d = new Date(); return FN.TODAY() + " " + d.toTimeString().slice(0, 5); },
    SUMIF: function (a) {
      var test = criteria(scalar(a[1])), src = grid(a[0]), sum = a[2] ? grid(a[2]) : src, t = 0;
      src.forEach(function (row, i) { row.forEach(function (v, j) { if (test(v)) { var x = sum[i] && sum[i][j]; if (typeof x === "number") t += x; } }); });
      return t;
    },
    COUNTIF: function (a) { var test = criteria(scalar(a[1])), c = 0; grid(a[0]).forEach(function (row) { row.forEach(function (v) { if (test(v)) c++; }); }); return c; },
    AVERAGEIF: function (a) {
      var test = criteria(scalar(a[1])), src = grid(a[0]), av = a[2] ? grid(a[2]) : src, t = 0, n = 0;
      src.forEach(function (row, i) { row.forEach(function (v, j) { if (test(v)) { var x = av[i] && av[i][j]; if (typeof x === "number") { t += x; n++; } } }); });
      if (!n) throw E.div; return t / n;
    },
    VLOOKUP: function (a) {
      var want = scalar(a[0]), tbl = grid(a[1]), col = num(scalar(a[2])), approx = a.length > 3 && a[3] !== null ? bool(scalar(a[3])) : true;
      if (col < 1 || col > (tbl[0] || []).length) throw E.ref;
      var hit = -1;
      for (var i = 0; i < tbl.length; i++) {
        var v = tbl[i][0];
        if (approx) { if (v !== null && compare(v, want, "<=")) hit = i; else if (v !== null) break; }
        else if (compare(v, want, "=") && v !== null) { hit = i; break; }
      }
      if (hit < 0) throw E.na; return tbl[hit][col - 1];
    },
    HLOOKUP: function (a) {
      var want = scalar(a[0]), tbl = grid(a[1]), row = num(scalar(a[2])), approx = a.length > 3 && a[3] !== null ? bool(scalar(a[3])) : true;
      if (row < 1 || row > tbl.length) throw E.ref;
      var hit = -1, top = tbl[0];
      for (var i = 0; i < top.length; i++) {
        if (approx) { if (top[i] !== null && compare(top[i], want, "<=")) hit = i; else if (top[i] !== null) break; }
        else if (top[i] !== null && compare(top[i], want, "=")) { hit = i; break; }
      }
      if (hit < 0) throw E.na; return tbl[row - 1][hit];
    },
    INDEX: function (a) {
      var t = grid(a[0]), r = num(scalar(a[1])), c = a.length > 2 ? num(scalar(a[2])) : 1;
      if (t.length === 1 && a.length === 2) { c = r; r = 1; }
      if (r < 1 || c < 1 || !t[r - 1] || t[r - 1][c - 1] === undefined) throw E.ref; return t[r - 1][c - 1];
    },
    MATCH: function (a) {
      var want = scalar(a[0]), list = flat([a[1]]), type = a.length > 2 ? num(scalar(a[2])) : 1;
      for (var i = 0, hit = -1; i < list.length; i++) {
        if (type === 0) { if (list[i] !== null && compare(list[i], want, "=")) return i + 1; }
        else if (type > 0) { if (list[i] !== null && compare(list[i], want, "<=")) hit = i; else break; }
        else { if (list[i] !== null && compare(list[i], want, ">=")) hit = i; else break; }
      }
      if (type === 0 || hit < 0) throw E.na; return hit + 1;
    },
    RANK: function (a) {
      var n = num(scalar(a[0])), list = nums([a[1]]), desc = !(a.length > 2 && num(scalar(a[2])));
      if (list.indexOf(n) < 0) throw E.na;
      return 1 + list.filter(function (x) { return desc ? x > n : x < n; }).length;
    }
  };

  /* Work out a formula (text after "="). Returns a value or an Err. */
  var cache = {};
  function run(text, ctx) {
    try {
      var tree = cache[text];
      if (!tree) { tree = parse(text); if (Object.keys(cache).length > 5000) cache = {}; cache[text] = tree; }
      var v = scalar(evaluate(tree, ctx));
      if (v === null) return 0;
      if (typeof v === "number" && !isFinite(v)) return E.num;
      return v;
    } catch (e) { if (isErr(e)) return e; return E.val; }
  }

  /* Change the cell names in a formula, leaving strings alone.
     fn(ref) gets {r,c,ac,ar} and returns a new one, or null for #REF!. */
  function mapRefs(text, fn) {
    var out = "", i = 0, s = text;
    while (i < s.length) {
      if (s[i] === '"') { var j = i + 1; while (j < s.length && !(s[j] === '"' && s[j + 1] !== '"')) j += s[j] === '"' ? 2 : 1; out += s.slice(i, j + 1); i = j + 1; continue; }
      var m = /^\$?[A-Za-z]{1,3}\$?\d+(?![A-Za-z0-9_(])/.exec(s.slice(i));
      var prev = i ? s[i - 1] : "";
      if (m && !/[A-Za-z0-9_.]/.test(prev)) {
        var r = parseRef(m[0]), n = r ? fn(r) : null;
        out += n ? (n.ac ? "$" : "") + colName(n.c) + (n.ar ? "$" : "") + (n.r + 1) : "#REF!";
        i += m[0].length; continue;
      }
      var w = /^[A-Za-z_][A-Za-z0-9_.]*/.exec(s.slice(i));
      if (w) { out += w[0]; i += w[0].length; continue; }
      out += s[i++];
    }
    return out;
  }
  function shift(text, dr, dc) {
    return mapRefs(text, function (r) {
      var n = { r: r.ar ? r.r : r.r + dr, c: r.ac ? r.c : r.c + dc, ar: r.ar, ac: r.ac };
      return n.r < 0 || n.c < 0 ? null : n;
    });
  }

  var api = { run: run, parse: parse, shift: shift, mapRefs: mapRefs, colName: colName, colNum: colNum, parseRef: parseRef, refName: refName,
    isErr: isErr, Err: Err, fmtNum: fmtNum, names: Object.keys(FN) };
  if (typeof module !== "undefined") module.exports = api; else root.Formula = api;
})(this);
