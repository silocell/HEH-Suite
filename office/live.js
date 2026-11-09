/* =====================================================================
   Launchpad Suite - working together live.

   The owner shares a file with people by their email ("Can edit" or
   "Can view"). Everyone who has it open joins one live room, and every
   change goes to the others straight away.

   How the changes fit together: a file is broken into small parts - each
   paragraph of a document, each cell of a sheet, each box on a slide,
   each shape on a board. Every part carries a stamp (a counter that only
   goes up, plus who changed it). When two people change different parts,
   both changes stay. When they change the same part at once, the later
   stamp wins for that part only. Lists (paragraphs, slides, shapes) keep
   anything either person added.

   Needs office.js first. An app turns it on by giving its autosaver a
   third argument:
     Office.autosaver(statusEl, getFile, {
       apply(body, title)   put a file that changed elsewhere on screen
       local: ["zoom"]      body keys that are only for this screen
       surface: element     where typing happens (blocked for viewers)
       where()              a CSS selector for what you're working on
       html: true           split body.html (and any "html") into paragraphs
     })
   and calling saver.opened() whenever it has loaded a file.
   ===================================================================== */
(function () {
  if (!window.Office) return;
  var O = window.Office, SEP = "\u0001";
  var me = Math.random().toString(36).slice(2, 10) + Date.now().toString(36).slice(-4);
  var COLS = ["#e11d48", "#2563eb", "#16a34a", "#9333ea", "#ea580c", "#0891b2", "#ca8a04", "#db2777"];
  function colFor(s) { var h = 0; for (var i = 0; i < s.length; i++) h = (h * 31 + s.charCodeAt(i)) >>> 0; return COLS[h % COLS.length]; }
  var esc = O.esc;

  /* ---------- remember when each opened file was last saved ---------- */
  var stamps = {};
  var get0 = O.get;
  O.get = function (id) {
    return get0(id).then(function (f) { if (f) stamps[f.id] = f.updated_at || ""; return f; });
  };

  /* ---------- paragraphs get a short id that stays with them ---------- */
  function bid() { return "b" + Math.random().toString(36).slice(2, 9); }
  function blockIds(el) {
    if (!el) return;
    var seen = {}, i = 0, legacy = !el.querySelector(":scope > [data-b]");
    for (var n = el.firstElementChild; n; n = n.nextElementSibling, i++) {
      var id = n.getAttribute("data-b");
      // a file from before: number its paragraphs the same way on every screen
      if (!id && legacy) id = "p" + i;
      if (!id || !/^[a-z0-9]{1,16}$/.test(id) || seen[id]) { id = bid(); }
      if (n.getAttribute("data-b") !== id) n.setAttribute("data-b", id);
      seen[id] = 1;
    }
  }
  /* Put new HTML into an element, keeping the paragraphs that didn't
     change exactly as they are (so the cursor stays where it was). */
  function patchBlocks(el, html) {
    var t = document.createElement("template"); t.innerHTML = html;
    var want = Array.prototype.slice.call(t.content.childNodes);
    var have = {}; Array.prototype.forEach.call(el.children, function (c) { var k = c.getAttribute("data-b"); if (k) have[k] = c; });
    var sel = getSelection(), caret = null;
    if (sel.rangeCount && el.contains(sel.anchorNode)) {
      var blk = sel.anchorNode.nodeType === 1 ? sel.anchorNode : sel.anchorNode.parentNode;
      while (blk && blk.parentNode !== el) blk = blk.parentNode;
      if (blk && blk.getAttribute) {
        var r = document.createRange(); r.setStart(blk, 0); r.setEnd(sel.anchorNode, sel.anchorOffset);
        caret = { id: blk.getAttribute("data-b"), off: r.toString().length, html: blk.outerHTML, text: blk.textContent };
      }
    }
    var keep = [], moved = false;
    want.forEach(function (n) {
      var k = n.nodeType === 1 && n.getAttribute("data-b"), old = k && have[k];
      if (old && old.outerHTML === n.outerHTML) { keep.push(old); delete have[k]; }
      else keep.push(n);
    });
    // swap in the new list, touching only what changed
    var cur = el.firstChild;
    keep.forEach(function (n) {
      if (n === cur) { cur = cur.nextSibling; return; }
      el.insertBefore(n, cur); moved = true;
    });
    while (cur) { var nx = cur.nextSibling; el.removeChild(cur); cur = nx; }
    if (caret) {
      var b = el.querySelector('[data-b="' + caret.id + '"]');
      if (b && b.outerHTML !== caret.html) {
        // the paragraph you were in changed: put the cursor back at the same spot
        // move it along with any words added or taken away before it
        var o = caret.text, nt = b.textContent, p = 0, q = 0, mn = Math.min(o.length, nt.length);
        while (p < mn && o.charCodeAt(p) === nt.charCodeAt(p)) p++;
        while (q < mn - p && o.charCodeAt(o.length - 1 - q) === nt.charCodeAt(nt.length - 1 - q)) q++;
        var off = caret.off <= p ? caret.off : caret.off >= o.length - q ? caret.off + nt.length - o.length : nt.length - q;
        var w = document.createTreeWalker(b, NodeFilter.SHOW_TEXT), left = off, tn, done = false;
        while ((tn = w.nextNode())) {
          if (left <= tn.nodeValue.length) { var rr = document.createRange(); rr.setStart(tn, left); rr.collapse(true); sel.removeAllRanges(); sel.addRange(rr); done = true; break; }
          left -= tn.nodeValue.length;
        }
        if (!done) { var r3 = document.createRange(); r3.selectNodeContents(b); r3.collapse(false); sel.removeAllRanges(); sel.addRange(r3); }
      }
    }
    return moved;
  }

  /* ---------- a file as a list of small parts ---------- */
  function keyedIds(a) {
    if (!a.length) return null;
    var ids = [], seen = {};
    for (var i = 0; i < a.length; i++) {
      var x = a[i];
      if (!x || typeof x !== "object" || Array.isArray(x) || typeof x.id !== "string" || !x.id || seen[x.id]) return null;
      seen[x.id] = 1; ids.push(x.id);
    }
    return ids;
  }
  function allObjects(a) { return a.length && a.every(function (x) { return x && typeof x === "object" && !Array.isArray(x); }); }
  function splitHTML(s) {
    if (s.indexOf("data-b=") < 0) return null;
    var t = document.createElement("template"); t.innerHTML = s;
    var out = [], ids = [];
    for (var n = t.content.firstChild; n; n = n.nextSibling) {
      if (n.nodeType === 3 && !n.nodeValue.trim()) continue;
      var k = n.nodeType === 1 && n.getAttribute("data-b");
      if (!k || ids.indexOf(k) >= 0) return null;
      ids.push(k); out.push([k, n.outerHTML]);
    }
    return { ids: ids, parts: out };
  }
  function flat(v, path, out, html) {
    if (Array.isArray(v)) {
      var ids = keyedIds(v);
      if (ids) { out[path + SEP + "#"] = JSON.stringify(ids); v.forEach(function (x) { flat(x, path + SEP + "@" + x.id, out, html); }); return; }
      if (allObjects(v)) { out[path + SEP + "#n"] = String(v.length); v.forEach(function (x, i) { flat(x, path + SEP + "$" + i, out, html); }); return; }
      out[path] = JSON.stringify(v); return;
    }
    if (v && typeof v === "object") {
      var ks = Object.keys(v);
      if (!ks.length) { out[path] = "{}"; return; }
      ks.forEach(function (k) { if (v[k] !== undefined) flat(v[k], path + SEP + "." + k, out, html); });
      return;
    }
    if (html && typeof v === "string" && /\.html$/.test(path)) {
      var sp = splitHTML(v);
      if (sp) { out[path + SEP + "%"] = JSON.stringify(sp.ids); sp.parts.forEach(function (p) { out[path + SEP + "~" + p[0]] = JSON.stringify(p[1]); }); return; }
    }
    out[path] = JSON.stringify(v === undefined ? null : v);
  }
  function unflat(map) {
    var root = { k: {} };
    Object.keys(map).forEach(function (key) {
      var segs = key.split(SEP), n = root;
      for (var i = 1; i < segs.length; i++) n = n.k[segs[i]] || (n.k[segs[i]] = { k: {} });
      n.leaf = map[key];
    });
    return build(root);
  }
  function build(n) {
    var ks = Object.keys(n.k);
    if (!ks.length) { if (n.leaf === undefined) return undefined; try { return JSON.parse(n.leaf); } catch (e) { return null; } }
    var c0 = ks.filter(function (k) { return k !== "#" && k !== "#n" && k !== "%"; })[0] || ks[0];
    var t = c0.charAt(0);
    if (t === "@" || c0 === "#") {
      var order = []; try { order = JSON.parse(n.k["#"] && n.k["#"].leaf || "[]"); } catch (e) {}
      var outA = [], used = {};
      order.forEach(function (id) { var kid = n.k["@" + id]; if (kid && !used[id]) { used[id] = 1; var v = build(kid); if (v !== undefined) outA.push(v); } });
      ks.forEach(function (k) { if (k.charAt(0) === "@" && !used[k.slice(1)]) { var v = build(n.k[k]); if (v !== undefined) outA.push(v); } });
      return outA;
    }
    if (t === "$" || c0 === "#n") {
      var len = parseInt(n.k["#n"] && n.k["#n"].leaf, 10); if (!(len >= 0)) len = 0;
      ks.forEach(function (k) { if (k.charAt(0) === "$" && n.k["#n"] === undefined) len = Math.max(len, +k.slice(1) + 1); });
      var outB = [];
      for (var i = 0; i < len; i++) { var kid = n.k["$" + i]; var v2 = kid ? build(kid) : undefined; outB.push(v2 === undefined ? {} : v2); }
      return outB;
    }
    if (t === "~" || c0 === "%") {
      var ord = []; try { ord = JSON.parse(n.k["%"] && n.k["%"].leaf || "[]"); } catch (e) {}
      var parts = [], u = {};
      ord.forEach(function (id) { var kid = n.k["~" + id]; if (kid && !u[id]) { u[id] = 1; parts.push(build(kid)); } });
      ks.forEach(function (k) { if (k.charAt(0) === "~" && !u[k.slice(1)]) parts.push(build(n.k[k])); });
      return parts.join("");
    }
    var o = {};
    ks.forEach(function (k) { if (k.charAt(0) === ".") { var v = build(n.k[k]); if (v !== undefined) o[k.slice(1)] = v; } });
    return o;
  }
  // the order lists, and the key of an item in them
  function orderKind(key) { var e = key.slice(key.lastIndexOf(SEP) + 1); return e === "#" ? "@" : e === "%" ? "~" : null; }

  /* ---------- typing in the same paragraph as someone else ----------
     A change to a piece of text is also sent as "what was there, around
     where, and what it became". If the other screen's copy has moved on
     too, the change is fitted in by finding those words, so both people's
     typing stays (as in Google Docs). If it can't be placed, the later
     change wins for that piece. */
  function hash(s) { var h = 2166136261; for (var i = 0; i < s.length; i++) { h ^= s.charCodeAt(i); h = Math.imul(h, 16777619); } return (h >>> 0).toString(36); }
  function mkPatch(a, b) {
    if (typeof a !== "string" || typeof b !== "string" || a.length < 3) return null;
    var p = 0, n = Math.min(a.length, b.length);
    while (p < n && a.charCodeAt(p) === b.charCodeAt(p)) p++;
    var q = 0;
    while (q < n - p && a.charCodeAt(a.length - 1 - q) === b.charCodeAt(b.length - 1 - q)) q++;
    var oldMid = a.slice(p, a.length - q), newMid = b.slice(p, b.length - q);
    if (oldMid.length + newMid.length > 4000) return null;
    return [hash(a), a.slice(Math.max(0, p - 24), p), oldMid, newMid, a.slice(a.length - q, a.length - q + 24), p];
  }
  function fit(v, pt) {
    if (!Array.isArray(pt) || pt.length < 6) return null;
    var before = String(pt[1]), oldMid = String(pt[2]), newMid = String(pt[3]), after = String(pt[4]), at = +pt[5] || 0;
    function near(needle, off) {
      // the match closest to where it was
      var best = -1, i = v.indexOf(needle);
      while (i >= 0) { if (best < 0 || Math.abs(i + off - at) < Math.abs(best + off - at)) best = i; i = v.indexOf(needle, i + 1); }
      return best;
    }
    // the words around it, then fewer of them, then just one side
    var tries = [];
    [24, 12, 6, 3].forEach(function (L) {
      var b0 = before.slice(-L), a0 = after.slice(0, L);
      tries.push([b0 + oldMid + a0, b0.length]);
    });
    [12, 6].forEach(function (L) { var b0 = before.slice(-L), a0 = after.slice(0, L); tries.push([oldMid + a0, 0]); tries.push([b0 + oldMid, b0.length]); });
    for (var t = 0; t < tries.length; t++) {
      if (tries[t][0].length < 4) continue;
      var i = near(tries[t][0], tries[t][1]);
      if (i >= 0) { var s0 = i + tries[t][1]; var out = v.slice(0, s0) + newMid + v.slice(s0 + oldMid.length); try { JSON.parse(out); return out; } catch (e) { return null; } }
    }
    return null;
  }

  /* ---------- the live session for the open file ---------- */
  function newer(a, b) { return !b || a.c > b.c || (a.c === b.c && a.by > b.by); }

  function session(fileId, mt, saver, getFile, opts) {
    var sb = O.sb(), S = { id: fileId, role: mt.role, meta: mt, units: {}, clock: 0, peers: [], ch: null, alive: true, ring: opts.ring };
    var view = mt.role === "view";
    function snapshot() {
      var f = getFile(), b = {};
      Object.keys(f.body || {}).forEach(function (k) { if ((opts.local || []).indexOf(k) < 0) b[k] = f.body[k]; });
      var out = {}; flat({ title: f.title, body: b }, "", out, opts.html);
      return out;
    }
    var base = "0" + (stamps[fileId] || "");
    var first = snapshot();
    Object.keys(first).forEach(function (k) { S.units[k] = { v: first[k], c: 0, by: base }; });
    S.local = first;      // how it is on this screen, as of the last look

    var joined = false;
    function send(ev, payload) { if (S.ch && S.alive && joined) S.ch.send({ type: "broadcast", event: ev, payload: payload }); }
    function sendUnits(map, ev, extra) {
      // in pieces of about 200 KB, as big messages may not get through
      var piece = {}, size = 0;
      function go() { if (size) { var p = { u: piece, from: me }; if (extra) Object.keys(extra).forEach(function (k) { p[k] = extra[k]; }); send(ev, p); piece = {}; size = 0; } }
      Object.keys(map).forEach(function (k) {
        var u = map[k], s = (u.v ? u.v.length : 4) + k.length + 30;
        if (size && size + s > 200000) go();
        piece[k] = u.pt ? [u.v, u.c, u.by, u.pt] : [u.v, u.c, u.by]; size += s + (u.pt ? 80 : 0);
      });
      go();
    }

    /* what changed here since last time - send it */
    var tmr = null;
    function diff() {
      tmr = null;
      if (!S.alive || view) return;
      if (getFile().id !== S.id) return;
      // only what was changed on this screen (not changes from others still waiting to be shown)
      var cur = snapshot(), out = {}, was = S.local;
      Object.keys(cur).forEach(function (k) {
        if (was[k] !== cur[k]) {
          var pt = typeof was[k] === "string" ? mkPatch(was[k], cur[k]) : null;
          S.units[k] = { v: cur[k], c: ++S.clock, by: me }; out[k] = { v: cur[k], c: S.clock, by: me, pt: pt };
        }
      });
      Object.keys(was).forEach(function (k) {
        if (!(k in cur) && S.units[k] && S.units[k].v !== null) { S.units[k] = { v: null, c: ++S.clock, by: me }; out[k] = S.units[k]; }
      });
      S.local = cur;
      if (Object.keys(out).length) sendUnits(out, "ch");
    }
    function soon() { if (!tmr) tmr = setTimeout(diff, 250); }

    /* changes from someone else */
    function take(u, from) {
      if (!S.alive || !u || typeof u !== "object") return;
      if (tmr) { clearTimeout(tmr); diff(); }            // send ours first, so they're stamped
      var won = {}, any = false, prev = {};
      Object.keys(u).forEach(function (k) {
        var a = u[k]; if (!Array.isArray(a)) return;
        var x = { v: a[0] === null || typeof a[0] === "string" ? a[0] : null, c: +a[1] || 0, by: String(a[2] || "") };
        if (x.c > S.clock) S.clock = x.c;
        var loc = S.units[k];
        if (loc && loc.v === x.v) { if (newer(x, loc)) S.units[k] = x; return; }
        if (a[3] && loc && loc.v !== null && x.v !== null) {
          if (hash(loc.v) === a[3][0]) { prev[k] = loc; S.units[k] = { v: x.v, c: Math.max(loc.c, x.c), by: x.by }; won[k] = 1; any = true; return; }
          var m = fit(loc.v, a[3]);   // we'd both changed it: fit their change into ours
          if (m !== null) { if (m !== loc.v) { prev[k] = loc; S.units[k] = { v: m, c: Math.max(loc.c, x.c), by: loc.by > x.by ? loc.by : x.by }; won[k] = 1; any = true; } return; }
        }
        if (newer(x, loc)) { prev[k] = loc; S.units[k] = x; won[k] = 1; any = true; }
      });
      if (!any) return;
      /* lists: keep anything we'd added that they hadn't seen */
      var fixed = {};
      Object.keys(won).forEach(function (k) {
        var kind = orderKind(k), p = prev[k];
        if (!kind || !p || p.v === null || S.units[k].v === null) return;
        var mine, theirs; try { mine = JSON.parse(p.v); theirs = JSON.parse(S.units[k].v); } catch (e) { return; }
        var stem = k.slice(0, k.lastIndexOf(SEP) + 1) + kind, merged = theirs.slice(), add = false;
        mine.forEach(function (id, i) {
          if (merged.indexOf(id) >= 0) return;
          var item = Object.keys(S.units).some(function (kk) { return (kk === stem + id || kk.indexOf(stem + id + SEP) === 0) && S.units[kk].v !== null; });
          if (!item) return;
          var at = 0;
          for (var j = i - 1; j >= 0; j--) { var w = merged.indexOf(mine[j]); if (w >= 0) { at = w + 1; break; } }
          merged.splice(at, 0, id); add = true;
        });
        if (add) { S.units[k] = { v: JSON.stringify(merged), c: ++S.clock, by: me }; fixed[k] = S.units[k]; }
      });
      if (Object.keys(fixed).length && !view) sendUnits(fixed, "ch");
      // not in the middle of a drag or a tap: wait until it's finished
      if (down) pending = true; else show();
    }
    function show() {
      pending = false;
      if (!S.alive) return;
      var map = {};
      Object.keys(S.units).forEach(function (k) { if (S.units[k].v !== null) map[k] = S.units[k].v; });
      var doc = unflat(map) || {}, f = getFile(), body = doc.body || {};
      (opts.local || []).forEach(function (k) { if (f.body && k in f.body) body[k] = f.body[k]; });
      try { opts.apply(body, typeof doc.title === "string" ? doc.title : f.title); } catch (e) { if (window.console) console.warn("live apply", e); }
      // how it looks here now (the app may tidy it a little) - not sent back as a change
      S.local = snapshot();
      saver.quiet();
    }
    var down = false, pending = false;
    function pdown() { down = true; }
    function pup() {
      down = false;
      if (pending) setTimeout(function () { if (down) return; if (tmr) { clearTimeout(tmr); diff(); } if (pending) show(); }, 40);
    }
    document.addEventListener("pointerdown", pdown, true);
    document.addEventListener("pointerup", pup, true);
    document.addEventListener("pointercancel", pup, true);

    /* who's here */
    var whereNow = "";
    function where() {
      var s = "";
      try {
        if (opts.where) s = opts.where() || "";
        else if (opts.surface) {
          var sel = getSelection(), n = sel.rangeCount ? sel.anchorNode : null;
          while (n && n !== opts.surface && !(n.nodeType === 1 && n.getAttribute("data-b") && n.parentNode === opts.surface)) n = n.parentNode;
          if (n && n !== opts.surface) s = '[data-b="' + n.getAttribute("data-b") + '"]';
        }
      } catch (e) {}
      return /^[\w\-\[\]="'.#: >(),]{0,200}$/.test(s) && s.indexOf("{") < 0 ? s : "";
    }
    function track() { if (S.ch && S.alive && joined) S.ch.track({ n: mt.me || "Someone", w: whereNow, r: S.role }); }
    var wt = null;
    function moved() { clearTimeout(wt); wt = setTimeout(function () { var w = where(); if (w !== whereNow) { whereNow = w; track(); } }, 300); }
    document.addEventListener("selectionchange", moved);
    document.addEventListener("pointerup", moved);
    document.addEventListener("keyup", moved);

    var ch = sb.channel("suite-" + mt.room, { config: { broadcast: { self: false }, presence: { key: me } } });
    S.ch = ch;
    ch.on("broadcast", { event: "ch" }, function (m) { take(m.payload && m.payload.u, m.payload && m.payload.from); });
    ch.on("broadcast", { event: "st" }, function (m) { if (!m.payload || m.payload.to === me || !m.payload.to) take(m.payload && m.payload.u); });
    ch.on("broadcast", { event: "hi" }, function (m) {
      // someone's just opened it: send them what's here (after a moment, so not everyone at once)
      var to = m.payload && m.payload.from;
      setTimeout(function () { if (S.alive) sendUnits(S.units, "st", { to: to }); }, 50 + Math.random() * 400);
    });
    ch.on("presence", { event: "sync" }, function () {
      var st = ch.presenceState(), list = [];
      Object.keys(st).forEach(function (k) { if (k === me) return; var p = st[k] && st[k][0]; if (p) list.push({ key: k, n: String(p.n || "Someone").slice(0, 40), w: String(p.w || ""), r: p.r }); });
      S.peers = list; drawPeers(S);
    });
    ch.subscribe(function (status) {
      if (status === "SUBSCRIBED") { joined = true; send("hi", { from: me }); whereNow = where(); track(); drawPeers(S); }
    });

    S.poke = soon;
    S.stop = function () {
      S.alive = false; clearTimeout(tmr);
      document.removeEventListener("selectionchange", moved); document.removeEventListener("pointerup", moved); document.removeEventListener("keyup", moved);
      document.removeEventListener("pointerdown", pdown, true); document.removeEventListener("pointerup", pup, true); document.removeEventListener("pointercancel", pup, true);
      try { sb.removeChannel(ch); } catch (e) {}
      S.peers = []; drawPeers(null);
    };
    return S;
  }

  /* ---------- the people bar and their highlights ---------- */
  function drawPeers(S) {
    var bar = document.getElementById("ofPeers"), css = document.getElementById("ofPeerCss");
    if (!css) { css = document.createElement("style"); css.id = "ofPeerCss"; document.head.appendChild(css); }
    if (!bar) return;
    var list = S ? S.peers : [];
    if (!list.length) { bar.hidden = true; css.textContent = ""; return; }
    bar.hidden = false;
    bar.innerHTML = list.slice(0, 5).map(function (p) {
      var c = colFor(p.key), ini = p.n.split(/\s+/).map(function (w) { return w.charAt(0); }).join("").slice(0, 2).toUpperCase();
      return '<span class="of-av" style="background:' + c + '" title="' + esc(p.n + (p.r === "view" ? " (viewing)" : " (here now)")) + '">' + esc(ini) + "</span>";
    }).join("") + (list.length > 5 ? '<span class="of-av more">+' + (list.length - 5) + "</span>" : "");
    // what each person is working on gets a ring in their colour
    css.textContent = list.filter(function (p) { return p.w; }).map(function (p) {
      return p.w + (S.ring === "shadow" ? "{filter:drop-shadow(0 0 2px " + colFor(p.key) + ") drop-shadow(0 0 4px " + colFor(p.key) + ")}"
        : "{outline:2px solid " + colFor(p.key) + "!important;outline-offset:2px}");
    }).join("\n");
  }

  /* ---------- viewers can look but not change anything ---------- */
  var blockers = null;
  function viewOnly(on, surface) {
    document.body.classList.toggle("of-viewonly", !!on);
    var ti = document.querySelector(".title-in"); if (ti) ti.readOnly = !!on;
    if (blockers) { blockers.forEach(function (b) { b[0].removeEventListener(b[1], b[2], true); }); blockers = null; }
    if (!on || !surface) return;
    blockers = [];
    ["beforeinput", "keydown", "paste", "drop", "cut", "pointerdown", "dblclick", "dragstart"].forEach(function (ev) {
      var fn = function (e) {
        if (ev === "keydown" && (/^(Arrow|Page|Home|End|Tab|Escape)/.test(e.key) || ((e.ctrlKey || e.metaKey) && /^[ac]$/i.test(e.key)))) return;
        if (ev === "pointerdown" && e.pointerType === "touch") return;   // still lets you scroll
        e.preventDefault(); e.stopPropagation();
        if (ev !== "pointerdown") O.toast("You can view this, but not change it.");
      };
      surface.addEventListener(ev, fn, true); blockers.push([surface, ev, fn]);
    });
  }

  /* ---------- the Share window ---------- */
  function rpc(name, args) {
    var p; try { p = O.sb().rpc(name, args); } catch (e) { return Promise.reject(e); }
    if (!p || typeof p.then !== "function") return Promise.reject(new Error("Sharing isn't available."));
    return p.then(function (r) { if (r.error) throw r.error; return r.data; });
  }
  function shareDialog(L) {
    var f = L.getFile();
    if (!O.isCloud()) { O.toast(O.user() ? "Sharing needs the latest Suite update on the server." : "Sign in on the Launchpad first, then you can share.", true); return; }
    var d = document.createElement("dialog"); d.className = "of-dialog of-share";
    document.body.appendChild(d);
    function close() { d.close(); d.remove(); }
    function paint(mt) {
      var owner = mt.role === "owner";
      d.innerHTML = '<div class="of-dhead"><h2>👥 Share “' + esc(f.title || "Untitled") + '”</h2><button type="button" class="of-x" data-a="x" aria-label="Close">✕</button></div>' +
        '<p class="of-where">' + (owner ? "People you add can open it from their Files, and you'll all see each other's changes as you type."
          : "Shared by " + esc(mt.owner) + ". You can " + (mt.role === "edit" ? "edit it - changes show for everyone straight away." : "view it.")) + "</p>" +
        (owner ? '<form class="of-addp"><input type="email" name="e" placeholder="Their school email" required autocomplete="off" aria-label="Email">' +
          '<select name="r" aria-label="What they can do"><option value="edit">Can edit</option>' + (L.editOnly ? "" : '<option value="view">Can view</option>') + "</select>" +
          '<button type="submit" class="of-btn">Share</button></form><p class="of-err" role="alert"></p>' : "") +
        '<ul class="of-plist"><li><span class="of-av" style="background:' + colFor(mt.owner) + '">' + esc((mt.owner || "?").charAt(0).toUpperCase()) + "</span><span><b>" + esc(mt.owner) + "</b><small>Owner</small></span></li>" +
        (mt.people || []).map(function (p) {
          return '<li><span class="of-av" style="background:' + colFor(p.name) + '">' + esc((p.name || "?").charAt(0).toUpperCase()) + "</span><span><b>" + esc(p.name) + "</b><small>" + esc(p.email || "") + "</small></span>" +
            (owner ? '<select data-p="' + esc(p.id) + '" aria-label="What ' + esc(p.name) + ' can do"><option value="edit"' + (p.role === "edit" ? " selected" : "") + '>Can edit</option><option value="view"' + (p.role === "view" ? " selected" : "") + '>Can view</option></select>' +
              '<button type="button" class="of-del" data-rm="' + esc(p.id) + '" aria-label="Take ' + esc(p.name) + ' off">✕</button>'
              : "<small>" + (p.role === "edit" ? "Can edit" : "Can view") + "</small>") + "</li>";
        }).join("") + "</ul>" +
        '<div class="of-new" style="margin-top:12px"><button type="button" class="of-btn ghost" data-a="link">🔗 Copy link</button>' +
        (!owner ? '<button type="button" class="of-btn ghost" data-a="leave">Remove from my files</button>' : "") + "</div>";
      var fm = d.querySelector(".of-addp"), err = d.querySelector(".of-err");
      if (fm) fm.onsubmit = function (e) {
        e.preventDefault();
        var btn = fm.querySelector("button"); btn.disabled = true; err.textContent = "";
        rpc("share_file", { fid: f.id, who: fm.e.value.trim(), r: fm.r.value }).then(function (m2) {
          O.toast("Shared."); paint(m2); L.reopen();
        }).catch(function (e2) { err.textContent = (e2.message || String(e2)).replace(/^.*?: /, ""); btn.disabled = false; });
      };
      d.querySelectorAll("select[data-p]").forEach(function (s) {
        s.onchange = function () {
          var p = (mt.people || []).filter(function (x) { return x.id === s.dataset.p; })[0];
          rpc("share_file", { fid: f.id, who: p.email, r: s.value }).then(function (m2) { O.toast("Changed."); paint(m2); }).catch(function (e2) { O.toast(e2.message || e2, true); });
        };
      });
    }
    d.onclick = function (e) {
      if (e.target === d) { close(); return; }
      var a = e.target.closest("[data-a],[data-rm]"); if (!a) return;
      if (a.dataset.a === "x") close();
      else if (a.dataset.a === "link") {
        var url = location.origin + location.pathname + "?id=" + encodeURIComponent(f.id);
        (navigator.clipboard ? navigator.clipboard.writeText(url) : Promise.reject()).then(function () { O.toast("Link copied - only people it's shared with can open it."); }, function () { O.toast(url); });
      } else if (a.dataset.a === "leave") {
        rpc("unshare_file", { fid: f.id, who: O.user().id }).then(function () { close(); O.toast("Taken out of your files."); L.left(); }).catch(function (e2) { O.toast(e2.message || e2, true); });
      } else if (a.dataset.rm) {
        rpc("unshare_file", { fid: f.id, who: a.dataset.rm }).then(function (m2) { paint(m2); L.reopen(); }).catch(function (e2) { O.toast(e2.message || e2, true); });
      }
    };
    d.innerHTML = '<p class="of-where">Loading…</p>'; d.showModal();
    // a brand-new file has to be saved before it can be shared
    L.saveNow().then(function () { return rpc("file_meta", { fid: f.id }); }).then(function (mt) {
      if (!mt) throw new Error("Sharing needs the latest Suite update on the server.");
      paint(mt);
    }).catch(function (e) { d.innerHTML = '<div class="of-dhead"><h2>Share</h2><button type="button" class="of-x" data-a="x" aria-label="Close">✕</button></div><p>' + esc(e.message || e) + "</p>"; });
  }

  /* ---------- hooking it into the apps' autosaver ---------- */
  var auto0 = O.autosaver;
  O.blockIds = blockIds; O.patchBlocks = patchBlocks;
  O.autosaver = function (statusEl, getFile, opts) {
    var S = null, role = "owner", wait = [];
    var getF = opts && opts.html && opts.surface ? function () { blockIds(opts.surface); return getFile(); } : getFile;
    var inner = auto0(statusEl, getF, {
      save: opts && opts.save,
      skip: function () { return role === "view" || (opts && opts.save && !getF().id); },
      done: function () { wait.splice(0).forEach(function (fn) { fn(); }); }
    });
    if (!opts) { inner.opened = inner.quiet = function () {}; return inner; }
    var L = {
      getFile: getF, editOnly: !!opts.editOnly,
      saveNow: function () { return new Promise(function (ok) { wait.push(ok); inner.now(); }); },
      reopen: function () { open(true); },
      left: function () { if (S) S.stop(); S = null; role = "owner"; viewOnly(false); setBadge(); }
    };
    // the Share button and the people bar, next to the theme button
    var top = document.querySelector("header.top"), anchor = document.getElementById("themeBtn");
    if (top && !document.getElementById("ofShare")) {
      var pb = document.createElement("span"); pb.id = "ofPeers"; pb.className = "of-peers"; pb.hidden = true;
      var sbtn = document.createElement("button"); sbtn.type = "button"; sbtn.className = "tbtn of-sharebtn"; sbtn.id = "ofShare";
      sbtn.innerHTML = '👥<span class="lbl"> ' + esc(opts.shareLabel || "Share") + "</span>";
      top.insertBefore(pb, anchor || null); top.insertBefore(sbtn, anchor || null);
      sbtn.onclick = function () { shareDialog(L); };
    }
    setBadge();
    function setBadge() {
      var st = document.getElementById("ofShare"); if (!st) return;
      st.innerHTML = role === "view" ? '👁<span class="lbl"> View only</span>' : role === "edit" ? '👥<span class="lbl"> Shared</span>' : '👥<span class="lbl"> ' + esc(opts.shareLabel || "Share") + "</span>";
      st.hidden = !getF().id;
    }
    var openN = 0;
    function open(force) {
      var f = getF(), n = ++openN;
      if (S && S.id === f.id && !force) return;
      if (S) { S.stop(); S = null; }
      role = "owner"; viewOnly(false); setBadge();
      if (!O.isCloud() || !O.sb() || !f.id) return;
      rpc("file_meta", { fid: f.id }).then(function (mt) {
        if (n !== openN || !mt || !mt.room || getF().id !== f.id) return;
        if (mt.locked) { lockedFile(f, mt); return; }
        role = mt.role; setBadge();
        if (role === "view") { viewOnly(true, opts.surface); statusEl.textContent = "👁 View only - shared by " + mt.owner; }
        else if (role === "edit") statusEl.textContent = "👥 Shared by " + mt.owner;
        S = session(f.id, mt, out, getF, opts);
      }).catch(function () {});
    }
    /* pay as you go: a file whose storage isn't paid for is locked - you can
       look at the copy kept on this device, or pay what's owed to unlock it */
    function lockedFile(f, mt) {
      role = "view"; setBadge(); viewOnly(true, opts.surface);
      var own = mt.role === "owner";
      statusEl.textContent = "🔒 Locked" + (own ? "" : " - shared by " + mt.owner);
      O.bill().then(function (b) {
        var info = b && b.byId[f.id], owed = info ? info.owed : null, have = b ? b.pence : null;
        O.get(f.id).then(function (g) {
          var d = document.createElement("dialog"); d.className = "of-dialog";
          d.innerHTML = '<div class="of-dhead"><h2>🔒 This file is locked</h2><button type="button" class="of-x" data-a="x" aria-label="Close">✕</button></div>' +
            (own ? "<p>Its storage hasn't been paid for, so it's locked. Nothing has been deleted.</p>" +
              (owed != null ? "<p>Unlocking it costs <b>" + owed + "p</b>" + (have != null ? " - you have <b>" + have + "p</b>" : "") + ". After that it's " + (b ? b.cost : 10) + "p a day (free in the school holidays).</p>" : "")
              : "<p>" + esc(mt.owner) + " hasn't paid for this file's storage, so it's locked. Ask them to open it and unlock it.</p>") +
            '<p class="of-where">' + (g && g.keptCopy ? "You're looking at the copy kept on this device (from " + new Date(g.keptCopy).toLocaleString("en-GB", { day: "numeric", month: "short", hour: "2-digit", minute: "2-digit" }) + "). You can read it and download it, but not change it."
              : "There's no copy of it on this device.") + "</p>" + '<p class="of-err" role="alert" style="color:#dc2626"></p>' +
            '<div class="of-new">' + (own ? '<button type="button" class="of-btn" data-a="pay">🔓 Unlock' + (owed != null ? " for " + owed + "p" : "") + "</button>" : "") +
            '<button type="button" class="of-btn ghost" data-a="x">' + (g && g.keptCopy ? "Look at my copy" : "Close") + "</button></div>";
          document.body.appendChild(d); d.showModal();
          d.onclick = function (e) {
            var bt = e.target.closest("[data-a]"); if (e.target === d) { d.close(); d.remove(); return; } if (!bt) return;
            if (bt.dataset.a === "x") { d.close(); d.remove(); return; }
            if (bt.dataset.a === "pay") {
              bt.disabled = true;
              O.unlock(f.id).then(function () { O.toast("Unlocked."); location.reload(); })
                .catch(function (er) { bt.disabled = false; d.querySelector(".of-err").textContent = (er.message || String(er)).replace(/^.*?: /, ""); });
            }
          };
        });
      });
    }
    var out = {
      poke: function () { inner.poke(); if (S) S.poke(); },
      now: function () { inner.now(); if (S) S.poke(); },
      quiet: function () { inner.poke(); },          // save what came in, without sending it back
      opened: function () { open(false); }
    };
    return out;
  };
})();
