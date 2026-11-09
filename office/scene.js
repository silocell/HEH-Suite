/* =====================================================================
   Launchpad Suite - drawing a design page or a video scene on a canvas.
   Used by Design (Canva-style) and Vids, for the screen and for export,
   so what you see is exactly what you get.

   A page/scene: { bg: {c:"#fff"} | {g:["#a","#b",angle]} | {img:src}, els: [...] }
   An element:   { id, t: "text"|"shape"|"img"|"emoji", x, y, w, h, r (degrees),
                   o (0-1), fx (flip), ...what it is }
   ===================================================================== */
(function () {
  /* Fonts from Google Fonts (the same on every computer, Chromebooks too),
     plus two that every computer has. [name, CSS family list] */
  var FONTS = [
    ["Inter", "'Inter', system-ui, sans-serif"], ["Poppins", "'Poppins', sans-serif"], ["Montserrat", "'Montserrat', sans-serif"],
    ["Nunito", "'Nunito', sans-serif"], ["Fredoka", "'Fredoka', sans-serif"], ["Oswald", "'Oswald', sans-serif"],
    ["Anton", "'Anton', Impact, sans-serif"], ["Bebas Neue", "'Bebas Neue', Impact, sans-serif"], ["Archivo Black", "'Archivo Black', 'Arial Black', sans-serif"],
    ["Bangers", "'Bangers', cursive"], ["Press Start 2P", "'Press Start 2P', monospace"],
    ["Playfair Display", "'Playfair Display', Georgia, serif"], ["Merriweather", "'Merriweather', Georgia, serif"], ["Lora", "'Lora', Georgia, serif"],
    ["Roboto Slab", "'Roboto Slab', serif"], ["Courier Prime", "'Courier Prime', 'Courier New', monospace"],
    ["Pacifico", "'Pacifico', cursive"], ["Dancing Script", "'Dancing Script', cursive"], ["Lobster", "'Lobster', cursive"],
    ["Caveat", "'Caveat', cursive"], ["Permanent Marker", "'Permanent Marker', cursive"], ["Comic Neue", "'Comic Neue', 'Comic Sans MS', cursive"],
    ["Arial", "Arial, Helvetica, sans-serif"], ["Georgia", "Georgia, 'Times New Roman', serif"]
  ];
  var FONTS_URL = "https://fonts.googleapis.com/css2?family=Inter:wght@400;700&family=Poppins:wght@400;700&family=Montserrat:wght@400;800&family=Nunito:wght@400;800" +
    "&family=Fredoka:wght@400;600&family=Oswald:wght@400;700&family=Anton&family=Bebas+Neue&family=Archivo+Black&family=Bangers&family=Press+Start+2P" +
    "&family=Playfair+Display:ital,wght@0,400;0,700;1,400&family=Merriweather:ital,wght@0,400;0,700;1,400&family=Lora:ital,wght@0,400;0,700;1,400" +
    "&family=Roboto+Slab:wght@400;700&family=Courier+Prime:wght@400;700&family=Pacifico&family=Dancing+Script:wght@400;700&family=Lobster" +
    "&family=Caveat:wght@400;700&family=Permanent+Marker&family=Comic+Neue:wght@400;700&display=swap";
  if (!document.querySelector('link[data-scene-fonts]')) {
    var lk = document.createElement("link"); lk.rel = "stylesheet"; lk.href = FONTS_URL; lk.setAttribute("data-scene-fonts", ""); document.head.appendChild(lk);
  }
  var FONT = {}; FONTS.forEach(function (f) { FONT[f[0]] = f[1]; });
  var SHAPES = ["rect", "round", "circle", "tri", "star", "heart", "arrow", "hex", "speech", "line", "blob", "badge"];

  function shapePath(ctx, k, w, h, rad) {
    ctx.beginPath();
    var i, a;
    switch (k) {
      case "round": rad = rad || Math.min(w, h) * 0.2; /* falls through */
      case "rect":
        var r = Math.max(0, Math.min(rad || 0, w / 2, h / 2));
        ctx.moveTo(r, 0); ctx.lineTo(w - r, 0); ctx.quadraticCurveTo(w, 0, w, r); ctx.lineTo(w, h - r); ctx.quadraticCurveTo(w, h, w - r, h);
        ctx.lineTo(r, h); ctx.quadraticCurveTo(0, h, 0, h - r); ctx.lineTo(0, r); ctx.quadraticCurveTo(0, 0, r, 0); break;
      case "circle": ctx.ellipse(w / 2, h / 2, w / 2, h / 2, 0, 0, Math.PI * 2); break;
      case "tri": ctx.moveTo(w / 2, 0); ctx.lineTo(w, h); ctx.lineTo(0, h); break;
      case "star":
        for (i = 0; i < 10; i++) { a = -Math.PI / 2 + i * Math.PI / 5; var rr = i % 2 ? 0.4 : 1; ctx.lineTo(w / 2 + Math.cos(a) * w / 2 * rr, h / 2 + Math.sin(a) * h / 2 * rr * 1.05 + h * 0.03); } break;
      case "heart":
        ctx.moveTo(w / 2, h * 0.28);
        ctx.bezierCurveTo(w * 0.5, 0, 0, 0, 0, h * 0.32); ctx.bezierCurveTo(0, h * 0.62, w * 0.35, h * 0.78, w / 2, h);
        ctx.bezierCurveTo(w * 0.65, h * 0.78, w, h * 0.62, w, h * 0.32); ctx.bezierCurveTo(w, 0, w * 0.5, 0, w / 2, h * 0.28); break;
      case "arrow": ctx.moveTo(0, h * 0.3); ctx.lineTo(w * 0.62, h * 0.3); ctx.lineTo(w * 0.62, 0); ctx.lineTo(w, h / 2); ctx.lineTo(w * 0.62, h); ctx.lineTo(w * 0.62, h * 0.7); ctx.lineTo(0, h * 0.7); break;
      case "hex": for (i = 0; i < 6; i++) { a = i * Math.PI / 3; ctx.lineTo(w / 2 + Math.cos(a) * w / 2, h / 2 + Math.sin(a) * h / 2); } break;
      case "speech":
        var sr = Math.min(w, h * 0.75) * 0.18, bh = h * 0.78;
        ctx.moveTo(sr, 0); ctx.lineTo(w - sr, 0); ctx.quadraticCurveTo(w, 0, w, sr); ctx.lineTo(w, bh - sr); ctx.quadraticCurveTo(w, bh, w - sr, bh);
        ctx.lineTo(w * 0.38, bh); ctx.lineTo(w * 0.18, h); ctx.lineTo(w * 0.22, bh); ctx.lineTo(sr, bh); ctx.quadraticCurveTo(0, bh, 0, bh - sr); ctx.lineTo(0, sr); ctx.quadraticCurveTo(0, 0, sr, 0); break;
      case "line": ctx.moveTo(0, h / 2); ctx.lineTo(w, h / 2); return;
      case "blob":
        ctx.moveTo(w * 0.5, 0); ctx.bezierCurveTo(w * 0.85, 0, w, h * 0.25, w * 0.95, h * 0.5); ctx.bezierCurveTo(w * 0.9, h * 0.8, w * 0.7, h, w * 0.45, h * 0.97);
        ctx.bezierCurveTo(w * 0.15, h * 0.94, 0, h * 0.7, w * 0.04, h * 0.42); ctx.bezierCurveTo(w * 0.08, h * 0.15, w * 0.25, 0, w * 0.5, 0); break;
      case "badge":
        for (i = 0; i < 32; i++) { a = i * Math.PI / 16; var br = i % 2 ? 0.88 : 1; ctx.lineTo(w / 2 + Math.cos(a) * w / 2 * br, h / 2 + Math.sin(a) * h / 2 * br); } break;
      default: ctx.rect(0, 0, w, h);
    }
    ctx.closePath();
  }

  /* a fill can be a colour or a gradient "lin:#a,#b,angle" */
  function paint(ctx, f, w, h) {
    if (typeof f !== "string" || !f) return "transparent";
    if (f.indexOf("lin:") === 0) {
      var p = f.slice(4).split(","), ang = (+p[2] || 135) * Math.PI / 180;
      var cx = w / 2, cy = h / 2, L = Math.abs(w * Math.sin(ang)) / 2 + Math.abs(h * Math.cos(ang)) / 2;
      var g = ctx.createLinearGradient(cx - Math.sin(ang) * L, cy + Math.cos(ang) * L, cx + Math.sin(ang) * L, cy - Math.cos(ang) * L);
      g.addColorStop(0, p[0] || "#000"); g.addColorStop(1, p[1] || "#fff"); return g;
    }
    return f;
  }
  function cssFill(f) {
    if (typeof f !== "string" || !f) return "transparent";
    if (f.indexOf("lin:") === 0) { var p = f.slice(4).split(","); return "linear-gradient(" + ((+p[2] || 135)) + "deg," + (p[0] || "#000") + "," + (p[1] || "#fff") + ")"; }
    return f;
  }

  function fontOf(el, scale) {
    return (el.italic ? "italic " : "") + (el.bold ? "700 " : "400 ") + Math.max(1, (el.size || 32) * (scale || 1)) + "px " + (FONT[el.font] || FONT.Inter);
  }
  /* A font is fetched the first time it's used; draw again once it's here. */
  var fload = {};
  function needFont(el, cb) {
    if (!document.fonts || !document.fonts.check) return;
    var f = fontOf(el);
    try { if (document.fonts.check(f)) return; } catch (e) { return; }
    if (fload[f] === true) return;
    if (fload[f]) { if (cb && fload[f].indexOf(cb) < 0) fload[f].push(cb); return; }
    fload[f] = cb ? [cb] : [];
    document.fonts.load(f).then(function () { var c = fload[f] || []; fload[f] = true; c.forEach(function (fn) { try { fn(); } catch (e) {} }); }, function () { fload[f] = true; });
  }
  function fontsReady(els) {
    if (!document.fonts) return Promise.resolve();
    return Promise.all((els || []).filter(function (e) { return e && e.t === "text"; }).map(function (e) { return document.fonts.load(fontOf(e)).catch(function () {}); }))
      .then(function () { return document.fonts.ready; });
  }
  /* Text broken into lines that fit the box (by words, then by letters). */
  function wrap(ctx, el) {
    ctx.font = fontOf(el);
    var ls = el.ls || 0, out = [], maxW = Math.max(10, el.w);
    function width(s) { return ctx.measureText(s).width + ls * Math.max(0, s.length - 1); }
    String(el.text == null ? "" : el.text).replace(/\r/g, "").split("\n").forEach(function (para) {
      if (el.upper) para = para.toUpperCase();
      var words = para.split(/(\s+)/), line = "";
      words.forEach(function (wd) {
        if (!wd) return;
        var test = line + wd;
        if (width(test) <= maxW || !line.trim()) {
          if (width(test) > maxW && !line) {
            // one long word: split it
            for (var i = 0; i < wd.length; i++) { if (width(line + wd[i]) > maxW && line) { out.push(line); line = ""; } line += wd[i]; }
          } else line = test;
        } else { out.push(line.replace(/\s+$/, "")); line = /^\s+$/.test(wd) ? "" : wd; }
      });
      out.push(line.replace(/\s+$/, ""));
    });
    return out;
  }
  function textHeight(ctx, el) { return Math.ceil(wrap(ctx, el).length * (el.size || 32) * (el.lh || 1.2) + (el.bgc ? (el.size || 32) * 0.5 : 0)); }

  var imgs = {}, waiting = [];
  /* pictures are kept loaded by their address; when one arrives, onload is called */
  function img(src, onload) {
    if (!src || src.indexOf("idb:") === 0) return null;   // things kept on the device are given by the app
    var m = imgs[src];
    if (m) { if (!m.complete && onload && m._cbs.indexOf(onload) < 0) m._cbs.push(onload); return m.complete && m.naturalWidth ? m : null; }
    m = new Image(); m._cbs = onload ? [onload] : []; m.decoding = "async";
    m.onload = m.onerror = function () { var c = m._cbs; m._cbs = []; c.forEach(function (f) { try { f(); } catch (e) {} }); };
    if (!/^data:|^blob:/.test(src)) m.crossOrigin = "anonymous";
    m.src = src; imgs[src] = m;
    return null;
  }
  function ready(srcs) {
    return Promise.all(srcs.filter(Boolean).map(function (s) {
      return new Promise(function (ok) { var m = img(s, ok); if (m || (imgs[s] && imgs[s].complete)) ok(); setTimeout(ok, 8000); });
    }));
  }

  function filterOf(el) {
    var f = el.fl || {}, s = [];
    if (f.b != null && f.b !== 100) s.push("brightness(" + f.b + "%)");
    if (f.c != null && f.c !== 100) s.push("contrast(" + f.c + "%)");
    if (f.s != null && f.s !== 100) s.push("saturate(" + f.s + "%)");
    if (f.g) s.push("grayscale(" + f.g + "%)");
    if (f.sp) s.push("sepia(" + f.sp + "%)");
    if (f.bl) s.push("blur(" + f.bl + "px)");
    return s.join(" ") || "none";
  }

  /* draw one element. a = { o, dx, dy, s, type } from an animation (Vids) */
  function drawEl(ctx, el, opt) {
    opt = opt || {}; var a = opt.a || {};
    var o = (el.o == null ? 1 : el.o) * (a.o == null ? 1 : a.o);
    if (o <= 0.001) return;
    ctx.save();
    ctx.globalAlpha = o;
    var cx = el.x + el.w / 2 + (a.dx || 0), cy = el.y + el.h / 2 + (a.dy || 0);
    ctx.translate(cx, cy);
    if (el.r || a.r) ctx.rotate(((el.r || 0) + (a.r || 0)) * Math.PI / 180);
    var sc = a.s == null ? 1 : a.s; if (sc !== 1) ctx.scale(sc, sc);
    if (el.fx) ctx.scale(-1, 1);
    ctx.translate(-el.w / 2, -el.h / 2);
    if (el.sh) { ctx.shadowColor = "rgba(0,0,0,.35)"; ctx.shadowBlur = Math.max(4, Math.min(el.w, el.h) * 0.06); ctx.shadowOffsetY = Math.max(2, Math.min(el.w, el.h) * 0.03); }
    var w = el.w, h = el.h;
    if (el.t === "shape") {
      shapePath(ctx, el.shape, w, h, el.rad);
      if (el.shape === "line") { ctx.strokeStyle = el.stroke || el.fill || "#111"; ctx.lineWidth = el.sw || 4; ctx.lineCap = "round"; ctx.stroke(); }
      else {
        if (el.fill && el.fill !== "none") { ctx.fillStyle = paint(ctx, el.fill, w, h); ctx.fill(); }
        if (el.sw && el.stroke) { ctx.shadowColor = "transparent"; ctx.strokeStyle = el.stroke; ctx.lineWidth = el.sw; ctx.lineJoin = "round"; ctx.stroke(); }
      }
    } else if (el.t === "img" || el.t === "video") {
      var m = opt.media || (el.t === "img" ? img(el.src, opt.onload) : null);
      if (m) {
        var iw = m.videoWidth || m.naturalWidth || m.width, ih = m.videoHeight || m.naturalHeight || m.height;
        if (iw && ih) {
          var z = Math.max(w / iw, h / ih) * (el.zoom || 1), dw = iw * z, dh = ih * z;
          var ox = (w - dw) * (el.px == null ? 0.5 : el.px), oy = (h - dh) * (el.py == null ? 0.5 : el.py);
          ctx.save();
          shapePath(ctx, el.mask || "rect", w, h, el.rad); ctx.clip();
          if (el.sh) { ctx.shadowColor = "transparent"; }
          var f = filterOf(el); if (f !== "none" && "filter" in ctx) ctx.filter = f;
          ctx.drawImage(m, ox, oy, dw, dh);
          ctx.restore();
          if (el.sw && el.stroke) { shapePath(ctx, el.mask || "rect", w, h, el.rad); ctx.strokeStyle = el.stroke; ctx.lineWidth = el.sw; ctx.stroke(); }
        }
      } else if (!opt.clean) {
        ctx.fillStyle = "rgba(127,127,127,.18)"; ctx.fillRect(0, 0, w, h);
      }
    } else if (el.t === "emoji") {
      ctx.font = Math.round(h * 0.86) + "px 'Apple Color Emoji','Segoe UI Emoji','Noto Color Emoji',sans-serif";
      ctx.textAlign = "center"; ctx.textBaseline = "middle";
      ctx.fillText(el.ch || "⭐", w / 2, h / 2 + h * 0.04);
    } else if (el.t === "text") {
      needFont(el, opt.onload);
      var lines = wrap(ctx, el), size = el.size || 32, lh = size * (el.lh || 1.2);
      var shown = lines;
      if (a.chars != null) { // typewriter: only so many letters
        var left = a.chars; shown = lines.map(function (l) { var s = l.slice(0, Math.max(0, left)); left -= l.length; return s; });
      }
      ctx.font = fontOf(el); ctx.textBaseline = "alphabetic";
      if ("letterSpacing" in ctx) ctx.letterSpacing = (el.ls || 0) + "px";
      var padY = el.bgc ? size * 0.25 : 0;
      if (el.bgc) {
        ctx.save(); ctx.shadowColor = "transparent";
        ctx.fillStyle = el.bgc;
        shapePath(ctx, "rect", w, Math.max(h, lines.length * lh + padY * 2), Math.min(size * 0.35, 24)); ctx.fill(); ctx.restore();
      }
      shown.forEach(function (l, i) {
        var lw = ctx.measureText(l).width, x = el.align === "center" ? (w - lw) / 2 : el.align === "right" ? w - lw : 0;
        if (el.bgc && el.align !== "center") x += el.align === "right" ? -size * 0.3 : size * 0.3;
        var y = padY + i * lh + size * 0.92 + (lh - size) / 2;
        if (el.stroke && el.sw) { ctx.lineJoin = "round"; ctx.strokeStyle = el.stroke; ctx.lineWidth = el.sw * 2; ctx.strokeText(l, x, y); }
        ctx.fillStyle = paint(ctx, el.color || "#111", w, h);
        ctx.fillText(l, x, y);
        if (el.under && l) { ctx.fillRect(x, y + size * 0.1, lw, Math.max(1, size / 16)); }
      });
    }
    ctx.restore();
  }

  function drawBg(ctx, bg, W, H, onload) {
    bg = bg || {};
    ctx.save();
    if (bg.img) {
      ctx.fillStyle = "#fff"; ctx.fillRect(0, 0, W, H);
      var m = img(bg.img, onload);
      if (m) { var z = Math.max(W / m.naturalWidth, H / m.naturalHeight); ctx.drawImage(m, (W - m.naturalWidth * z) / 2, (H - m.naturalHeight * z) / 2, m.naturalWidth * z, m.naturalHeight * z); }
      if (bg.dim) { ctx.fillStyle = "rgba(0,0,0," + bg.dim + ")"; ctx.fillRect(0, 0, W, H); }
    } else {
      ctx.fillStyle = paint(ctx, bg.g ? "lin:" + bg.g.join(",") : (bg.c || "#ffffff"), W, H);
      ctx.fillRect(0, 0, W, H);
    }
    ctx.restore();
  }
  /* a whole page: background then everything on it, in order */
  function drawPage(ctx, page, W, H, opt) {
    opt = opt || {};
    drawBg(ctx, page.bg, W, H, opt.onload);
    (page.els || []).forEach(function (el) {
      if (opt.skip && opt.skip === el.id) return;
      drawEl(ctx, el, { onload: opt.onload, clean: opt.clean, a: opt.anim ? opt.anim(el) : null, media: opt.media ? opt.media(el) : null });
    });
  }
  /* is a point (in page units) on this element? (allowing for turning it) */
  function hit(el, px, py, pad) {
    pad = pad || 0;
    var cx = el.x + el.w / 2, cy = el.y + el.h / 2, a = -(el.r || 0) * Math.PI / 180;
    var dx = px - cx, dy = py - cy, x = dx * Math.cos(a) - dy * Math.sin(a), y = dx * Math.sin(a) + dy * Math.cos(a);
    return Math.abs(x) <= el.w / 2 + pad && Math.abs(y) <= el.h / 2 + pad;
  }

  /* safety: everything that comes from a file or from someone else */
  function num(v, d, lo, hi) { v = +v; if (!isFinite(v)) v = d; return Math.max(lo, Math.min(hi, v)); }
  function colour(v, d) { v = String(v == null ? "" : v); return /^(#[0-9a-f]{3,8}|rgba?\([\d.,\s%]+\)|transparent|none|lin:#[0-9a-f]{3,8},#[0-9a-f]{3,8}(,-?\d{1,3})?)$/i.test(v) ? v : d; }
  function src(v) { v = String(v || ""); return /^data:image\/(png|jpe?g|gif|webp|svg\+xml);base64,[a-z0-9+\/=]+$/i.test(v) || /^idb:[a-z0-9]{1,32}$/i.test(v) || /^https:\/\/[^\s"'<>]+$/i.test(v) ? v : ""; }
  function cleanEl(e, i) {
    if (!e || typeof e !== "object") return null;
    var t = ["text", "shape", "img", "emoji", "video"].indexOf(e.t) >= 0 ? e.t : null; if (!t) return null;
    var o = { id: String(e.id || "").replace(/[^a-z0-9]/gi, "").slice(0, 16) || ("e" + i + Math.random().toString(36).slice(2, 6)), t: t,
      x: num(e.x, 0, -20000, 20000), y: num(e.y, 0, -20000, 20000), w: num(e.w, 100, 1, 20000), h: num(e.h, 100, 1, 20000),
      r: num(e.r, 0, -360, 360), o: num(e.o, 1, 0, 1) };
    if (e.fx) o.fx = true; if (e.sh) o.sh = true; if (e.lock) o.lock = true;
    if (e.anim) o.anim = String(e.anim).replace(/[^a-z]/g, "").slice(0, 12);
    if (e.at != null) o.at = num(e.at, 0, 0, 600);
    if (t === "text") {
      o.text = String(e.text == null ? "" : e.text).slice(0, 5000); o.font = FONT[e.font] ? e.font : "Inter"; o.size = num(e.size, 32, 4, 800);
      o.color = colour(e.color, "#111111"); if (e.bold) o.bold = true; if (e.italic) o.italic = true; if (e.under) o.under = true; if (e.upper) o.upper = true;
      o.align = /^(left|center|right)$/.test(e.align) ? e.align : "left"; o.lh = num(e.lh, 1.2, 0.6, 4); o.ls = num(e.ls, 0, -10, 80);
      if (e.bgc) o.bgc = colour(e.bgc, ""); if (e.stroke) { o.stroke = colour(e.stroke, ""); o.sw = num(e.sw, 2, 0, 40); }
    } else if (t === "shape") {
      o.shape = SHAPES.indexOf(e.shape) >= 0 ? e.shape : "rect"; o.fill = colour(e.fill, "#7c3aed");
      if (e.stroke) o.stroke = colour(e.stroke, ""); o.sw = num(e.sw, 0, 0, 80); o.rad = num(e.rad, 0, 0, 2000);
    } else if (t === "img") {
      o.src = src(e.src); if (!o.src) return null; o.rad = num(e.rad, 0, 0, 2000); o.zoom = num(e.zoom, 1, 1, 6);
      o.px = num(e.px, 0.5, 0, 1); o.py = num(e.py, 0.5, 0, 1);
      if (e.mask && SHAPES.indexOf(e.mask) >= 0) o.mask = e.mask; if (e.stroke) { o.stroke = colour(e.stroke, ""); o.sw = num(e.sw, 0, 0, 80); }
      if (e.fl && typeof e.fl === "object") { o.fl = {}; ["b", "c", "s", "g", "sp", "bl"].forEach(function (k) { if (e.fl[k] != null) o.fl[k] = num(e.fl[k], k === "b" || k === "c" || k === "s" ? 100 : 0, 0, k === "bl" ? 40 : 300); }); }
    } else if (t === "video") {
      o.src = String(e.src || "").replace(/[^a-z0-9:]/gi, "").slice(0, 40); o.rad = num(e.rad, 0, 0, 2000);
      o.from = num(e.from, 0, 0, 36000); o.vol = num(e.vol, 1, 0, 1); if (e.mute) o.mute = true;
    } else if (t === "emoji") {
      o.ch = String(e.ch || "⭐").slice(0, 16);
    }
    return o;
  }
  function cleanBg(b) {
    b = b && typeof b === "object" ? b : {};
    if (b.img && src(b.img)) return { img: src(b.img), dim: num(b.dim, 0, 0, 0.9) };
    if (Array.isArray(b.g)) return { g: [colour(b.g[0], "#7c3aed"), colour(b.g[1], "#db2777"), num(b.g[2], 135, -360, 360)] };
    return { c: colour(b.c, "#ffffff") };
  }

  window.Scene = { FONTS: FONTS, FONT: FONT, SHAPES: SHAPES, shapePath: shapePath, paint: paint, cssFill: cssFill, fontOf: fontOf,
    wrap: wrap, textHeight: textHeight, needFont: needFont, fontsReady: fontsReady, img: img, ready: ready, drawEl: drawEl, drawBg: drawBg, drawPage: drawPage, hit: hit,
    cleanEl: cleanEl, cleanBg: cleanBg, colour: colour, num: num, imgSrc: src };
})();
