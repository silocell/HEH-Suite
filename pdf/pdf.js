/* Launchpad PDF - read, mark up, sign, fill in, reorder, merge and split
   PDFs, and turn pictures into a PDF. Everything happens on this device:
   pdf.js (Mozilla) draws the pages and pdf-lib writes the new PDF. */
import * as pdfjs from "/pdf/lib/pdf.min.mjs";
pdfjs.GlobalWorkerOptions.workerSrc = "/pdf/lib/pdf.worker.min.mjs";
const { PDFDocument, rgb, degrees, StandardFonts, PDFTextField, PDFCheckBox, PDFDropdown, PDFRadioGroup, PDFName, PDFDict } = window.PDFLib;
const $ = (id) => document.getElementById(id);

let docs = [];          // { bytes, js, name, fields: [...] }
let pages = [];         // { id, d, i, rot, inh, w, h, ann: [] }   d = -1 for a blank page
let cur = 0, zoom = 1, tool = "select", signature = null, undo = [];
let uid = 0;
const formVals = {};    // "doc:field" -> value

Office.themeToggle($("themeBtn"));

/* ---------- opening ---------- */
async function addPdf(bytes, name) {
  const js = await pdfjs.getDocument({ data: bytes.slice(0), isEvalSupported: false, enableXfa: false }).promise;
  const d = docs.length;
  const entry = { bytes, js, name, fields: [] };
  docs.push(entry);
  for (let i = 0; i < js.numPages; i++) {
    const p = await js.getPage(i + 1), v = p.getViewport({ scale: 1, rotation: 0 });
    pages.push({ id: ++uid, d, i, rot: 0, inh: p.rotate || 0, w: v.width, h: v.height, ann: [] });
  }
  try { // form fields, so they can be filled in
    const lib = await PDFDocument.load(bytes, { ignoreEncryption: true });
    lib.getForm().getFields().forEach((f) => {
      const n = f.getName();
      if (f instanceof PDFTextField) entry.fields.push({ n, t: "text", v: f.getText() || "" });
      else if (f instanceof PDFCheckBox) entry.fields.push({ n, t: "check", v: f.isChecked() });
      else if (f instanceof PDFDropdown) entry.fields.push({ n, t: "pick", v: (f.getSelected() || [])[0] || "", o: f.getOptions() });
      else if (f instanceof PDFRadioGroup) entry.fields.push({ n, t: "pick", v: f.getSelected() || "", o: f.getOptions() });
    });
  } catch (e) { /* not every PDF can be read by pdf-lib - it can still be shown */ }
  $("formBtn").hidden = !docs.some((x) => x.fields.length);
}
async function openFiles(files, merge) {
  const list = Array.from(files || []);
  if (!list.length) return;
  if (!merge) { docs = []; pages = []; undo = []; cur = 0; $("formBox").hidden = true; }
  for (const f of list) {
    if (!/pdf$/i.test(f.type) && !/\.pdf$/i.test(f.name)) { Office.toast(f.name + " isn't a PDF.", true); continue; }
    if (f.size > 200 * 1024 * 1024) { Office.toast(f.name + " is too big (over 200 MB).", true); continue; }
    try {
      $("status").textContent = "Opening " + f.name + "…";
      await addPdf(new Uint8Array(await f.arrayBuffer()), f.name);
      if (!merge && docs.length === 1) $("title").value = f.name.replace(/\.pdf$/i, "");
    } catch (e) {
      Office.toast(/password/i.test(e.message || "") ? f.name + " is locked with a password." : "Couldn't open " + f.name + ": " + (e.message || e), true);
    }
  }
  $("status").textContent = "Your PDFs stay on this device";
  draw();
  if (merge && list.length) Office.toast("Added " + list.length + (list.length === 1 ? " PDF" : " PDFs") + " at the end.");
}
async function picturesToPdf(files) {
  const imgs = Array.from(files || []).filter((f) => /^image\//.test(f.type));
  if (!imgs.length) return;
  $("status").textContent = "Making a PDF from " + imgs.length + " pictures…";
  const out = await PDFDocument.create();
  for (const f of imgs) {
    // every picture goes through a canvas, so any type (HEIC from an iPad too) becomes a JPEG
    const src = await Office.shrinkImage(f, 2400).catch(() => null);
    if (!src) continue;
    const bytes = Uint8Array.from(atob(src.split(",")[1]), (c) => c.charCodeAt(0));
    const im = /png/.test(src.slice(0, 30)) ? await out.embedPng(bytes) : await out.embedJpg(bytes);
    const A4 = [595.28, 841.89], land = im.width > im.height, pw = land ? A4[1] : A4[0], ph = land ? A4[0] : A4[1];
    const k = Math.min((pw - 40) / im.width, (ph - 40) / im.height);
    const p = out.addPage([pw, ph]);
    p.drawImage(im, { x: (pw - im.width * k) / 2, y: (ph - im.height * k) / 2, width: im.width * k, height: im.height * k });
  }
  const bytes = await out.save();
  const had = pages.length;
  await addPdf(bytes, "Pictures.pdf");
  if (!had) $("title").value = "Pictures";
  $("status").textContent = "Your PDFs stay on this device";
  draw(); Office.toast("Made " + imgs.length + (imgs.length === 1 ? " page" : " pages") + " from your pictures.");
}
["openIn", "openIn2"].forEach((id) => $(id).addEventListener("change", (e) => { openFiles(e.target.files, false); e.target.value = ""; }));
$("mergeIn").addEventListener("change", (e) => { openFiles(e.target.files, pages.length > 0); e.target.value = ""; });
["imgIn", "imgIn2"].forEach((id) => $(id).addEventListener("change", (e) => { picturesToPdf(e.target.files); e.target.value = ""; }));
const view = $("view");
view.addEventListener("dragover", (e) => { e.preventDefault(); view.classList.add("drop"); });
view.addEventListener("dragleave", () => view.classList.remove("drop"));
view.addEventListener("drop", (e) => {
  e.preventDefault(); view.classList.remove("drop");
  const fs = Array.from(e.dataTransfer.files || []);
  if (fs.some((f) => /^image\//.test(f.type))) picturesToPdf(fs); else openFiles(fs, pages.length > 0);
});
Office.setImporter((f) => { if (/^image\//.test(f.type)) picturesToPdf([f]); else openFiles([f], false); });

/* ---------- drawing the pages ---------- */
function R(p) { return ((p.inh + p.rot) % 360 + 360) % 360; }
function dispSize(p) { const r = R(p); return r % 180 ? [p.h, p.w] : [p.w, p.h]; }
function matrix(p) {
  const r = R(p), w = p.w, h = p.h;
  return r === 90 ? `matrix(0,1,-1,0,${h},0)` : r === 180 ? `matrix(-1,0,0,-1,${w},${h})` : r === 270 ? `matrix(0,-1,1,0,0,${w})` : "";
}
function fromDisp(p, X, Y) {
  const r = R(p), w = p.w, h = p.h;
  return r === 90 ? [Y, h - X] : r === 180 ? [w - X, h - Y] : r === 270 ? [w - Y, X] : [X, Y];
}
const io = new IntersectionObserver((ents) => ents.forEach((en) => { if (en.isIntersecting) paint(en.target); }), { root: view, rootMargin: "600px 0px" });
function draw() {
  view.querySelectorAll(".pg").forEach((x) => x.remove());
  $("empty").hidden = pages.length > 0; $("rail").hidden = !pages.length;
  pages.forEach((p, k) => {
    const [dw, dh] = dispSize(p);
    const el = document.createElement("div"); el.className = "pg"; el.dataset.id = p.id;
    el.style.width = dw * zoom + "px"; el.style.height = dh * zoom + "px";
    el.innerHTML = `<span class="lbl">Page ${k + 1}</span><canvas></canvas><svg viewBox="0 0 ${dw} ${dh}" xmlns="http://www.w3.org/2000/svg"><g transform="${matrix(p)}"></g></svg>`;
    view.appendChild(el); io.observe(el); drawAnn(p, el);
    hook(el, p);
  });
  thumbs(); setTool(tool); info();
}
async function paint(el) {
  const p = pages.find((x) => x.id === +el.dataset.id); if (!p || el.dataset.painted === String(zoom) + ":" + R(p)) return;
  el.dataset.painted = String(zoom) + ":" + R(p);
  const cv = el.querySelector("canvas"), dpr = Math.min(2, window.devicePixelRatio || 1);
  const [dw, dh] = dispSize(p);
  cv.width = Math.round(dw * zoom * dpr); cv.height = Math.round(dh * zoom * dpr);
  const g = cv.getContext("2d");
  if (p.d < 0) { g.fillStyle = "#fff"; g.fillRect(0, 0, cv.width, cv.height); return; }
  try {
    const pg = await docs[p.d].js.getPage(p.i + 1);
    await pg.render({ canvasContext: g, viewport: pg.getViewport({ scale: zoom * dpr, rotation: R(p) }) }).promise;
  } catch (e) { }
}
function drawAnn(p, el) {
  el = el || view.querySelector(`.pg[data-id="${p.id}"]`); if (!el) return;
  const g = el.querySelector("g");
  g.innerHTML = p.ann.map((a, k) => {
    if (a.t === "ink") { let d = ""; for (let i = 0; i < a.p.length; i += 2) d += (i ? "L" : "M") + a.p[i].toFixed(1) + " " + a.p[i + 1].toFixed(1); if (a.p.length === 2) d += "l0.1 0.1";
      return `<path data-k="${k}" d="${d}" fill="none" stroke="${a.c}" stroke-width="${a.w}" stroke-linecap="round" stroke-linejoin="round" opacity="${a.o || 1}"/>`; }
    if (a.t === "rect") return `<rect data-k="${k}" x="${a.x}" y="${a.y}" width="${a.w}" height="${a.h}" ${a.fill ? 'fill="#ffffff" stroke="none"' : `fill="none" stroke="${a.c}" stroke-width="${a.lw || 2}"`}/>`;
    if (a.t === "text") return `<text data-k="${k}" x="${a.x}" y="${a.y + a.s}" font-family="Helvetica, Arial, sans-serif" font-size="${a.s}" fill="${a.c}">` +
      a.v.split("\n").map((l, i) => `<tspan x="${a.x}" dy="${i ? a.s * 1.2 : 0}">${Office.esc(l)}</tspan>`).join("") + "</text>";
    if (a.t === "img") return `<image data-k="${k}" href="${a.src}" x="${a.x}" y="${a.y}" width="${a.w}" height="${a.h}"/>`;
    return "";
  }).join("");
}
function thumbs() {
  const rail = $("rail"); rail.innerHTML = "";
  pages.forEach((p, k) => {
    const b = document.createElement("div"); b.className = "th"; b.dataset.k = k; b.tabIndex = 0; b.setAttribute("role", "button"); b.setAttribute("aria-label", "Page " + (k + 1));
    if (k === cur) b.setAttribute("aria-current", "true");
    b.innerHTML = `<canvas></canvas><span class="n">${k + 1}</span><input type="checkbox" aria-label="Tick page ${k + 1}" ${p.tick ? "checked" : ""}>`;
    b.querySelector("input").onclick = (e) => { e.stopPropagation(); p.tick = e.target.checked; info(); };
    b.onclick = () => go(k);
    rail.appendChild(b);
    const cv = b.querySelector("canvas"), [dw, dh] = dispSize(p), s = 120 / dw;
    cv.width = 120; cv.height = Math.round(dh * s);
    const g = cv.getContext("2d"); g.fillStyle = "#fff"; g.fillRect(0, 0, cv.width, cv.height);
    if (p.d >= 0) docs[p.d].js.getPage(p.i + 1).then((pg) => pg.render({ canvasContext: g, viewport: pg.getViewport({ scale: s, rotation: R(p) }) }).promise).catch(() => {});
  });
}
let jumping = 0;
function go(k) {
  cur = Math.max(0, Math.min(pages.length - 1, k)); jumping = Date.now();
  const el = view.querySelectorAll(".pg")[cur]; if (el) el.scrollIntoView({ block: "start", behavior: "smooth" });
  $("rail").querySelectorAll(".th").forEach((t, i) => t.setAttribute("aria-current", String(i === cur)));
  info();
}
view.addEventListener("scroll", () => {
  if (Date.now() - jumping < 900) return;
  const els = view.querySelectorAll(".pg"), top = view.getBoundingClientRect().top + 80;
  let k = 0; els.forEach((el, i) => { if (el.getBoundingClientRect().top <= top) k = i; });
  if (k !== cur) { cur = k; $("rail").querySelectorAll(".th").forEach((t, i) => t.setAttribute("aria-current", String(i === cur))); info(); }
});
function info() {
  const t = pages.filter((p) => p.tick).length;
  $("pgInfo").textContent = pages.length ? `Page ${cur + 1} of ${pages.length}` + (t ? ` · ${t} ticked` : "") : "";
  $("zoomN").textContent = Math.round(zoom * 100) + "%";
}

/* ---------- marking up ---------- */
function setTool(t) {
  tool = t;
  document.querySelectorAll("[data-tool]").forEach((b) => b.setAttribute("aria-pressed", String(b.dataset.tool === t)));
  view.querySelectorAll(".pg svg").forEach((s) => { s.style.pointerEvents = t === "select" ? "none" : "auto"; s.style.cursor = t === "erase" ? "pointer" : t === "text" ? "text" : "crosshair"; });
}
document.querySelectorAll("[data-tool]").forEach((b) => b.onclick = () => {
  if (b.dataset.tool === "sign" && !signature) { signPad(() => setTool("sign")); return; }
  setTool(b.dataset.tool);
  if (b.dataset.tool === "sign") Office.toast("Tap where the signature goes.");
  if (b.dataset.tool === "text") Office.toast("Tap where the words go.");
});
$("col").addEventListener("input", (e) => { $("colBar").style.background = e.target.value; });
function snap(p) { undo.push({ id: p.id, ann: JSON.stringify(p.ann) }); if (undo.length > 100) undo.shift(); }
$("undoBtn").onclick = () => { const u = undo.pop(); if (!u) return; const p = pages.find((x) => x.id === u.id); if (p) { p.ann = JSON.parse(u.ann); drawAnn(p); } };
function hook(el, p) {
  const svg = el.querySelector("svg");
  let live = null;
  const pt = (e) => { const r = svg.getBoundingClientRect(), [dw, dh] = dispSize(p); return fromDisp(p, (e.clientX - r.left) / r.width * dw, (e.clientY - r.top) / r.height * dh); };
  svg.addEventListener("pointerdown", (e) => {
    if (tool === "select") return;
    e.preventDefault(); cur = pages.indexOf(p); info();
    const [x, y] = pt(e), c = $("col").value, sz = +$("size").value;
    if (tool === "erase") { const t = e.target.closest("[data-k]"); if (t) { snap(p); p.ann.splice(+t.dataset.k, 1); drawAnn(p); } return; }
    if (tool === "text") { textBox((v) => { if (!v) return; snap(p); p.ann.push({ t: "text", x, y, s: sz >= 12 ? sz : 14, c, v }); drawAnn(p); }); return; }
    if (tool === "sign") {
      const w = 170, h = w * signature.h / signature.w; snap(p);
      p.ann.push({ t: "img", src: signature.src, x: x - w / 2, y: y - h / 2, w, h }); drawAnn(p); return;
    }
    svg.setPointerCapture(e.pointerId); snap(p);
    if (tool === "pen" || tool === "hl") {
      live = { t: "ink", c: tool === "hl" ? "#facc15" : c, w: tool === "hl" ? Math.max(10, sz * 4) : (sz >= 12 ? 3 : sz) * (e.pointerType === "pen" && e.pressure ? 0.6 + e.pressure : 1), o: tool === "hl" ? 0.4 : 1, p: [x, y] };
    } else if (tool === "box" || tool === "white") live = { t: "rect", c, x, y, w: 0, h: 0, fill: tool === "white", ox: x, oy: y, lw: sz >= 12 ? 2 : sz };
    if (live) { p.ann.push(live); drawAnn(p); }
  });
  svg.addEventListener("pointermove", (e) => {
    if (!live) { if (tool === "erase" && e.buttons) { const t = document.elementFromPoint(e.clientX, e.clientY); const k = t && t.closest && t.closest("[data-k]"); if (k && svg.contains(k)) { p.ann.splice(+k.dataset.k, 1); drawAnn(p); } } return; }
    const [x, y] = pt(e);
    if (live.t === "ink") { const L = live.p.length; if (Math.hypot(live.p[L - 2] - x, live.p[L - 1] - y) < 0.8 || L > 6000) return; live.p.push(x, y); }
    else { live.x = Math.min(live.ox, x); live.y = Math.min(live.oy, y); live.w = Math.abs(x - live.ox); live.h = Math.abs(y - live.oy); }
    drawAnn(p);
  });
  const end = () => {
    if (live && live.t === "rect") { delete live.ox; delete live.oy; if (live.w < 3 || live.h < 3) { p.ann.pop(); undo.pop(); } }
    live = null; drawAnn(p);
  };
  svg.addEventListener("pointerup", end); svg.addEventListener("pointercancel", end);
}
function textBox(done) {
  const d = document.createElement("dialog"); d.className = "of-dialog";
  d.innerHTML = '<div class="of-dhead"><h2>Type on the PDF</h2><button type="button" class="of-x" aria-label="Close">✕</button></div>' +
    '<textarea id="tbx" rows="3" style="width:100%;border:1px solid var(--rule);border-radius:10px;padding:8px;background:var(--tool);font:inherit;margin:12px 0;box-sizing:border-box"></textarea>' +
    '<div class="of-new"><button type="button" class="of-btn" id="tOk">Put it on</button></div>';
  document.body.appendChild(d); d.showModal(); const ta = d.querySelector("textarea"); ta.focus();
  const close = (v) => { d.close(); d.remove(); done(v); };
  d.querySelector(".of-x").onclick = () => close("");
  d.querySelector("#tOk").onclick = () => close(ta.value.trim());
}
function signPad(after) {
  const d = document.createElement("dialog"); d.className = "of-dialog";
  d.innerHTML = '<div class="of-dhead"><h2>Your signature</h2><button type="button" class="of-x" aria-label="Close">✕</button></div>' +
    '<p class="of-where">Sign with your finger, a pencil or the mouse. It stays on this device.</p><canvas class="sigpad" width="900" height="300"></canvas>' +
    '<div class="of-new" style="margin-top:12px"><button type="button" class="of-btn" id="sOk">Use this signature</button><button type="button" class="of-btn ghost" id="sClr">Clear</button></div>';
  document.body.appendChild(d); d.showModal();
  const cv = d.querySelector("canvas"), g = cv.getContext("2d"); let down = false, any = false;
  g.lineWidth = 6; g.lineCap = "round"; g.lineJoin = "round"; g.strokeStyle = "#111827";
  const at = (e) => { const r = cv.getBoundingClientRect(); return [(e.clientX - r.left) / r.width * cv.width, (e.clientY - r.top) / r.height * cv.height]; };
  cv.onpointerdown = (e) => { down = true; any = true; cv.setPointerCapture(e.pointerId); const [x, y] = at(e); g.beginPath(); g.moveTo(x, y); };
  cv.onpointermove = (e) => { if (!down) return; const [x, y] = at(e); g.lineTo(x, y); g.stroke(); };
  cv.onpointerup = () => { down = false; };
  d.querySelector("#sClr").onclick = () => { g.clearRect(0, 0, cv.width, cv.height); any = false; };
  d.querySelector(".of-x").onclick = () => { d.close(); d.remove(); };
  d.querySelector("#sOk").onclick = () => {
    if (!any) { Office.toast("Sign in the box first."); return; }
    // trim the empty space round the signature
    const px = g.getImageData(0, 0, cv.width, cv.height).data; let x0 = cv.width, y0 = cv.height, x1 = 0, y1 = 0;
    for (let y = 0; y < cv.height; y += 2) for (let x = 0; x < cv.width; x += 2) if (px[(y * cv.width + x) * 4 + 3] > 20) { if (x < x0) x0 = x; if (y < y0) y0 = y; if (x > x1) x1 = x; if (y > y1) y1 = y; }
    const w = Math.max(10, x1 - x0 + 12), h = Math.max(10, y1 - y0 + 12), t = document.createElement("canvas"); t.width = w; t.height = h;
    t.getContext("2d").drawImage(cv, x0 - 6, y0 - 6, w, h, 0, 0, w, h);
    signature = { src: t.toDataURL("image/png"), w, h };
    d.close(); d.remove(); after && after(); Office.toast("Tap where the signature goes.");
  };
}

/* ---------- pages ---------- */
function targets() { const t = pages.filter((p) => p.tick); return t.length ? t : pages[cur] ? [pages[cur]] : []; }
$("rotL").onclick = () => { targets().forEach((p) => { p.rot = (p.rot + 270) % 360; }); draw(); go(cur); };
$("rotR").onclick = () => { targets().forEach((p) => { p.rot = (p.rot + 90) % 360; }); draw(); go(cur); };
$("delBtn").onclick = function () {
  const t = targets(); if (!t.length) return;
  if (t.length >= pages.length) { Office.toast("A PDF needs at least one page."); return; }
  if (!this.dataset.armed) { this.dataset.armed = 1; this.textContent = "Delete " + t.length + "?"; setTimeout(() => { delete this.dataset.armed; this.textContent = "🗑"; }, 2500); return; }
  delete this.dataset.armed; this.textContent = "🗑";
  pages = pages.filter((p) => t.indexOf(p) < 0); cur = Math.min(cur, pages.length - 1); draw(); go(cur);
};
function move(d) { const j = cur + d; if (j < 0 || j >= pages.length) return; const p = pages.splice(cur, 1)[0]; pages.splice(j, 0, p); cur = j; draw(); go(cur); }
$("upBtn").onclick = () => move(-1); $("downBtn").onclick = () => move(1);
$("blankBtn").onclick = () => { const ref = pages[cur] || { w: 595.28, h: 841.89 }; pages.splice(cur + 1, 0, { id: ++uid, d: -1, i: 0, rot: 0, inh: 0, w: ref.w, h: ref.h, ann: [] }); draw(); go(cur + 1); };
$("zIn").onclick = () => { zoom = Math.min(3, +(zoom + 0.25).toFixed(2)); draw(); go(cur); };
$("zOut").onclick = () => { zoom = Math.max(0.5, +(zoom - 0.25).toFixed(2)); draw(); go(cur); };
$("numBtn").onclick = function () { const on = this.getAttribute("aria-pressed") !== "true"; this.setAttribute("aria-pressed", String(on)); Office.toast(on ? "Page numbers will be added when you save." : "No page numbers."); };
document.addEventListener("keydown", (e) => {
  if (/INPUT|TEXTAREA|SELECT/.test(document.activeElement.tagName)) return;
  if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "z") { e.preventDefault(); $("undoBtn").click(); }
  if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "s") { e.preventDefault(); save("pdf"); }
  if (e.key === "PageDown") { e.preventDefault(); go(cur + 1); } if (e.key === "PageUp") { e.preventDefault(); go(cur - 1); }
});

/* ---------- forms ---------- */
$("formBtn").onclick = () => {
  const box = $("formBox"); box.hidden = !box.hidden; if (box.hidden) return;
  let h = '<h3 style="margin:0 0 6px;display:flex;justify-content:space-between">Fill in the form <button type="button" class="of-x" id="fClose" aria-label="Close">✕</button></h3><p class="of-where">What you type goes into the PDF when you save.</p>';
  docs.forEach((d, di) => d.fields.forEach((f, fi) => {
    const key = di + ":" + f.n, v = key in formVals ? formVals[key] : f.v, id = "ff" + di + "_" + fi;
    if (f.t === "check") h += `<label><input type="checkbox" id="${id}" data-key="${Office.esc(key)}" ${v ? "checked" : ""}> ${Office.esc(f.n)}</label>`;
    else if (f.t === "pick") h += `<label for="${id}">${Office.esc(f.n)}</label><select id="${id}" data-key="${Office.esc(key)}" style="width:100%;height:34px">` + f.o.map((o) => `<option ${o === v ? "selected" : ""}>${Office.esc(o)}</option>`).join("") + "</select>";
    else h += `<label for="${id}">${Office.esc(f.n)}</label><input type="text" id="${id}" data-key="${Office.esc(key)}" value="${Office.esc(v)}">`;
  }));
  box.innerHTML = h;
  box.querySelectorAll("[data-key]").forEach((i) => i.addEventListener("input", () => { formVals[i.dataset.key] = i.type === "checkbox" ? i.checked : i.value; }));
  box.querySelectorAll("input[type=checkbox][data-key]").forEach((i) => i.addEventListener("change", () => { formVals[i.dataset.key] = i.checked; }));
  $("fClose").onclick = () => { box.hidden = true; };
};

/* ---------- saving ---------- */
function hexRgb(h) { const n = parseInt(h.slice(1), 16); return rgb((n >> 16 & 255) / 255, (n >> 8 & 255) / 255, (n & 255) / 255); }
function safeText(font, s) { return Array.from(s).map((ch) => { try { font.encodeText(ch); return ch; } catch (e) { return "?"; } }).join(""); }
async function build(list) {
  const out = await PDFDocument.create(), src = {};
  const font = await out.embedFont(StandardFonts.Helvetica);
  for (const p of list) {
    if (p.d >= 0 && !src[p.d]) {
      const lib = await PDFDocument.load(docs[p.d].bytes, { ignoreEncryption: true });
      try { // put the form answers in, then make them part of the page
        const form = lib.getForm();
        docs[p.d].fields.forEach((f) => {
          const key = p.d + ":" + f.n; if (!(key in formVals)) return;
          const v = formVals[key], fld = form.getField(f.n);
          if (f.t === "text") fld.setText(String(v)); else if (f.t === "check") v ? fld.check() : fld.uncheck(); else fld.select(String(v));
        });
        if (form.getFields().length) form.flatten();
      } catch (e) { }
      src[p.d] = lib;
    }
  }
  const imgCache = {};
  let n = 0;
  for (const p of list) {
    n++;
    let page;
    if (p.d < 0) page = out.addPage([p.w, p.h]);
    else {
      const [cp] = await out.copyPages(src[p.d], [p.i]); page = out.addPage(cp);
      // form boxes were turned into part of the page; drop the leftover links to them
      const an = page.node.lookup(PDFName.of("Annots"));
      if (an && an.size) {
        const keep = [];
        for (let k = 0; k < an.size(); k++) {
          const ref = an.get(k), o = out.context.lookup(ref);
          if (o instanceof PDFDict && String(o.get(PDFName.of("Subtype"))) !== "/Widget") keep.push(ref);
        }
        if (keep.length) page.node.set(PDFName.of("Annots"), out.context.obj(keep)); else page.node.delete(PDFName.of("Annots"));
      }
    }
    page.setRotation(degrees(R(p)));
    const box = page.getCropBox(), x0 = box.x, y0 = box.y, H = box.height;
    for (const a of p.ann) {
      if (a.t === "ink" && a.p.length >= 2) {
        let d = ""; for (let i = 0; i < a.p.length; i += 2) d += (i ? "L" : "M") + a.p[i].toFixed(2) + " " + a.p[i + 1].toFixed(2);
        if (a.p.length === 2) d += "L" + (a.p[0] + 0.1).toFixed(2) + " " + (a.p[1] + 0.1).toFixed(2);
        page.drawSvgPath(d, { x: x0, y: y0 + H, borderColor: hexRgb(a.c), borderWidth: a.w, borderOpacity: a.o || 1, borderLineCap: 1 });
      } else if (a.t === "rect") {
        if (a.fill) page.drawRectangle({ x: x0 + a.x, y: y0 + H - a.y - a.h, width: a.w, height: a.h, color: rgb(1, 1, 1) });
        else page.drawRectangle({ x: x0 + a.x, y: y0 + H - a.y - a.h, width: a.w, height: a.h, borderColor: hexRgb(a.c), borderWidth: a.lw || 2 });
      } else if (a.t === "text") {
        a.v.split("\n").forEach((line, i) => page.drawText(safeText(font, line), { x: x0 + a.x, y: y0 + H - a.y - a.s * (0.85 + i * 1.2), size: a.s, font, color: hexRgb(a.c) }));
      } else if (a.t === "img") {
        if (!imgCache[a.src]) imgCache[a.src] = await out.embedPng(Uint8Array.from(atob(a.src.split(",")[1]), (c) => c.charCodeAt(0)));
        page.drawImage(imgCache[a.src], { x: x0 + a.x, y: y0 + H - a.y - a.h, width: a.w, height: a.h });
      }
    }
    if ($("numBtn").getAttribute("aria-pressed") === "true") {
      const t = String(n), w = font.widthOfTextAtSize(t, 10), W = box.width, r = R(p);
      // the bottom-middle of the page as you look at it
      const spots = { 0: [x0 + W / 2 - w / 2, y0 + 18, 0], 90: [x0 + W - 18, y0 + H / 2 - w / 2, 90], 180: [x0 + W / 2 + w / 2, y0 + H - 18, 180], 270: [x0 + 18, y0 + H / 2 + w / 2, 270] }[r];
      page.drawText(t, { x: spots[0], y: spots[1], size: 10, font, color: rgb(0.35, 0.35, 0.4), rotate: degrees(spots[2]) });
    }
  }
  out.setTitle($("title").value || "PDF"); out.setProducer("Launchpad Suite"); out.setCreator("Launchpad Suite PDF");
  return out.save({ useObjectStreams: false });
}
async function save(kind) {
  if (!pages.length) { Office.toast("Open a PDF first."); return; }
  const name = $("title").value.trim() || "PDF";
  try {
    if (kind === "png") return pagePng();
    if (kind === "text" || kind === "write") return words(kind);
    const list = kind === "ticked" ? pages.filter((p) => p.tick) : pages;
    if (!list.length) { Office.toast("Tick the pages you want first (the boxes on the left)."); return; }
    $("status").textContent = "Making your PDF…";
    const bytes = await build(list);
    Office.download(new Blob([bytes], { type: "application/pdf" }), Office.safeName(name + (kind === "ticked" ? " (pages)" : ""), ".pdf"));
    $("status").textContent = "Saved a copy - your PDFs stay on this device";
  } catch (e) { $("status").textContent = "Your PDFs stay on this device"; Office.toast("Couldn't make the PDF: " + (e.message || e), true); }
}
async function pagePng() {
  const p = pages[cur], [dw, dh] = dispSize(p), s = 2, cv = document.createElement("canvas");
  cv.width = Math.round(dw * s); cv.height = Math.round(dh * s);
  const g = cv.getContext("2d"); g.fillStyle = "#fff"; g.fillRect(0, 0, cv.width, cv.height);
  if (p.d >= 0) { const pg = await docs[p.d].js.getPage(p.i + 1); await pg.render({ canvasContext: g, viewport: pg.getViewport({ scale: s, rotation: R(p) }) }).promise; }
  const svg = view.querySelector(`.pg[data-id="${p.id}"] svg`).cloneNode(true);
  svg.setAttribute("width", cv.width); svg.setAttribute("height", cv.height);
  const img = new Image();
  await new Promise((ok) => { img.onload = ok; img.onerror = ok; img.src = "data:image/svg+xml;charset=utf-8," + encodeURIComponent(new XMLSerializer().serializeToString(svg)); });
  try { g.drawImage(img, 0, 0); } catch (e) { }
  cv.toBlob((b) => Office.download(b, Office.safeName(($("title").value || "PDF") + " - page " + (cur + 1), ".png")));
}
async function words(kind) {
  const out = [];
  for (const p of pages) {
    if (p.d < 0) continue;
    const pg = await docs[p.d].js.getPage(p.i + 1), tc = await pg.getTextContent();
    let lines = [], line = "", lastY = null;
    tc.items.forEach((it) => { const y = it.transform && it.transform[5]; if (lastY !== null && Math.abs(y - lastY) > 2) { lines.push(line); line = ""; } line += it.str; if (it.hasEOL) { lines.push(line); line = ""; } lastY = y; });
    if (line) lines.push(line);
    out.push(lines.join("\n").replace(/\n{3,}/g, "\n\n").trim());
  }
  const text = out.filter(Boolean).join("\n\n");
  if (!text.trim()) { Office.toast("This PDF has no words in it that can be copied - it may be a scan.", true); return; }
  const name = $("title").value.trim() || "PDF";
  if (kind === "text") { Office.download(new Blob([text], { type: "text/plain;charset=utf-8" }), Office.safeName(name, ".txt")); return; }
  const st = await Office.init("doc");
  const html = text.split(/\n{2,}/).map((para) => "<p>" + Office.esc(para).replace(/\n/g, " ") + "</p>").join("");
  const f = await Office.save({ title: name, body: { html } });
  Office.toast(st.cloud ? "Saved to your account - opening Write…" : "Opening Write…");
  setTimeout(() => { location.href = "/write/?id=" + encodeURIComponent(f.id); }, 500);
}
$("saveBtn").onclick = () => save("pdf");
const mb = $("moreBtn"), mm = $("moreMenu");
mb.onclick = (e) => { e.stopPropagation(); mm.hidden = !mm.hidden; };
document.addEventListener("click", (e) => { if (!mm.hidden && !mm.contains(e.target)) mm.hidden = true; });
mm.onclick = (e) => { const b = e.target.closest("[data-ex]"); if (!b) return; mm.hidden = true; save(b.dataset.ex); };
window.addEventListener("beforeunload", (e) => { if (pages.some((p) => p.ann.length) || Object.keys(formVals).length) { e.preventDefault(); e.returnValue = ""; } });
draw();

/* the ribbon tabs and the File screen, as in Write */
Office.ribbon();
$("fileTab").onclick = () => {
  const ex = (k) => { const b = document.querySelector('#moreMenu [data-ex="' + k + '"]'); if (b) b.click(); };
  Office.backstage({
    print: () => ex("pdf"),
    pages: {
      home: { label: "Home", title: $("title").value, items: [
        { a: "open", icon: "📂", title: "Open a PDF", sub: "From this device - nothing is uploaded" },
        { a: "merge", icon: "➕", title: "Merge", sub: "Add another PDF's pages at the end" },
        { a: "pics", icon: "🖼", title: "Pictures to PDF", sub: "Turn photos into PDF pages" }] },
      save: { label: "Save a copy", title: "Save a copy", items: [
        { a: "ex-pdf", icon: "📕", title: "Save the whole PDF", sub: "With your marks, in this page order" },
        { a: "ex-ticked", icon: "✂️", title: "Save ticked pages only", sub: "Split out the pages you ticked" },
        { a: "ex-png", icon: "🖼", title: "This page as a picture", sub: ".png" },
        { a: "ex-text", icon: "📝", title: "Copy out the words", sub: ".txt of all the text" },
        { a: "ex-write", icon: "✍️", title: "Open the words in Write", sub: "Edit the text as a document" }] },
      info: { label: "Info", html: () => '<h2>Info</h2><dl class="info"><dt>Name</dt><dd>' + Office.esc($("title").value) + "</dd><dt>Pages</dt><dd>" + pages.length + "</dd><dt>Saved</dt><dd>On this device only - nothing is uploaded</dd></dl>" }
    },
    onPick: (a) => {
      if (a === "open") $("openIn").click(); else if (a === "merge") $("mergeIn").click(); else if (a === "pics") $("imgIn").click();
      else if (a.indexOf("ex-") === 0) ex(a.slice(3));
    }
  });
};
