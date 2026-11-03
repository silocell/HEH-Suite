/* OpenDocument (.odt, .ods) and RTF for the Launchpad Suite - the formats
   LibreOffice, OpenOffice and older Word use. Written from scratch on JSZip. */
(function () {
  var NS = {
    office: "urn:oasis:names:tc:opendocument:xmlns:office:1.0", text: "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
    table: "urn:oasis:names:tc:opendocument:xmlns:table:1.0", style: "urn:oasis:names:tc:opendocument:xmlns:style:1.0",
    fo: "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0", draw: "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0",
    svg: "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0", xlink: "http://www.w3.org/1999/xlink",
    of: "urn:oasis:names:tc:opendocument:xmlns:of:1.2"
  };
  var DECL = Object.keys(NS).map(function (k) { return 'xmlns:' + k + '="' + NS[k] + '"'; }).join(" ");
  var HEAD = '<?xml version="1.0" encoding="UTF-8"?>';
  function x(s) { return String(s).replace(/[&<>"]/g, function (c) { return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]; }).replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g, ""); }
  function e(s) { return String(s).replace(/[&<>"]/g, function (c) { return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]; }); }
  function hex(c) {
    if (!c) return null; var m = String(c).match(/^#([0-9a-f]{6})$/i); if (m) return "#" + m[1].toLowerCase();
    m = String(c).match(/^rgba?\(\s*(\d+)[,\s]+(\d+)[,\s]+(\d+)/i);
    return m ? "#" + [m[1], m[2], m[3]].map(function (n) { return ("0" + (+n).toString(16)).slice(-2); }).join("") : null;
  }
  function pack(mime, content, styles, extra) {
    var z = new JSZip();
    z.file("mimetype", mime, { compression: "STORE" });   // must come first, not squashed
    z.file("content.xml", content);
    z.file("styles.xml", styles);
    z.file("meta.xml", HEAD + '<office:document-meta ' + DECL + ' xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" office:version="1.3"><office:meta><meta:generator>Launchpad Suite</meta:generator></office:meta></office:document-meta>');
    var files = ['<manifest:file-entry manifest:full-path="/" manifest:version="1.3" manifest:media-type="' + mime + '"/>',
      '<manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>', '<manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/>',
      '<manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/>'];
    (extra || []).forEach(function (f) { z.file(f.path, f.data, { base64: true }); files.push('<manifest:file-entry manifest:full-path="' + f.path + '" manifest:media-type="' + f.type + '"/>'); });
    z.file("META-INF/manifest.xml", HEAD + '<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3">' + files.join("") + "</manifest:manifest>");
    return z.generateAsync({ type: "blob", mimeType: mime, compression: "DEFLATE" });
  }

  /* ================= .odt writing ================= */
  function writeOdt(root, title, lay) {
    lay = lay || {};
    var tstyles = {}, pstyles = {}, auto = [], pics = [], n = 0;
    function tStyle(p) {
      var k = [p.b ? "b" : "", p.i ? "i" : "", p.u ? "u" : "", p.s ? "s" : "", p.color || "", p.sz || "", p.bg || "", p.va || ""].join("|");
      if (k === "|||||||") return null;
      if (!tstyles[k]) {
        var nm = "T" + (++n), a = "";
        if (p.b) a += ' fo:font-weight="bold"'; if (p.i) a += ' fo:font-style="italic"';
        if (p.u) a += ' style:text-underline-style="solid" style:text-underline-width="auto" style:text-underline-color="font-color"';
        if (p.s) a += ' style:text-line-through-style="solid"';
        if (p.color) a += ' fo:color="' + p.color + '"'; if (p.bg) a += ' fo:background-color="' + p.bg + '"';
        if (p.sz) a += ' fo:font-size="' + p.sz + 'pt"'; if (p.va) a += ' style:text-position="' + (p.va === "sup" ? "super" : "sub") + ' 58%"';
        auto.push('<style:style style:name="' + nm + '" style:family="text"><style:text-properties' + a + "/></style:style>");
        tstyles[k] = nm;
      }
      return tstyles[k];
    }
    function pStyle(parent, align) {
      if (!align || align === "left" || align === "start") return parent;
      var k = parent + "|" + align;
      if (!pstyles[k]) { var nm = "P" + (++n); auto.push('<style:style style:name="' + nm + '" style:family="paragraph" style:parent-style-name="' + parent + '"><style:paragraph-properties fo:text-align="' + (align === "justify" ? "justify" : align === "center" ? "center" : "end") + '"/></style:style>'); pstyles[k] = nm; }
      return pstyles[k];
    }
    function props(el, base) {
      var p = Object.assign({}, base); if (el.nodeType !== 1) return p;
      var t = el.tagName, st = el.style;
      if (/^(B|STRONG|TH)$/.test(t) || /bold|[6-9]00/.test(st.fontWeight)) p.b = 1;
      if (/^(I|EM)$/.test(t) || st.fontStyle === "italic") p.i = 1;
      if (t === "U" || /underline/.test(st.textDecorationLine || st.textDecoration || "")) p.u = 1;
      if (/^(S|STRIKE|DEL)$/.test(t) || /line-through/.test(st.textDecorationLine || st.textDecoration || "")) p.s = 1;
      if (t === "SUP") p.va = "sup"; if (t === "SUB") p.va = "sub";
      if (hex(st.color)) p.color = hex(st.color); if (hex(st.backgroundColor)) p.bg = hex(st.backgroundColor);
      var m = String(st.fontSize || "").match(/^([\d.]+)(pt|px)$/); if (m) p.sz = Math.round((m[2] === "pt" ? +m[1] : +m[1] * 0.75) * 10) / 10;
      return p;
    }
    function inline(node, p) {
      var o = "";
      for (var c = node.firstChild; c; c = c.nextSibling) {
        if (c.nodeType === 3) { var s = x(c.nodeValue.replace(/\u00a0/g, " ")).replace(/ {2,}/g, function (m) { return ' <text:s text:c="' + (m.length - 1) + '"/>'; }); var ts = tStyle(p); o += ts ? '<text:span text:style-name="' + ts + '">' + s + "</text:span>" : s; continue; }
        if (c.nodeType !== 1) continue;
        if (c.tagName === "BR") { o += "<text:line-break/>"; continue; }
        if (c.tagName === "IMG") {
          var m = (c.getAttribute("src") || "").match(/^data:image\/(png|jpe?g|gif);base64,(.+)$/i); if (!m) continue;
          var ext = /png/i.test(m[1]) ? "png" : /gif/i.test(m[1]) ? "gif" : "jpg", path = "Pictures/img" + (pics.length + 1) + "." + ext;
          pics.push({ path: path, data: m[2], type: "image/" + (ext === "jpg" ? "jpeg" : ext) });
          var w = c.width || c.naturalWidth || 400, h = c.height || c.naturalHeight || 300; if (w > 640) { h = h * 640 / w; w = 640; }
          o += '<draw:frame draw:name="Picture' + pics.length + '" text:anchor-type="as-char" svg:width="' + (w / 96).toFixed(3) + 'in" svg:height="' + (h / 96).toFixed(3) + 'in"><draw:image xlink:href="' + path + '" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/></draw:frame>';
          continue;
        }
        if (c.tagName === "A" && /^https?:|^mailto:/.test(c.getAttribute("href") || "")) { o += '<text:a xlink:type="simple" xlink:href="' + x(c.getAttribute("href")) + '">' + inline(c, props(c, p)) + "</text:a>"; continue; }
        if (/^(UL|OL|TABLE|P|DIV|H1|H2|H3|BLOCKQUOTE|HR|PRE)$/.test(c.tagName)) continue;
        o += inline(c, props(c, p));
      }
      return o;
    }
    var body = [];
    function list(el) {
      var o = "<text:list" + (el.tagName === "OL" ? ' text:style-name="Numbering_20_123"' : ' text:style-name="List_20_1"') + ">";
      Array.prototype.forEach.call(el.children, function (li) {
        if (li.tagName !== "LI") return;
        var pre = el.classList.contains("todo") ? (li.classList.contains("done") ? "☑ " : "☐ ") : "";
        o += "<text:list-item><text:p text:style-name=\"" + pStyle("List_20_Paragraph", li.style.textAlign) + '">' + x(pre) + inline(li, {}) + "</text:p>";
        Array.prototype.forEach.call(li.children, function (c) { if (c.tagName === "UL" || c.tagName === "OL") o += list(c); });
        o += "</text:list-item>";
      });
      return o + "</text:list>";
    }
    function blocks(box, out) {
      var loose = null;
      function flush() { if (loose) { out.push('<text:p text:style-name="Standard">' + inline(loose, {}) + "</text:p>"); loose = null; } }
      Array.prototype.slice.call(box.childNodes).forEach(function (c) {
        var t = c.nodeType === 1 ? c.tagName : "";
        if (t === "P" || t === "DIV") { flush(); out.push('<text:p text:style-name="' + pStyle(c.classList.contains("doc-title") ? "Title" : c.classList.contains("doc-subtitle") ? "Subtitle" : "Standard", c.style.textAlign) + '">' + inline(c, {}) + "</text:p>"); }
        else if (/^H[1-3]$/.test(t)) { flush(); out.push('<text:h text:style-name="' + pStyle("Heading_20_" + t[1], c.style.textAlign) + '" text:outline-level="' + t[1] + '">' + inline(c, {}) + "</text:h>"); }
        else if (t === "BLOCKQUOTE") { flush(); out.push('<text:p text:style-name="Quotations">' + inline(c, {}) + "</text:p>"); }
        else if (t === "PRE") { flush(); c.textContent.split("\n").forEach(function (l) { out.push('<text:p text:style-name="Preformatted_20_Text">' + x(l).replace(/ {2,}/g, function (m) { return ' <text:s text:c="' + (m.length - 1) + '"/>'; }) + "</text:p>"); }); }
        else if (t === "UL" || t === "OL") { flush(); if (c.classList.contains("toc")) Array.prototype.forEach.call(c.children, function (li) { out.push('<text:p text:style-name="Standard">' + inline(li, li.classList.contains("toch") ? { b: 1 } : {}) + "</text:p>"); }); else out.push(list(c)); }
        else if (t === "HR") { flush(); out.push(c.classList.contains("pb") ? '<text:p text:style-name="PageBreak"/>' : '<text:p text:style-name="Rule"/>'); }
        else if (t === "TABLE") {
          flush();
          var rows = Array.prototype.filter.call(c.querySelectorAll("tr"), function (r) { return r.closest("table") === c; }), cols = 0;
          rows.forEach(function (r) { cols = Math.max(cols, r.children.length); });
          var o = '<table:table table:name="Table' + (++n) + '" table:style-name="Tbl"><table:table-column table:number-columns-repeated="' + Math.max(1, cols) + '"/>';
          rows.forEach(function (r) {
            o += "<table:table-row>";
            Array.prototype.forEach.call(r.children, function (td) { var inner = []; blocks(td, inner); o += '<table:table-cell table:style-name="Cell" office:value-type="string">' + (inner.join("") || "<text:p/>") + "</table:table-cell>"; });
            for (var k = r.children.length; k < cols; k++) o += '<table:table-cell table:style-name="Cell" office:value-type="string"><text:p/></table:table-cell>';
            o += "</table:table-row>";
          });
          out.push(o + "</table:table>");
        } else if (c.nodeType === 3 && !c.nodeValue.trim()) { }
        else { if (!loose) loose = document.createElement("p"); loose.appendChild(c.cloneNode(true)); }
      });
      flush();
    }
    blocks(root, body);
    var SZ = { A4: ["21cm", "29.7cm"], Letter: ["8.5in", "11in"], A5: ["14.8cm", "21cm"] }[lay.size] || ["21cm", "29.7cm"], land = lay.orient === "landscape";
    var M = { normal: "2.54cm", narrow: "1.27cm", moderate: "1.9cm", wide: "3.81cm" }[lay.margins] || "2.54cm";
    var content = HEAD + '<office:document-content ' + DECL + ' office:version="1.3"><office:automatic-styles>' +
      '<style:style style:name="Tbl" style:family="table"><style:table-properties style:width="16.5cm" table:align="margins"/></style:style>' +
      '<style:style style:name="Cell" style:family="table-cell"><style:table-cell-properties fo:padding="0.1cm" fo:border="0.5pt solid #9aa3ae"/></style:style>' +
      '<style:style style:name="PageBreak" style:family="paragraph" style:parent-style-name="Standard"><style:paragraph-properties fo:break-after="page"/></style:style>' +
      '<style:style style:name="Rule" style:family="paragraph" style:parent-style-name="Standard"><style:paragraph-properties fo:border-bottom="0.5pt solid #9aa3ae" fo:padding="0.05cm"/></style:style>' +
      auto.join("") + "</office:automatic-styles><office:body><office:text>" + (body.join("") || "<text:p/>") + "</office:text></office:body></office:document-content>";
    var ls = Math.round((+lay.ls || 1.15) * 100) + "%";
    var styles = HEAD + '<office:document-styles ' + DECL + ' office:version="1.3"><office:font-face-decls/><office:styles>' +
      '<style:default-style style:family="paragraph"><style:paragraph-properties fo:line-height="' + ls + '"/><style:text-properties style:font-name="Calibri" fo:font-family="Calibri, Carlito, sans-serif" fo:font-size="12pt" fo:language="en" fo:country="GB"/></style:default-style>' +
      '<style:style style:name="Standard" style:family="paragraph" style:class="text"><style:paragraph-properties fo:margin-bottom="0.2cm"/></style:style>' +
      '<style:style style:name="Title" style:family="paragraph" style:parent-style-name="Standard" style:class="chapter"><style:text-properties fo:font-size="28pt"/></style:style>' +
      '<style:style style:name="Subtitle" style:family="paragraph" style:parent-style-name="Standard" style:class="chapter"><style:text-properties fo:font-size="15pt" fo:color="#5a5a5a"/></style:style>' +
      [["1", "20pt", "#1f3864"], ["2", "15pt", "#2f5496"], ["3", "13pt", "#2f5496"]].map(function (h) {
        return '<style:style style:name="Heading_20_' + h[0] + '" style:display-name="Heading ' + h[0] + '" style:family="paragraph" style:parent-style-name="Standard" style:next-style-name="Standard" style:default-outline-level="' + h[0] + '" style:class="text"><style:paragraph-properties fo:margin-top="0.4cm" fo:margin-bottom="0.15cm" fo:keep-with-next="always"/><style:text-properties fo:font-size="' + h[1] + '" fo:font-weight="bold" fo:color="' + h[2] + '"/></style:style>';
      }).join("") +
      '<style:style style:name="Quotations" style:family="paragraph" style:parent-style-name="Standard"><style:paragraph-properties fo:margin-left="1cm" fo:border-left="2pt solid #b4c6e7" fo:padding-left="0.3cm"/><style:text-properties fo:font-style="italic" fo:color="#404040"/></style:style>' +
      '<style:style style:name="Preformatted_20_Text" style:display-name="Preformatted Text" style:family="paragraph" style:parent-style-name="Standard"><style:paragraph-properties fo:margin-bottom="0cm" fo:background-color="#f3f4f6"/><style:text-properties style:font-name="Courier New" fo:font-family="Consolas, &apos;Courier New&apos;, monospace" fo:font-size="10pt"/></style:style>' +
      '<style:style style:name="List_20_Paragraph" style:display-name="List Paragraph" style:family="paragraph" style:parent-style-name="Standard"><style:paragraph-properties fo:margin-bottom="0.05cm"/></style:style>' +
      '<text:list-style style:name="List_20_1" style:display-name="List 1"><text:list-level-style-bullet text:level="1" text:bullet-char="•"><style:list-level-properties text:list-level-position-and-space-mode="label-alignment"><style:list-level-label-alignment text:label-followed-by="listtab" text:list-tab-stop-position="1cm" fo:text-indent="-0.5cm" fo:margin-left="1cm"/></style:list-level-properties></text:list-level-style-bullet>' +
      '<text:list-level-style-bullet text:level="2" text:bullet-char="◦"><style:list-level-properties text:list-level-position-and-space-mode="label-alignment"><style:list-level-label-alignment text:label-followed-by="listtab" text:list-tab-stop-position="1.6cm" fo:text-indent="-0.5cm" fo:margin-left="1.6cm"/></style:list-level-properties></text:list-level-style-bullet></text:list-style>' +
      '<text:list-style style:name="Numbering_20_123" style:display-name="Numbering 123"><text:list-level-style-number text:level="1" style:num-suffix="." style:num-format="1"><style:list-level-properties text:list-level-position-and-space-mode="label-alignment"><style:list-level-label-alignment text:label-followed-by="listtab" text:list-tab-stop-position="1cm" fo:text-indent="-0.5cm" fo:margin-left="1cm"/></style:list-level-properties></text:list-level-style-number>' +
      '<text:list-level-style-number text:level="2" style:num-suffix="." style:num-format="a"><style:list-level-properties text:list-level-position-and-space-mode="label-alignment"><style:list-level-label-alignment text:label-followed-by="listtab" text:list-tab-stop-position="1.6cm" fo:text-indent="-0.5cm" fo:margin-left="1.6cm"/></style:list-level-properties></text:list-level-style-number></text:list-style>' +
      '</office:styles><office:automatic-styles><style:page-layout style:name="pm1"><style:page-layout-properties fo:page-width="' + (land ? SZ[1] : SZ[0]) + '" fo:page-height="' + (land ? SZ[0] : SZ[1]) + '" style:print-orientation="' + (land ? "landscape" : "portrait") + '" fo:margin-top="' + M + '" fo:margin-bottom="' + M + '" fo:margin-left="' + M + '" fo:margin-right="' + M + '"' +
      (/^#[0-9a-f]{6}$/i.test(lay.pc || "") && lay.pc.toLowerCase() !== "#ffffff" ? ' fo:background-color="' + lay.pc + '"' : "") + '>' + (+lay.cols > 1 ? '<style:columns fo:column-count="' + (+lay.cols) + '" fo:column-gap="0.8cm"/>' : "") + "</style:page-layout-properties></style:page-layout></office:automatic-styles>" +
      '<office:master-styles><style:master-page style:name="Standard" style:page-layout-name="pm1"/></office:master-styles></office:document-styles>';
    return pack("application/vnd.oasis.opendocument.text", content, styles, pics);
  }

  /* ================= .odt reading ================= */
  function attr(el, ns, name) { return el.getAttributeNS(NS[ns], name) || el.getAttribute(ns + ":" + name); }
  function kids(el) { var o = []; for (var c = el.firstChild; c; c = c.nextSibling) o.push(c); return o; }
  function readOdt(file) {
    return JSZip.loadAsync(file).then(function (z) {
      var cf = z.file("content.xml"); if (!cf) throw new Error("this doesn't look like an OpenDocument file");
      return Promise.all([cf.async("string"), z.file("styles.xml") ? z.file("styles.xml").async("string") : ""]).then(function (a) {
        var P = new DOMParser(), doc = P.parseFromString(a[0], "application/xml"), sdoc = a[1] ? P.parseFromString(a[1], "application/xml") : null;
        var st = {};
        function collect(d) {
          if (!d) return;
          Array.prototype.forEach.call(d.getElementsByTagNameNS(NS.style, "style"), function (s) {
            var tp = s.getElementsByTagNameNS(NS.style, "text-properties")[0], pp = s.getElementsByTagNameNS(NS.style, "paragraph-properties")[0], o = { parent: attr(s, "style", "parent-style-name") };
            if (tp) {
              if (attr(tp, "fo", "font-weight") === "bold") o.b = 1; if (attr(tp, "fo", "font-style") === "italic") o.i = 1;
              var u = attr(tp, "style", "text-underline-style"); if (u && u !== "none") o.u = 1;
              var lt = attr(tp, "style", "text-line-through-style"); if (lt && lt !== "none") o.s = 1;
              var c = attr(tp, "fo", "color"); if (/^#[0-9a-f]{6}$/i.test(c || "") && c.toLowerCase() !== "#000000") o.color = c;
              var bg = attr(tp, "fo", "background-color"); if (/^#[0-9a-f]{6}$/i.test(bg || "")) o.bg = bg;
              var fs = attr(tp, "fo", "font-size"); if (/^[\d.]+pt$/.test(fs || "")) o.sz = fs;
              var pos = attr(tp, "style", "text-position");
              if (pos) { if (/^super/.test(pos) || parseFloat(pos) > 0) o.va = "sup"; else if (/^sub/.test(pos) || parseFloat(pos) < 0) o.va = "sub"; }
            }
            if (pp) { var al = attr(pp, "fo", "text-align"); if (al) o.align = { center: "center", end: "right", right: "right", justify: "justify" }[al]; if (attr(pp, "fo", "break-before") === "page") o.pbBefore = 1; if (attr(pp, "fo", "break-after") === "page") o.pbAfter = 1; }
            st[attr(s, "style", "name")] = o;
          });
        }
        collect(sdoc); collect(doc);
        function look(name, k) { var seen = 0; while (name && st[name] && seen++ < 10) { if (st[name][k]) return st[name][k]; name = st[name].parent; } return null; }
        var pics = {}, jobs = [];
        Array.prototype.forEach.call(doc.getElementsByTagNameNS(NS.draw, "image"), function (im) {
          var h = im.getAttributeNS(NS.xlink, "href"), m = /\.(png|jpe?g|gif)$/i.exec(h || ""), f = h && z.file(h);
          if (m && f) jobs.push(f.async("base64").then(function (b) { pics[h] = "data:image/" + (/png/i.test(m[1]) ? "png" : /gif/i.test(m[1]) ? "gif" : "jpeg") + ";base64," + b; }));
        });
        return Promise.all(jobs).then(function () {
          function run(n) {
            var o = "";
            kids(n).forEach(function (c) {
              if (c.nodeType === 3) { o += e(c.nodeValue); return; }
              if (c.nodeType !== 1) return;
              var ln = c.localName;
              if (ln === "s") o += " ".repeat(Math.min(50, +attr(c, "text", "c") || 1));
              else if (ln === "tab") o += "&nbsp;&nbsp;&nbsp;&nbsp;";
              else if (ln === "line-break") o += "<br>";
              else if (ln === "span") {
                var nm = attr(c, "text", "style-name"), t = run(c), css = [];
                if (look(nm, "b")) t = "<b>" + t + "</b>"; if (look(nm, "i")) t = "<i>" + t + "</i>"; if (look(nm, "u")) t = "<u>" + t + "</u>"; if (look(nm, "s")) t = "<s>" + t + "</s>";
                var va = look(nm, "va"); if (va) t = "<" + va + ">" + t + "</" + va + ">";
                if (look(nm, "color")) css.push("color:" + look(nm, "color")); if (look(nm, "bg")) css.push("background-color:" + look(nm, "bg")); if (look(nm, "sz")) css.push("font-size:" + look(nm, "sz"));
                o += css.length ? '<span style="' + css.join(";") + '">' + t + "</span>" : t;
              } else if (ln === "a") { var href = c.getAttributeNS(NS.xlink, "href") || ""; o += /^(https?:|mailto:)/i.test(href) ? '<a href="' + e(href) + '">' + run(c) + "</a>" : run(c); }
              else if (ln === "frame") { var im = c.getElementsByTagNameNS(NS.draw, "image")[0], h = im && im.getAttributeNS(NS.xlink, "href"); if (h && pics[h]) o += '<img src="' + pics[h] + '" alt="">'; }
              else if (ln !== "note" && ln !== "bookmark" && ln !== "annotation") o += run(c);
            });
            return o;
          }
          var out = [];
          function blocks(n) {
            kids(n).forEach(function (c) {
              if (c.nodeType !== 1) return;
              var ln = c.localName, nm = attr(c, "text", "style-name"), al = look(nm, "align"), sty = al ? ' style="text-align:' + al + '"' : "";
              if (look(nm, "pbBefore")) out.push('<hr class="pb">');
              if (ln === "h") { var lv = Math.min(3, +attr(c, "text", "outline-level") || 1); out.push("<h" + lv + sty + ">" + (run(c) || "<br>") + "</h" + lv + ">"); }
              else if (ln === "p") {
                var pn = (nm || "").toLowerCase(), inner = run(c) || "<br>";
                if (/^title/.test(pn) || look(nm, "parent") === "Title") out.push('<p class="doc-title"' + sty + ">" + inner + "</p>");
                else if (/^subtitle/.test(pn)) out.push('<p class="doc-subtitle"' + sty + ">" + inner + "</p>");
                else if (/quotation/.test(pn)) out.push("<blockquote" + sty + ">" + inner + "</blockquote>");
                else if (/preformatted/.test(pn)) out.push("<pre>" + c.textContent.replace(/[&<>]/g, function (k) { return { "&": "&amp;", "<": "&lt;", ">": "&gt;" }[k]; }) + "</pre>");
                else out.push("<p" + sty + ">" + inner + "</p>");
              } else if (ln === "list") { out.push(list(c)); }
              else if (ln === "table") {
                var t = "<table><tbody>";
                Array.prototype.forEach.call(c.getElementsByTagNameNS(NS.table, "table-row"), function (r) {
                  if (r.parentNode !== c && r.parentNode.parentNode !== c) return;
                  t += "<tr>"; Array.prototype.forEach.call(r.children, function (td) { if (td.localName !== "table-cell") return; var save = out; out = []; blocks(td); t += "<td>" + (out.join("") || "<br>") + "</td>"; out = save; }); t += "</tr>";
                });
                out.push(t + "</tbody></table>");
              } else if (ln === "section" || ln === "table-of-content" || ln === "index-body") blocks(c);
              if (look(nm, "pbAfter")) out.push('<hr class="pb">');
            });
          }
          function list(l) {
            var numbered = /number|123|numbering/i.test(attr(l, "text", "style-name") || "");
            var o = numbered ? "<ol>" : "<ul>";
            kids(l).forEach(function (it) {
              if (it.nodeType !== 1 || (it.localName !== "list-item" && it.localName !== "list-header")) return;
              o += "<li>";
              kids(it).forEach(function (c) { if (c.nodeType !== 1) return; if (c.localName === "list") o += list(c); else o += run(c); });
              o += "</li>";
            });
            return o + (numbered ? "</ol>" : "</ul>");
          }
          var body = doc.getElementsByTagNameNS(NS.office, "text")[0]; if (body) blocks(body);
          return out.join("");
        });
      });
    });
  }

  /* ================= .ods ================= */
  function toOf(f) { // =SUM(A1:B2) -> of:=SUM([.A1:.B2])
    var s = Formula.mapRefs(f, function (r) { return r; });
    s = s.replace(/(\$?[A-Z]{1,3}\$?\d+):(\$?[A-Z]{1,3}\$?\d+)/g, "[.$1:.$2]").replace(/(^|[^.\w$\[])(\$?[A-Z]{1,3}\$?\d+)(?![\w(])/g, "$1[.$2]");
    var out = "", q = false; for (var i = 0; i < s.length; i++) { var c = s[i]; if (c === '"') q = !q; out += !q && c === "," ? ";" : c; }
    return "of:=" + out;
  }
  function fromOf(f) {
    var s = String(f).replace(/^(of|oooc|msoxl):/, "");
    if (s[0] !== "=") s = "=" + s;
    s = s.replace(/\[\$?[^\].]*\.(\$?[A-Z]{1,3}\$?\d+)(?::\.(\$?[A-Z]{1,3}\$?\d+))?\]/g, function (m, a, b) { return b ? a + ":" + b : a; });
    var out = "", q = false; for (var i = 0; i < s.length; i++) { var c = s[i]; if (c === '"') q = !q; out += !q && c === ";" ? "," : c; }
    return out;
  }
  function writeOds(sheets, value) {
    var auto = {}, autoX = [], n = 0;
    function cs(c) {
      var k = [c.b, c.i, c.u, c.fg, c.bg, c.al, c.nf].join("|"); if (k === "||||||") return "";
      if (!auto[k]) {
        var nm = "ce" + (++n), t = "", cp = "", pp = "";
        if (c.b) t += ' fo:font-weight="bold"'; if (c.i) t += ' fo:font-style="italic"'; if (c.u) t += ' style:text-underline-style="solid" style:text-underline-width="auto" style:text-underline-color="font-color"';
        if (c.fg) t += ' fo:color="' + c.fg + '"'; if (c.bg) cp += ' fo:background-color="' + c.bg + '"';
        if (c.al) pp = '<style:paragraph-properties fo:text-align="' + { l: "start", c: "center", r: "end" }[c.al] + '"/>';
        var df = c.nf === "pct" ? ' style:data-style-name="Npct"' : c.nf === "gbp" ? ' style:data-style-name="Ngbp"' : c.nf === "n2" ? ' style:data-style-name="Nn2"' : c.nf === "int" ? ' style:data-style-name="Nint"' : "";
        autoX.push('<style:style style:name="' + nm + '" style:family="table-cell" style:parent-style-name="Default"' + df + ">" + (cp ? "<style:table-cell-properties" + cp + "/>" : "") + pp + (t ? "<style:text-properties" + t + "/>" : "") + "</style:style>");
        auto[k] = nm;
      }
      return ' table:style-name="' + auto[k] + '"';
    }
    var tables = sheets.map(function (sh, si) {
      var maxR = -1, maxC = -1;
      Object.keys(sh.cells).forEach(function (k) { var p = Formula.parseRef(k); if (p) { maxR = Math.max(maxR, p.r); maxC = Math.max(maxC, p.c); } });
      var o = '<table:table table:name="' + x(sh.name || "Sheet" + (si + 1)) + '">';
      for (var c = 0; c <= Math.max(0, maxC); c++) {
        var w = (sh.w || {})[Formula.colName(c)]; if (w) { autoX.push('<style:style style:name="co' + si + "_" + c + '" style:family="table-column"><style:table-column-properties style:column-width="' + (w / 96 * 2.54).toFixed(2) + 'cm"/></style:style>'); o += '<table:table-column table:style-name="co' + si + "_" + c + '"/>'; }
        else o += "<table:table-column/>";
      }
      for (var r = 0; r <= maxR; r++) {
        o += "<table:table-row>";
        for (c = 0; c <= maxC; c++) {
          var key = Formula.refName(r, c), cell = sh.cells[key];
          if (!cell) { o += "<table:table-cell/>"; continue; }
          var raw = cell.v == null ? "" : String(cell.v), v = value(si, key), sa = cs(cell), f = raw[0] === "=" && raw.length > 1 ? ' table:formula="' + x(toOf(raw.slice(1))) + '"' : "";
          if (raw === "") { o += "<table:table-cell" + sa + "/>"; continue; }
          if (typeof v === "number") o += "<table:table-cell" + sa + f + ' office:value-type="' + (cell.nf === "pct" ? "percentage" : cell.nf === "gbp" ? "currency" : "float") + '"' + (cell.nf === "gbp" ? ' office:currency="GBP"' : "") + ' office:value="' + v + '"><text:p>' + v + "</text:p></table:table-cell>";
          else if (typeof v === "boolean") o += "<table:table-cell" + sa + f + ' office:value-type="boolean" office:boolean-value="' + v + '"><text:p>' + (v ? "TRUE" : "FALSE") + "</text:p></table:table-cell>";
          else { var t = Formula.isErr(v) ? v.err : v == null ? "" : String(v); o += "<table:table-cell" + sa + f + ' office:value-type="string"><text:p>' + x(t) + "</text:p></table:table-cell>"; }
        }
        o += "</table:table-row>";
      }
      if (maxR < 0) o += "<table:table-row><table:table-cell/></table:table-row>";
      return o + "</table:table>";
    }).join("");
    var numberNS = ' xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0"';
    var dataStyles = '<number:number-style style:name="Nn2"><number:number number:decimal-places="2" number:min-decimal-places="2" number:min-integer-digits="1" number:grouping="true"/></number:number-style>' +
      '<number:number-style style:name="Nint"><number:number number:decimal-places="0" number:min-integer-digits="1" number:grouping="true"/></number:number-style>' +
      '<number:percentage-style style:name="Npct"><number:number number:decimal-places="0" number:min-integer-digits="1"/><number:text>%</number:text></number:percentage-style>' +
      '<number:currency-style style:name="Ngbp"><number:currency-symbol number:language="en" number:country="GB">£</number:currency-symbol><number:number number:decimal-places="2" number:min-decimal-places="2" number:min-integer-digits="1" number:grouping="true"/></number:currency-style>';
    var content = HEAD + '<office:document-content ' + DECL + numberNS + ' office:version="1.3"><office:automatic-styles>' + dataStyles + autoX.join("") + "</office:automatic-styles><office:body><office:spreadsheet>" + tables + "</office:spreadsheet></office:body></office:document-content>";
    var styles = HEAD + '<office:document-styles ' + DECL + numberNS + ' office:version="1.3"><office:styles><style:style style:name="Default" style:family="table-cell"><style:text-properties fo:font-size="11pt"/></style:style></office:styles></office:document-styles>';
    return pack("application/vnd.oasis.opendocument.spreadsheet", content, styles);
  }
  function readOds(file) {
    return JSZip.loadAsync(file).then(function (z) {
      var cf = z.file("content.xml"); if (!cf) throw new Error("this doesn't look like an OpenDocument file");
      return cf.async("string").then(function (s) {
        var doc = new DOMParser().parseFromString(s, "application/xml"), bold = {};
        Array.prototype.forEach.call(doc.getElementsByTagNameNS(NS.style, "style"), function (st) {
          var tp = st.getElementsByTagNameNS(NS.style, "text-properties")[0], cp = st.getElementsByTagNameNS(NS.style, "table-cell-properties")[0], o = {};
          if (tp && attr(tp, "fo", "font-weight") === "bold") o.b = 1; if (tp && attr(tp, "fo", "font-style") === "italic") o.i = 1;
          var c = tp && attr(tp, "fo", "color"); if (/^#[0-9a-f]{6}$/i.test(c || "") && c !== "#000000") o.fg = c;
          var bg = cp && attr(cp, "fo", "background-color"); if (/^#[0-9a-f]{6}$/i.test(bg || "")) o.bg = bg;
          bold[attr(st, "style", "name")] = o;
        });
        var out = [];
        Array.prototype.forEach.call(doc.getElementsByTagNameNS(NS.table, "table"), function (t) {
          if (out.length >= 20) return;
          var cells = {}, r = 0, maxR = 0, maxC = 0;
          Array.prototype.forEach.call(t.getElementsByTagNameNS(NS.table, "table-row"), function (row) {
            var rep = Math.min(+attr(row, "table", "number-rows-repeated") || 1, 5000 - r); if (rep <= 0) return;
            var c = 0, made = [];
            Array.prototype.forEach.call(row.children, function (cell) {
              if (cell.localName !== "table-cell" && cell.localName !== "covered-table-cell") return;
              var crep = Math.min(+attr(cell, "table", "number-columns-repeated") || 1, 52 - c);
              var type = attr(cell, "office", "value-type"), f = attr(cell, "table", "formula"), v = "";
              if (f) v = fromOf(f);
              else if (type === "float" || type === "percentage" || type === "currency") v = attr(cell, "office", "value") || "";
              else if (type === "boolean") v = attr(cell, "office", "boolean-value") === "true" ? "TRUE" : "FALSE";
              else if (type === "date") v = (attr(cell, "office", "date-value") || "").slice(0, 10);
              else v = Array.prototype.map.call(cell.getElementsByTagNameNS(NS.text, "p"), function (p) { return p.textContent; }).join("\n");
              var stl = bold[attr(cell, "table", "style-name")] || {};
              for (var k = 0; k < crep && c < 52; k++, c++) {
                if (v === "" && !stl.b && !stl.bg) continue;
                var o = {}; if (v !== "") o.v = v; if (stl.b) o.b = 1; if (stl.i) o.i = 1; if (stl.fg) o.fg = stl.fg; if (stl.bg) o.bg = stl.bg;
                if (type === "percentage") o.nf = "pct"; if (type === "currency") o.nf = "gbp";
                made.push([c, o]);
              }
            });
            for (var k2 = 0; k2 < rep && r < 5000; k2++, r++) made.forEach(function (m) { cells[Formula.refName(r, m[0])] = JSON.parse(JSON.stringify(m[1])); maxR = Math.max(maxR, r); maxC = Math.max(maxC, m[0]); });
          });
          out.push({ name: (attr(t, "table", "name") || "Sheet" + (out.length + 1)).slice(0, 31), cells: cells, w: {}, rows: Math.max(100, maxR + 21), cols: Math.max(26, maxC + 1) });
        });
        if (!out.length) throw new Error("there were no sheets in it");
        return out;
      });
    });
  }

  /* ================= RTF ================= */
  var CP1252 = { 128: 8364, 130: 8218, 131: 402, 132: 8222, 133: 8230, 134: 8224, 135: 8225, 136: 710, 137: 8240, 138: 352, 139: 8249, 140: 338, 142: 381, 145: 8216, 146: 8217, 147: 8220, 148: 8221, 149: 8226, 150: 8211, 151: 8212, 152: 732, 153: 8482, 154: 353, 155: 8250, 156: 339, 158: 382, 159: 376 };
  function readRtf(src) {
    if (!/^\s*\{\\rtf/.test(src)) throw new Error("this doesn't look like an RTF file");
    var i = 0, stack = [], st = { b: 0, i: 0, u: 0, skip: 0, uc: 1 }, paras = [], cur = [], skipN = 0;
    var SKIP = /^(fonttbl|colortbl|stylesheet|info|pict|object|header|footer|headerl|headerr|footerl|footerr|footnote|field|fldinst|themedata|colorschememapping|latentstyles|datastore|xmlnstbl|listtable|listoverridetable|rsidtbl|generator|pgdsctbl|mmathPr|filetbl|revtbl)$/;
    function emit(t) { if (st.skip || !t) return; if (skipN > 0) { var k = Math.min(skipN, t.length); t = t.slice(k); skipN -= k; if (!t) return; } cur.push({ t: t, b: st.b, i: st.i, u: st.u }); }
    function par() { paras.push(cur); cur = []; }
    while (i < src.length) {
      var ch = src[i];
      if (ch === "{") { stack.push(Object.assign({}, st)); i++; if (src.substr(i, 2) === "\\*") st.skip = 1; continue; }
      if (ch === "}") { st = stack.pop() || st; i++; continue; }
      if (ch === "\\") {
        var m = /^\\([a-z]{1,32})(-?\d{1,10})? ?/i.exec(src.slice(i, i + 50));
        if (m) {
          i += m[0].length; var w = m[1], n = m[2] === undefined ? null : +m[2];
          if (SKIP.test(w)) { st.skip = 1; continue; }
          if (w === "par" || w === "sect") par(); else if (w === "line") emit("\n"); else if (w === "tab") emit("\t");
          else if (w === "b") st.b = n === 0 ? 0 : 1; else if (w === "i") st.i = n === 0 ? 0 : 1;
          else if (w === "ul") st.u = n === 0 ? 0 : 1; else if (w === "ulnone") st.u = 0; else if (w === "plain") { st.b = st.i = st.u = 0; }
          else if (w === "uc") st.uc = n || 1;
          else if (w === "u") { var cp = n < 0 ? n + 65536 : n; emit(String.fromCharCode(cp)); skipN = st.uc; }
          else if (w === "emdash") emit("—"); else if (w === "endash") emit("–"); else if (w === "bullet") emit("•");
          else if (w === "lquote") emit("‘"); else if (w === "rquote") emit("’"); else if (w === "ldblquote") emit("“"); else if (w === "rdblquote") emit("”");
          continue;
        }
        var nx = src[i + 1];
        if (nx === "'") { var c = parseInt(src.substr(i + 2, 2), 16); emit(String.fromCharCode(CP1252[c] || c)); i += 4; continue; }
        if (nx === "~") emit("\u00a0"); else if (nx === "-" || nx === "_") emit(nx === "_" ? "-" : ""); else if (nx === "\n" || nx === "\r") par(); else emit(nx);
        i += 2; continue;
      }
      if (ch === "\r" || ch === "\n") { i++; continue; }
      var j = i; while (j < src.length && !/[\\{}\r\n]/.test(src[j])) j++;
      emit(src.slice(i, j)); i = j;
    }
    if (cur.length) par();
    return paras.map(function (p) {
      var h = p.map(function (r) { var t = e(r.t).replace(/\n/g, "<br>").replace(/\t/g, "&nbsp;&nbsp;&nbsp;&nbsp;"); if (r.b) t = "<b>" + t + "</b>"; if (r.i) t = "<i>" + t + "</i>"; if (r.u) t = "<u>" + t + "</u>"; return t; }).join("");
      return "<p>" + (h || "<br>") + "</p>";
    }).join("");
  }
  function rtfText(s) {
    return Array.from(s).map(function (c) {
      var n = c.codePointAt(0);
      if (c === "\\" || c === "{" || c === "}") return "\\" + c;
      if (c === "\t") return "\\tab ";
      if (n < 128) return c;
      if (n > 0xffff) { var hi = Math.floor((n - 0x10000) / 0x400) + 0xd800, lo = (n - 0x10000) % 0x400 + 0xdc00; return "\\u" + (hi - 65536) + "?\\u" + (lo - 65536) + "?"; }
      return "\\u" + (n > 32767 ? n - 65536 : n) + "?";
    }).join("");
  }
  function writeRtf(root) {
    var o = ["{\\rtf1\\ansi\\ansicpg1252\\deff0{\\fonttbl{\\f0\\fswiss Calibri;}{\\f1\\fmodern Courier New;}}{\\colortbl;\\red31\\green56\\blue100;}\\paperw11906\\paperh16838\\margl1440\\margr1440\\margt1440\\margb1440\\f0\\fs24\n"];
    function inline(n, p) {
      var s = "";
      n.childNodes.forEach(function (c) {
        if (c.nodeType === 3) { s += rtfText(c.nodeValue.replace(/\u00a0/g, " ")); return; }
        if (c.nodeType !== 1) return;
        var t = c.tagName, st = c.style || {};
        if (t === "BR") { s += "\\line "; return; }
        if (/^(UL|OL|TABLE|P|DIV|H1|H2|H3|BLOCKQUOTE|HR|PRE|IMG)$/.test(t)) return;
        var on = "", off = "";
        if (/^(B|STRONG|TH)$/.test(t) || /bold|[6-9]00/.test(st.fontWeight || "")) { on += "\\b "; off = "\\b0 " + off; }
        if (/^(I|EM)$/.test(t) || st.fontStyle === "italic") { on += "\\i "; off = "\\i0 " + off; }
        if (t === "U") { on += "\\ul "; off = "\\ulnone " + off; }
        if (/^(S|STRIKE|DEL)$/.test(t)) { on += "\\strike "; off = "\\strike0 " + off; }
        if (t === "SUP") { on += "\\super "; off = "\\nosupersub " + off; } if (t === "SUB") { on += "\\sub "; off = "\\nosupersub " + off; }
        s += "{" + on + inline(c) + off + "}";
      });
      return s;
    }
    function al(el) { var a = el.style && el.style.textAlign; return a === "center" ? "\\qc" : a === "right" ? "\\qr" : a === "justify" ? "\\qj" : "\\ql"; }
    function list(el, depth) {
      var k = 1;
      Array.prototype.forEach.call(el.children, function (li) {
        if (li.tagName !== "LI") return;
        var mark = el.classList.contains("todo") ? (li.classList.contains("done") ? "\\u9745?" : "\\u9744?") : el.tagName === "OL" ? (k++) + "." : "\\bullet";
        o.push("{\\pard\\li" + (720 + depth * 360) + "\\fi-360\\sa60 " + mark + "\\tab " + inline(li) + "\\par}\n");
        Array.prototype.forEach.call(li.children, function (c) { if (c.tagName === "UL" || c.tagName === "OL") list(c, depth + 1); });
      });
    }
    function blocks(box) {
      Array.prototype.forEach.call(box.childNodes, function (c) {
        if (c.nodeType === 3) { if (c.nodeValue.trim()) o.push("{\\pard\\sa120 " + rtfText(c.nodeValue) + "\\par}\n"); return; }
        if (c.nodeType !== 1) return;
        var t = c.tagName;
        if (/^H[1-3]$/.test(t)) o.push("{\\pard" + al(c) + "\\sb240\\sa120\\keepn\\b\\cf1\\fs" + { 1: 40, 2: 30, 3: 26 }[t[1]] + " " + inline(c) + "\\par}\n");
        else if (t === "P" || t === "DIV") { var big = c.classList.contains("doc-title") ? "\\fs56 " : c.classList.contains("doc-subtitle") ? "\\fs30 " : ""; o.push("{\\pard" + al(c) + "\\sa120 " + big + inline(c) + "\\par}\n"); }
        else if (t === "BLOCKQUOTE") o.push("{\\pard\\li720\\sa120\\i " + inline(c) + "\\par}\n");
        else if (t === "PRE") c.textContent.split("\n").forEach(function (l) { o.push("{\\pard\\f1\\fs20 " + rtfText(l) + "\\par}\n"); });
        else if (t === "UL" || t === "OL") { if (c.classList.contains("toc")) Array.prototype.forEach.call(c.children, function (li) { o.push("{\\pard\\sa60 " + inline(li) + "\\par}\n"); }); else list(c, 0); }
        else if (t === "HR") o.push(c.classList.contains("pb") ? "\\page\n" : "{\\pard\\brdrb\\brdrs\\brdrw10\\brsp20 \\par}\n");
        else if (t === "TABLE") {
          Array.prototype.forEach.call(c.querySelectorAll("tr"), function (tr) {
            var cells = Array.prototype.slice.call(tr.children), w = Math.floor(9026 / Math.max(1, cells.length)), row = "\\trowd\\trgaph108";
            cells.forEach(function (td, j) { row += "\\clbrdrt\\brdrs\\clbrdrl\\brdrs\\clbrdrb\\brdrs\\clbrdrr\\brdrs\\cellx" + w * (j + 1); });
            o.push(row + " " + cells.map(function (td) { return "\\pard\\intbl " + (td.tagName === "TH" ? "\\b " : "") + rtfText(td.textContent) + (td.tagName === "TH" ? "\\b0" : "") + "\\cell "; }).join("") + "\\row\n");
          });
        } else o.push("{\\pard\\sa120 " + inline(c) + "\\par}\n");
      });
    }
    blocks(root);
    o.push("}");
    return o.join("");
  }

  window.Odf = { writeOdt: writeOdt, readOdt: readOdt, writeOds: writeOds, readOds: readOds, readRtf: readRtf, writeRtf: writeRtf };
})();
