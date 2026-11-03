/* Word files (.docx) for Launchpad Write, written from scratch.
   A .docx is a zip of XML files; this makes and reads the parts that
   matter for homework: headings, styles, bold/italic/underline, colour,
   highlight, sizes, fonts, lists, links, tables, pictures and lines. */
(function () {
  var W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
  var R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

  function x(s) { return String(s).replace(/[&<>"]/g, function (c) { return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]; })
    .replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g, ""); }
  function hex(c) {
    if (!c) return null; c = String(c).trim();
    var m = c.match(/^#([0-9a-f]{3}|[0-9a-f]{6})$/i);
    if (m) { var h = m[1]; if (h.length === 3) h = h.replace(/./g, "$&$&"); return h.toUpperCase(); }
    m = c.match(/^rgba?\(\s*(\d+)[,\s]+(\d+)[,\s]+(\d+)(?:[,\s\/]+([\d.]+))?/i);
    if (m) { if (m[4] !== undefined && +m[4] === 0) return null;
      return [m[1], m[2], m[3]].map(function (n) { return ("0" + (+n).toString(16)).slice(-2); }).join("").toUpperCase(); }
    var named = { red: "FF0000", blue: "0000FF", green: "008000", black: "000000", white: "FFFFFF", yellow: "FFFF00", orange: "FFA500", purple: "800080", grey: "808080", gray: "808080" };
    return named[c.toLowerCase()] || null;
  }
  function pt(v) {
    var m = String(v || "").match(/^([\d.]+)(pt|px)$/); if (!m) return null;
    return m[2] === "pt" ? +m[1] : +m[1] * 0.75;
  }

  /* ================= writing ================= */
  function build(root, title) {
    var rels = [], media = [], nums = [], body = [];
    function rel(type, target, ext) {
      var id = "rId" + (rels.length + 10);
      rels.push('<Relationship Id="' + id + '" Type="' + R + "/" + type + '" Target="' + x(target) + '"' + (ext ? ' TargetMode="External"' : "") + "/>");
      return id;
    }
    var imgJobs = [];

    function runProps(el, base) {
      var p = Object.assign({}, base);
      if (el.nodeType !== 1) return p;
      var t = el.tagName, st = el.style;
      if (t === "B" || t === "STRONG" || t === "TH") p.b = true;
      if (t === "I" || t === "EM") p.i = true;
      if (t === "U") p.u = true;
      if (t === "S" || t === "STRIKE" || t === "DEL") p.s = true;
      if (t === "SUP") p.va = "superscript";
      if (t === "SUB") p.va = "subscript";
      if (st) {
        if (/bold|[6-9]00/.test(st.fontWeight)) p.b = true;
        if (st.fontWeight === "normal" || st.fontWeight === "400") p.b = false;
        if (st.fontStyle === "italic") p.i = true;
        var td = st.textDecorationLine || st.textDecoration || "";
        if (/underline/.test(td)) p.u = true;
        if (/line-through/.test(td)) p.s = true;
        if (st.color && hex(st.color)) p.color = hex(st.color);
        if (st.backgroundColor && hex(st.backgroundColor)) p.shd = hex(st.backgroundColor);
        if (st.fontSize && pt(st.fontSize)) p.sz = Math.round(pt(st.fontSize) * 2);
        if (st.fontFamily) p.font = st.fontFamily.split(",")[0].replace(/["']/g, "").trim();
        if (st.verticalAlign === "super") p.va = "superscript";
        if (st.verticalAlign === "sub") p.va = "subscript";
      }
      return p;
    }
    function rPr(p) {
      var o = "";
      if (p.font) o += '<w:rFonts w:ascii="' + x(p.font) + '" w:hAnsi="' + x(p.font) + '" w:cs="' + x(p.font) + '"/>';
      if (p.b) o += "<w:b/>"; if (p.i) o += "<w:i/>"; if (p.s) o += "<w:strike/>";
      if (p.color) o += '<w:color w:val="' + p.color + '"/>';
      if (p.sz) o += '<w:sz w:val="' + p.sz + '"/><w:szCs w:val="' + p.sz + '"/>';
      if (p.u) o += '<w:u w:val="single"/>';
      if (p.shd) o += '<w:shd w:val="clear" w:color="auto" w:fill="' + p.shd + '"/>';
      if (p.va) o += '<w:vertAlign w:val="' + p.va + '"/>';
      if (p.link) o += '<w:rStyle w:val="Hyperlink"/>';
      return o ? "<w:rPr>" + o + "</w:rPr>" : "";
    }
    function textRun(t, p) {
      return "<w:r>" + rPr(p) + '<w:t xml:space="preserve">' + x(t.replace(/ /g, " ")) + "</w:t></w:r>";
    }
    function imageRun(img) {
      var src = img.getAttribute("src") || "", m = src.match(/^data:image\/(png|jpe?g|gif);base64,(.+)$/i);
      if (!m) return "";
      var ext = m[1].toLowerCase() === "png" ? "png" : m[1].toLowerCase() === "gif" ? "gif" : "jpeg";
      var n = media.length + 1, name = "image" + n + "." + ext;
      media.push({ name: name, data: m[2] });
      var id = rel("image", "media/" + name);
      var w = img.naturalWidth || img.width || 400, h = img.naturalHeight || img.height || 300;
      var shown = img.width || w; if (shown && w) { h = h * shown / w; w = shown; }
      var max = 602; if (w > max) { h = h * max / w; w = max; }  // 6.27 in of A4 text width at 96 dpi
      var cx = Math.round(w * 9525), cy = Math.round(h * 9525);
      return '<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="' + cx + '" cy="' + cy + '"/>' +
        '<wp:docPr id="' + n + '" name="Picture ' + n + '" descr="' + x(img.alt || "") + '"/>' +
        '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
        '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="' + n + '" name="' + name + '"/><pic:cNvPicPr/></pic:nvPicPr>' +
        '<pic:blipFill><a:blip r:embed="' + id + '"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>' +
        '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="' + cx + '" cy="' + cy + '"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>' +
        "</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>";
    }
    function inline(node, p) {
      var o = "";
      for (var n = node.firstChild; n; n = n.nextSibling) {
        if (n.nodeType === 3) { if (n.nodeValue) o += textRun(n.nodeValue, p); continue; }
        if (n.nodeType !== 1) continue;
        var t = n.tagName;
        if (t === "BR") { o += "<w:r><w:br/></w:r>"; continue; }
        if (t === "IMG") { o += imageRun(n); continue; }
        if (t === "INPUT") { o += textRun(n.checked || n.hasAttribute("checked") ? "☑ " : "☐ ", p); continue; }
        if (t === "A" && /^(https?:|mailto:)/i.test(n.getAttribute("href") || "")) {
          var id = rel("hyperlink", n.getAttribute("href"), true);
          o += '<w:hyperlink r:id="' + id + '">' + inline(n, Object.assign(runProps(n, p), { link: true })) + "</w:hyperlink>";
          continue;
        }
        if (/^(UL|OL|TABLE|P|DIV|H1|H2|H3|BLOCKQUOTE|HR)$/.test(t)) continue; // blocks inside a line are handled by the block walker
        o += inline(n, runProps(n, p));
      }
      return o;
    }
    function para(el, style, extra, base) {
      var pp = "";
      if (style) pp += '<w:pStyle w:val="' + style + '"/>';
      if (extra && extra.num) pp += '<w:numPr><w:ilvl w:val="' + extra.lvl + '"/><w:numId w:val="' + extra.num + '"/></w:numPr>';
      if (extra && extra.ind) pp += '<w:ind w:left="' + extra.ind + '"/>';
      var al = el && el.style && el.style.textAlign;
      var jc = { center: "center", right: "right", justify: "both", left: "left" }[al];
      if (jc) pp += '<w:jc w:val="' + jc + '"/>';
      return "<w:p>" + (pp ? "<w:pPr>" + pp + "</w:pPr>" : "") + (el ? inline(el, base || {}) : "") + "</w:p>";
    }
    function list(el, lvl, out) {
      var abs = el.tagName === "OL" ? 1 : 0;
      nums.push(abs); var num = nums.length;
      Array.prototype.forEach.call(el.children, function (li) {
        if (li.tagName !== "LI") return;
        out.push(para(li, "ListParagraph", { num: num, lvl: Math.min(lvl, 8) }));
        Array.prototype.forEach.call(li.children, function (c) { if (c.tagName === "UL" || c.tagName === "OL") list(c, lvl + 1, out); });
      });
    }
    function table(el, out) {
      var rows = Array.prototype.filter.call(el.querySelectorAll("tr"), function (tr) { return tr.closest("table") === el; });
      var cols = 0;
      rows.forEach(function (tr) { var c = 0; Array.prototype.forEach.call(tr.children, function (td) { c += +td.getAttribute("colspan") || 1; }); cols = Math.max(cols, c); });
      if (!cols) return;
      var tw = 9026, cw = Math.floor(tw / cols), o = '<w:tbl><w:tblPr><w:tblW w:w="' + tw + '" w:type="dxa"/>' +
        '<w:tblBorders><w:top w:val="single" w:sz="4" w:color="9AA3AE"/><w:left w:val="single" w:sz="4" w:color="9AA3AE"/><w:bottom w:val="single" w:sz="4" w:color="9AA3AE"/><w:right w:val="single" w:sz="4" w:color="9AA3AE"/><w:insideH w:val="single" w:sz="4" w:color="9AA3AE"/><w:insideV w:val="single" w:sz="4" w:color="9AA3AE"/></w:tblBorders>' +
        '<w:tblCellMar><w:left w:w="100" w:type="dxa"/><w:right w:w="100" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid>';
      for (var i = 0; i < cols; i++) o += '<w:gridCol w:w="' + cw + '"/>';
      o += "</w:tblGrid>";
      rows.forEach(function (tr) {
        o += "<w:tr>";
        Array.prototype.forEach.call(tr.children, function (td) {
          var span = +td.getAttribute("colspan") || 1, inner = [];
          blocks(td, inner, td.tagName === "TH" ? { b: true } : {});
          if (!inner.length) inner.push("<w:p/>");
          o += '<w:tc><w:tcPr><w:tcW w:w="' + cw * span + '" w:type="dxa"/>' + (span > 1 ? '<w:gridSpan w:val="' + span + '"/>' : "") +
            (td.tagName === "TH" ? '<w:shd w:val="clear" w:color="auto" w:fill="EEF2F8"/>' : "") + "</w:tcPr>" + inner.join("") + "</w:tc>";
        });
        o += "</w:tr>";
      });
      out.push(o + "</w:tbl><w:p/>");
    }
    /* Walk the blocks of a container. Loose text and inline tags between
       blocks are gathered into one paragraph, like a browser shows them. */
    function blocks(box, out, base) {
      var loose = document.createElement("p"), any = false;
      function flush() {
        if (any) out.push(para(loose, null, null, base));
        loose = document.createElement("p"); any = false;
      }
      Array.prototype.slice.call(box.childNodes).forEach(function (n) {
        if (n.nodeType === 1) {
          var t = n.tagName;
          if (t === "P" || t === "DIV") {
            flush();
            if (n.querySelector("ul,ol,table,p,div,h1,h2,h3,blockquote")) { blocks(n, out, base); return; }
            out.push(para(n, n.classList.contains("doc-title") ? "Title" : null, null, base)); return;
          }
          if (/^H[1-3]$/.test(t)) { flush(); out.push(para(n, "Heading" + t[1], null, base)); return; }
          if (t === "BLOCKQUOTE") {
            flush();
            var indent = /margin/.test(n.getAttribute("style") || "");
            if (n.querySelector("p,div,ul,ol,h1,h2,h3")) {
              var tmp = []; blocks(n, tmp, base);
              tmp.forEach(function (b) { out.push(b); });
            } else out.push(para(n, indent ? null : "Quote", indent ? { ind: 720 } : null, base));
            return;
          }
          if (t === "UL" || t === "OL") { flush(); list(n, 0, out); return; }
          if (t === "TABLE") { flush(); table(n, out); return; }
          if (t === "HR") { flush(); out.push('<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="9AA3AE"/></w:pBdr></w:pPr></w:p>'); return; }
        }
        if (n.nodeType === 3 && !n.nodeValue.trim() && !any) return;
        loose.appendChild(n.cloneNode(true)); any = true;
      });
      flush();
    }

    // images need their real size, so wait for them to be measured
    Array.prototype.forEach.call(root.querySelectorAll("img"), function (im) {
      if (!im.complete) imgJobs.push(new Promise(function (ok) { im.onload = im.onerror = ok; }));
    });
    return Promise.all(imgJobs).then(function () {
      blocks(root, body, {});
      var numbering = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering xmlns:w="' + W + '">';
      [0, 1].forEach(function (abs) {
        numbering += '<w:abstractNum w:abstractNumId="' + abs + '"><w:multiLevelType w:val="hybridMultilevel"/>';
        for (var l = 0; l < 9; l++) {
          var bullet = ["•", "◦", "▪"][l % 3], fmt = abs ? ["decimal", "lowerLetter", "lowerRoman"][l % 3] : "bullet";
          numbering += '<w:lvl w:ilvl="' + l + '"><w:start w:val="1"/><w:numFmt w:val="' + fmt + '"/><w:lvlText w:val="' + (abs ? "%" + (l + 1) + "." : bullet) + '"/><w:lvlJc w:val="left"/>' +
            '<w:pPr><w:ind w:left="' + (720 + 360 * l) + '" w:hanging="360"/></w:pPr>' + (abs ? "" : '<w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/></w:rPr>') + "</w:lvl>";
        }
        numbering += "</w:abstractNum>";
      });
      nums.forEach(function (abs, i) {
        numbering += '<w:num w:numId="' + (i + 1) + '"><w:abstractNumId w:val="' + abs + '"/>' +
          (abs ? '<w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride>' : "") + "</w:num>";
      });
      numbering += "</w:numbering>";

      var doc = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<w:document xmlns:w="' + W + '" xmlns:r="' + R + '" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">' +
        "<w:body>" + (body.join("") || "<w:p/>") +
        '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>' +
        "</w:body></w:document>";

      var styles = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="' + W + '">' +
        '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/><w:sz w:val="24"/><w:szCs w:val="24"/><w:lang w:val="en-GB"/></w:rPr></w:rPrDefault>' +
        '<w:pPrDefault><w:pPr><w:spacing w:after="120" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>' +
        '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>' +
        '<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:spacing w:after="120"/></w:pPr><w:rPr><w:rFonts w:ascii="Calibri Light" w:hAnsi="Calibri Light"/><w:sz w:val="56"/></w:rPr></w:style>' +
        '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="240" w:after="80"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:color w:val="1F3864"/><w:sz w:val="40"/></w:rPr></w:style>' +
        '<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="200" w:after="80"/><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:color w:val="2F5496"/><w:sz w:val="30"/></w:rPr></w:style>' +
        '<w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="heading 3"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="160" w:after="60"/><w:outlineLvl w:val="2"/></w:pPr><w:rPr><w:color w:val="2F5496"/><w:sz w:val="26"/></w:rPr></w:style>' +
        '<w:style w:type="paragraph" w:styleId="Quote"><w:name w:val="Quote"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:pBdr><w:left w:val="single" w:sz="24" w:space="8" w:color="B4C6E7"/></w:pBdr><w:ind w:left="360"/></w:pPr><w:rPr><w:i/><w:color w:val="404040"/></w:rPr></w:style>' +
        '<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:spacing w:after="40"/><w:contextualSpacing/></w:pPr></w:style>' +
        '<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr></w:style>' +
        "</w:styles>";

      var z = new JSZip();
      z.file("[Content_Types].xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' +
        '<Default Extension="png" ContentType="image/png"/><Default Extension="jpeg" ContentType="image/jpeg"/><Default Extension="gif" ContentType="image/gif"/>' +
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
        '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
        '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>' +
        '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' +
        '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>');
      z.file("_rels/.rels", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="' + R + '/officeDocument" Target="word/document.xml"/>' +
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>' +
        '<Relationship Id="rId3" Type="' + R + '/extended-properties" Target="docProps/app.xml"/></Relationships>');
      var now = new Date().toISOString().replace(/\.\d+Z$/, "Z");
      z.file("docProps/core.xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
        "<dc:title>" + x(title) + "</dc:title><dc:creator>Homework Launchpad</dc:creator>" +
        '<dcterms:created xsi:type="dcterms:W3CDTF">' + now + '</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">' + now + "</dcterms:modified></cp:coreProperties>");
      z.file("docProps/app.xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Launchpad Write</Application></Properties>');
      z.file("word/document.xml", doc);
      z.file("word/styles.xml", styles);
      z.file("word/numbering.xml", numbering);
      z.file("word/_rels/document.xml.rels", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="' + R + '/styles" Target="styles.xml"/><Relationship Id="rId2" Type="' + R + '/numbering" Target="numbering.xml"/>' + rels.join("") + "</Relationships>");
      media.forEach(function (m) { z.file("word/media/" + m.name, m.data, { base64: true }); });
      return z.generateAsync({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document", compression: "DEFLATE" });
    });
  }

  /* ================= reading ================= */
  function kids(el, name) { var o = []; for (var n = el && el.firstChild; n; n = n.nextSibling) if (n.nodeType === 1 && (!name || n.localName === name)) o.push(n); return o; }
  function kid(el, name) { return kids(el, name)[0] || null; }
  function att(el, name) { return el ? (el.getAttributeNS(W, name) || el.getAttribute("w:" + name)) : null; }
  function e(s) { return String(s).replace(/[&<>"]/g, function (c) { return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]; }); }
  var HL = { yellow: "#ffff00", green: "#00ff00", cyan: "#00ffff", magenta: "#ff00ff", blue: "#0000ff", red: "#ff0000", darkBlue: "#000080",
    darkCyan: "#008080", darkGreen: "#008000", darkMagenta: "#800080", darkRed: "#800000", darkYellow: "#808000", darkGray: "#808080", lightGray: "#c0c0c0", black: "#000000" };

  function read(file) {
    return JSZip.loadAsync(file).then(function (z) {
      var docF = z.file("word/document.xml"); if (!docF) throw new Error("this doesn't look like a Word file");
      var get = function (p) { var f = z.file(p); return f ? f.async("string") : Promise.resolve(null); };
      return Promise.all([docF.async("string"), get("word/_rels/document.xml.rels"), get("word/styles.xml"), get("word/numbering.xml")]).then(function (a) {
        var P = new DOMParser();
        var doc = P.parseFromString(a[0], "application/xml");
        var rels = {}, styles = {}, numFmt = {};
        if (a[1]) Array.prototype.forEach.call(P.parseFromString(a[1], "application/xml").getElementsByTagName("Relationship"), function (r) {
          rels[r.getAttribute("Id")] = { t: r.getAttribute("Target"), ext: r.getAttribute("TargetMode") === "External" };
        });
        if (a[2]) Array.prototype.forEach.call(P.parseFromString(a[2], "application/xml").getElementsByTagNameNS(W, "style"), function (s) {
          var n = kid(s, "name"); styles[att(s, "styleId")] = (n ? att(n, "val") : "").toLowerCase();
        });
        if (a[3]) {
          var nd = P.parseFromString(a[3], "application/xml"), absF = {};
          Array.prototype.forEach.call(nd.getElementsByTagNameNS(W, "abstractNum"), function (an) {
            var lv = {}; kids(an, "lvl").forEach(function (l) { var f = kid(l, "numFmt"); lv[att(l, "ilvl")] = f ? att(f, "val") : "bullet"; });
            absF[att(an, "abstractNumId")] = lv;
          });
          Array.prototype.forEach.call(nd.getElementsByTagNameNS(W, "num"), function (n) {
            var ab = kid(n, "abstractNumId"); numFmt[att(n, "numId")] = absF[ab ? att(ab, "val") : ""] || {};
          });
        }
        // pictures: load the ones the text uses
        var need = {};
        Array.prototype.forEach.call(doc.getElementsByTagNameNS("http://schemas.openxmlformats.org/drawingml/2006/main", "blip"), function (b) {
          var id = b.getAttributeNS(R, "embed"); if (id && rels[id]) need[id] = rels[id].t;
        });
        var pics = {};
        return Promise.all(Object.keys(need).map(function (id) {
          var path = need[id].replace(/^\//, ""); if (!/^word\//.test(path)) path = "word/" + path;
          var ext = (path.match(/\.(png|jpe?g|gif)$/i) || [])[1]; var f = z.file(path);
          if (!ext || !f) return null;
          return f.async("base64").then(function (b) { pics[id] = "data:image/" + (ext.toLowerCase() === "png" ? "png" : ext.toLowerCase() === "gif" ? "gif" : "jpeg") + ";base64," + b; });
        })).then(function () { return convert(doc, rels, styles, numFmt, pics); });
      });
    });
  }

  function convert(doc, rels, styles, numFmt, pics) {
    var body = doc.getElementsByTagNameNS(W, "body")[0]; if (!body) return "";
    var out = [], listStack = [];   // [{tag, lvl}]
    function closeLists(to) { while (listStack.length > to) { out.push("</li></" + listStack.pop().tag + ">"); } }

    function runs(el) {
      var o = "";
      kids(el).forEach(function (n) {
        var ln = n.localName;
        if (ln === "r") o += run(n);
        else if (ln === "hyperlink") {
          var id = n.getAttributeNS(R, "id"), r = id && rels[id], inner = runs(n);
          o += r && /^(https?:|mailto:)/i.test(r.t) ? '<a href="' + e(r.t) + '">' + inner + "</a>" : inner;
        } else if (/^(ins|smartTag|sdt|sdtContent|fldSimple|customXml|bdo|dir)$/.test(ln)) o += runs(n);
      });
      return o;
    }
    function run(r) {
      var pr = kid(r, "rPr"), t = "";
      kids(r).forEach(function (c) {
        var ln = c.localName;
        if (ln === "t") t += e(c.textContent);
        else if (ln === "tab") t += "&nbsp;&nbsp;&nbsp;&nbsp;";
        else if (ln === "br" || ln === "cr") t += "<br>";
        else if (ln === "drawing" || ln === "pict") {
          var b = c.getElementsByTagNameNS("http://schemas.openxmlformats.org/drawingml/2006/main", "blip")[0];
          var id = b && b.getAttributeNS(R, "embed"), ext = c.getElementsByTagNameNS("http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing", "extent")[0];
          if (id && pics[id]) t += '<img src="' + pics[id] + '" alt=""' + (ext ? ' width="' + Math.round(+ext.getAttribute("cx") / 9525) + '"' : "") + ">";
        } else if (ln === "sym") t += "&#x" + (att(c, "char") || "20").replace(/^F0/i, "") + ";";
      });
      if (!t || !pr) return t;
      var st = [];
      var on = function (n) { var k = kid(pr, n); return k && att(k, "val") !== "0" && att(k, "val") !== "false" && att(k, "val") !== "none"; };
      var c = kid(pr, "color"), sz = kid(pr, "sz"), hl = kid(pr, "highlight"), shd = kid(pr, "shd"), va = kid(pr, "vertAlign"), f = kid(pr, "rFonts");
      if (c && /^[0-9a-f]{6}$/i.test(att(c, "val") || "")) st.push("color:#" + att(c, "val"));
      if (sz && +att(sz, "val")) st.push("font-size:" + (+att(sz, "val") / 2) + "pt");
      if (hl && HL[att(hl, "val")]) st.push("background-color:" + HL[att(hl, "val")]);
      else if (shd && /^[0-9a-f]{6}$/i.test(att(shd, "fill") || "")) st.push("background-color:#" + att(shd, "fill"));
      if (f && att(f, "ascii")) st.push("font-family:" + att(f, "ascii").replace(/[^a-z0-9 -]/gi, ""));
      if (on("b")) t = "<b>" + t + "</b>";
      if (on("i")) t = "<i>" + t + "</i>";
      if (on("u")) t = "<u>" + t + "</u>";
      if (on("strike") || on("dstrike")) t = "<s>" + t + "</s>";
      if (va && att(va, "val") === "superscript") t = "<sup>" + t + "</sup>";
      if (va && att(va, "val") === "subscript") t = "<sub>" + t + "</sub>";
      return st.length ? '<span style="' + st.join(";") + '">' + t + "</span>" : t;
    }
    function para(p) {
      var pr = kid(p, "pPr"), ps = pr && kid(pr, "pStyle"), sid = ps ? att(ps, "val") : "", name = styles[sid] || (sid || "").toLowerCase();
      var jc = pr && kid(pr, "jc"), al = jc ? { center: "center", right: "right", end: "right", both: "justify", distribute: "justify" }[att(jc, "val")] : "";
      var inner = runs(p) || "<br>", style = al ? ' style="text-align:' + al + '"' : "";
      var numPr = pr && kid(pr, "numPr");
      if (numPr) {
        var lvl = +att(kid(numPr, "ilvl"), "val") || 0, nid = att(kid(numPr, "numId"), "val");
        if (nid && nid !== "0") {
          var fmt = (numFmt[nid] || {})[lvl] || "bullet", tag = fmt === "bullet" || fmt === "none" ? "ul" : "ol";
          if (listStack.length > lvl + 1) closeLists(lvl + 1);
          if (listStack.length === lvl + 1 && listStack[lvl].tag !== tag) closeLists(lvl);
          if (listStack.length === lvl + 1) out.push("</li>");
          while (listStack.length < lvl + 1) { out.push("<" + tag + ">"); listStack.push({ tag: tag }); }
          out.push("<li" + style + ">" + inner);
          return;
        }
      }
      closeLists(0);
      var m = name.match(/^heading ([1-6])$/);
      if (m) { var h = Math.min(3, +m[1]); out.push("<h" + h + style + ">" + inner + "</h" + h + ">"); }
      else if (name === "title") out.push('<p class="doc-title"' + style + ">" + inner + "</p>");
      else if (name === "subtitle") out.push("<h2" + style + ">" + inner + "</h2>");
      else if (/quote/.test(name)) out.push("<blockquote" + style + ">" + inner + "</blockquote>");
      else out.push("<p" + style + ">" + inner + "</p>");
    }
    function tbl(t) {
      closeLists(0);
      var o = "<table><tbody>";
      kids(t, "tr").forEach(function (tr) {
        o += "<tr>";
        kids(tr, "tc").forEach(function (tc) {
          var pr = kid(tc, "tcPr"), gs = pr && kid(pr, "gridSpan"), span = gs ? +att(gs, "val") : 1;
          var save = out; out = []; walk(tc); closeLists(0); var inner = out.join(""); out = save;
          o += "<td" + (span > 1 ? ' colspan="' + span + '"' : "") + ">" + (inner || "<br>") + "</td>";
        });
        o += "</tr>";
      });
      out.push(o + "</tbody></table>");
    }
    function walk(box) {
      kids(box).forEach(function (n) {
        if (n.localName === "p") para(n);
        else if (n.localName === "tbl") tbl(n);
        else if (/^(sdt|sdtContent|customXml|ins)$/.test(n.localName)) walk(n);
      });
    }
    walk(body); closeLists(0);
    return out.join("");
  }

  window.Docx = { build: build, read: read };
})();
