/* Excel files (.xlsx) for Launchpad Grid, written from scratch on JSZip.
   Keeps values, formulas, bold/italic/underline, colours, fills,
   alignment, number formats, column widths and every sheet tab. */
(function () {
  var NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
  var RNS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  function x(s) { return String(s).replace(/[&<>"]/g, function (c) { return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]; })
    .replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g, ""); }
  var NF = { n2: 4, int: 3, pct: 10, gbp: 164 };
  function argb(h) { h = String(h || "").replace("#", ""); if (h.length === 3) h = h.replace(/./g, "$&$&"); return /^[0-9a-f]{6}$/i.test(h) ? "FF" + h.toUpperCase() : null; }

  /* sheets: [{name, cells, w}], value(sheetIndex, key) -> computed value */
  function build(sheets, value) {
    var fonts = ['<font><sz val="11"/><name val="Calibri"/></font>'], fontKey = { "": 0 };
    var fills = ['<fill><patternFill patternType="none"/></fill>', '<fill><patternFill patternType="gray125"/></fill>'], fillKey = { "": 0 };
    var xfs = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'], xfKey = { "": 0 };
    function styleOf(c) {
      var fk = (c.b ? "b" : "") + (c.i ? "i" : "") + (c.u ? "u" : "") + (c.fg ? c.fg : "");
      if (!(fk in fontKey)) {
        var fg = argb(c.fg);
        fontKey[fk] = fonts.length;
        fonts.push("<font>" + (c.b ? "<b/>" : "") + (c.i ? "<i/>" : "") + (c.u ? "<u/>" : "") + '<sz val="11"/>' + (fg ? '<color rgb="' + fg + '"/>' : "") + '<name val="Calibri"/></font>');
      }
      var bg = argb(c.bg), lk = bg || "";
      if (!(lk in fillKey)) { fillKey[lk] = fills.length; fills.push('<fill><patternFill patternType="solid"><fgColor rgb="' + bg + '"/><bgColor indexed="64"/></patternFill></fill>'); }
      var nf = NF[c.nf] || 0, al = { l: "left", c: "center", r: "right" }[c.al] || "";
      var k = fontKey[fk] + "/" + fillKey[lk] + "/" + nf + "/" + al;
      if (k === "0/0/0/") return 0;
      if (!(k in xfKey)) {
        xfKey[k] = xfs.length;
        xfs.push('<xf numFmtId="' + nf + '" fontId="' + fontKey[fk] + '" fillId="' + fillKey[lk] + '" borderId="0" xfId="0"' +
          (nf ? ' applyNumberFormat="1"' : "") + (fontKey[fk] ? ' applyFont="1"' : "") + (fillKey[lk] ? ' applyFill="1"' : "") +
          (al ? ' applyAlignment="1"><alignment horizontal="' + al + '"/></xf>' : "/>"));
      }
      return xfKey[k];
    }

    var z = new JSZip(), sheetFiles = [];
    sheets.forEach(function (sh, si) {
      var rows = {};
      Object.keys(sh.cells).forEach(function (k) {
        var p = Formula.parseRef(k); if (!p) return;
        var c = sh.cells[k]; if (!c || (c.v === "" || c.v == null) && !c.b && !c.i && !c.u && !c.bg && !c.fg && !c.nf && !c.al) return;
        (rows[p.r] = rows[p.r] || []).push({ c: p.c, k: k, cell: c });
      });
      var o = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="' + NS + '" xmlns:r="' + RNS + '">';
      if (si === 0) o += '<sheetViews><sheetView workbookViewId="0" tabSelected="1"/></sheetViews>';
      var ws = Object.keys(sh.w || {});
      if (ws.length) {
        o += "<cols>" + ws.map(function (col) { var n = Formula.colNum(col) + 1; return '<col min="' + n + '" max="' + n + '" width="' + (Math.round(sh.w[col] / 7 * 100) / 100) + '" customWidth="1"/>'; })
          .sort(function (a, b) { return +a.match(/min="(\d+)"/)[1] - +b.match(/min="(\d+)"/)[1]; }).join("") + "</cols>";
      }
      o += "<sheetData>";
      Object.keys(rows).map(Number).sort(function (a, b) { return a - b; }).forEach(function (r) {
        o += '<row r="' + (r + 1) + '">';
        rows[r].sort(function (a, b) { return a.c - b.c; }).forEach(function (e) {
          var c = e.cell, s = styleOf(c), sa = s ? ' s="' + s + '"' : "", raw = c.v == null ? "" : String(c.v);
          if (raw === "") { o += '<c r="' + e.k + '"' + sa + "/>"; return; }
          if (raw[0] === "=" && raw.length > 1) {
            var v = value(si, e.k), f = "<f>" + x(raw.slice(1)) + "</f>";
            if (Formula.isErr(v)) o += '<c r="' + e.k + '"' + sa + ' t="e">' + f + "<v>" + x(v.err) + "</v></c>";
            else if (typeof v === "number") o += '<c r="' + e.k + '"' + sa + ">" + f + "<v>" + v + "</v></c>";
            else if (typeof v === "boolean") o += '<c r="' + e.k + '"' + sa + ' t="b">' + f + "<v>" + (v ? 1 : 0) + "</v></c>";
            else o += '<c r="' + e.k + '"' + sa + ' t="str">' + f + "<v>" + x(v) + "</v></c>";
            return;
          }
          var n = value(si, e.k);
          if (typeof n === "number") o += '<c r="' + e.k + '"' + sa + "><v>" + n + "</v></c>";
          else if (typeof n === "boolean") o += '<c r="' + e.k + '"' + sa + ' t="b"><v>' + (n ? 1 : 0) + "</v></c>";
          else o += '<c r="' + e.k + '"' + sa + ' t="inlineStr"><is><t xml:space="preserve">' + x(raw) + "</t></is></c>";
        });
        o += "</row>";
      });
      o += '</sheetData><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>';
      z.file("xl/worksheets/sheet" + (si + 1) + ".xml", o);
      sheetFiles.push(si + 1);
    });

    var names = {};
    var wb = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="' + NS + '" xmlns:r="' + RNS + '"><bookViews><workbookView/></bookViews><sheets>' +
      sheets.map(function (sh, i) {
        var n = String(sh.name || "Sheet" + (i + 1)).replace(/[\\\/?*\[\]:]/g, " ").slice(0, 31) || "Sheet" + (i + 1);
        while (names[n.toLowerCase()]) n = n.slice(0, 28) + " " + (i + 1); names[n.toLowerCase()] = 1;
        return '<sheet name="' + x(n) + '" sheetId="' + (i + 1) + '" r:id="rId' + (i + 1) + '"/>';
      }).join("") + "</sheets><calcPr calcId=\"191029\" fullCalcOnLoad=\"1\"/></workbook>";
    var wbRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      sheetFiles.map(function (n) { return '<Relationship Id="rId' + n + '" Type="' + RNS + '/worksheet" Target="worksheets/sheet' + n + '.xml"/>'; }).join("") +
      '<Relationship Id="rId' + (sheets.length + 1) + '" Type="' + RNS + '/styles" Target="styles.xml"/></Relationships>';
    var styles = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="' + NS + '">' +
      '<numFmts count="1"><numFmt numFmtId="164" formatCode="&quot;£&quot;#,##0.00"/></numFmts>' +
      '<fonts count="' + fonts.length + '">' + fonts.join("") + "</fonts>" +
      '<fills count="' + fills.length + '">' + fills.join("") + "</fills>" +
      '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
      '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
      '<cellXfs count="' + xfs.length + '">' + xfs.join("") + "</cellXfs>" +
      '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>';

    z.file("[Content_Types].xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
      sheetFiles.map(function (n) { return '<Override PartName="/xl/worksheets/sheet' + n + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'; }).join("") +
      '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>');
    z.file("_rels/.rels", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="' + RNS + '/officeDocument" Target="xl/workbook.xml"/></Relationships>');
    z.file("xl/workbook.xml", wb);
    z.file("xl/_rels/workbook.xml.rels", wbRels);
    z.file("xl/styles.xml", styles);
    return z.generateAsync({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", compression: "DEFLATE" });
  }

  /* ---------- reading ---------- */
  function els(doc, name) { return Array.prototype.slice.call(doc.getElementsByTagNameNS(NS, name)); }
  function first(el, name) { var l = el.getElementsByTagNameNS(NS, name); return l.length ? l[0] : null; }
  function read(file) {
    return JSZip.loadAsync(file).then(function (z) {
      var P = new DOMParser(), xml = function (s) { return P.parseFromString(s, "application/xml"); };
      var get = function (p) { var f = z.file(p); return f ? f.async("string") : Promise.resolve(null); };
      return Promise.all([get("xl/workbook.xml"), get("xl/_rels/workbook.xml.rels"), get("xl/sharedStrings.xml"), get("xl/styles.xml")]).then(function (a) {
        if (!a[0]) throw new Error("this doesn't look like an Excel file");
        var wb = xml(a[0]), rels = {};
        if (a[1]) Array.prototype.forEach.call(xml(a[1]).getElementsByTagName("Relationship"), function (r) { rels[r.getAttribute("Id")] = r.getAttribute("Target"); });
        var shared = [];
        if (a[2]) els(xml(a[2]), "si").forEach(function (si) {
          shared.push(Array.prototype.map.call(si.getElementsByTagNameNS(NS, "t"), function (t) {
            return t.parentNode.localName === "rPh" ? "" : t.textContent; }).join(""));
        });
        var xfs = [];
        if (a[3]) {
          var st = xml(a[3]), fonts = [], fills = [], fmts = {};
          els(st, "numFmt").forEach(function (n) { fmts[n.getAttribute("numFmtId")] = n.getAttribute("formatCode") || ""; });
          var fontsEl = first(st, "fonts"); if (fontsEl) Array.prototype.forEach.call(fontsEl.children, function (f) {
            var col = first(f, "color"), rgb = col && col.getAttribute("rgb");
            var on = function (n) { var e = first(f, n); return e && e.getAttribute("val") !== "0" && e.getAttribute("val") !== "none"; };
            fonts.push({ b: on("b"), i: on("i"), u: on("u"), fg: rgb && rgb.length === 8 && rgb.slice(2) !== "000000" ? "#" + rgb.slice(2).toLowerCase() : null });
          });
          var fillsEl = first(st, "fills"); if (fillsEl) Array.prototype.forEach.call(fillsEl.children, function (f) {
            var pf = first(f, "patternFill"), fg = pf && pf.getAttribute("patternType") === "solid" && first(pf, "fgColor"), rgb = fg && fg.getAttribute("rgb");
            fills.push(rgb && rgb.length === 8 ? "#" + rgb.slice(2).toLowerCase() : null);
          });
          var cx = first(st, "cellXfs"); if (cx) Array.prototype.forEach.call(cx.children, function (xf) {
            var f = fonts[+xf.getAttribute("fontId") || 0] || {}, al = first(xf, "alignment"), h = al && al.getAttribute("horizontal");
            var id = +xf.getAttribute("numFmtId") || 0, code = fmts[id] || "", nf = null;
            if (id === 9 || id === 10 || /%/.test(code)) nf = "pct";
            else if (/£/.test(code) || id === 7 || id === 8 || id === 5 || id === 6) nf = "gbp";
            else if (id === 4 || id === 2 || /0\.00/.test(code)) nf = "n2";
            else if (id === 3 || id === 1) nf = "int";
            xfs.push({ b: f.b, i: f.i, u: f.u, fg: f.fg, bg: fills[+xf.getAttribute("fillId") || 0] || null, nf: nf,
              al: { left: "l", center: "c", centerContinuous: "c", right: "r" }[h] || null });
          });
        }
        var list = els(wb, "sheet").slice(0, 20);
        return Promise.all(list.map(function (s) {
          var id = s.getAttributeNS(RNS, "id") || s.getAttribute("r:id"), t = rels[id] || "";
          var path = t.replace(/^\/?(xl\/)?/, "xl/");
          return get(path).then(function (body) { return { name: s.getAttribute("name"), body: body }; });
        })).then(function (got) {
          var out = [];
          got.forEach(function (g) {
            if (!g.body) return;
            var d = xml(g.body), cells = {}, w = {}, maxR = 0, maxC = 0, n = 0;
            els(d, "col").forEach(function (c) {
              var lo = +c.getAttribute("min"), hi = Math.min(+c.getAttribute("max"), 52), wd = +c.getAttribute("width");
              if (wd && lo) for (var i = lo; i <= hi; i++) w[Formula.colName(i - 1)] = Math.max(30, Math.min(600, Math.round(wd * 7)));
            });
            els(d, "c").forEach(function (c) {
              if (n > 60000) return;
              var r = c.getAttribute("r"), p = Formula.parseRef(r || ""); if (!p || p.r > 4999 || p.c > 51) return;
              var t = c.getAttribute("t"), v = first(c, "v"), f = first(c, "f"), val = "";
              if (f && f.textContent) val = "=" + f.textContent;
              else if (t === "s" && v) val = shared[+v.textContent] || "";
              else if (t === "inlineStr") { var is = first(c, "is"); val = is ? is.textContent : ""; }
              else if (t === "b" && v) val = v.textContent === "1" ? "TRUE" : "FALSE";
              else if (v) val = v.textContent;
              var st = xfs[+c.getAttribute("s") || 0] || {}, cell = {};
              if (val !== "") cell.v = val;
              ["b", "i", "u", "fg", "bg", "nf", "al"].forEach(function (k) { if (st[k]) cell[k] = st[k]; });
              if (Object.keys(cell).length) { cells[Formula.refName(p.r, p.c)] = cell; n++; maxR = Math.max(maxR, p.r); maxC = Math.max(maxC, p.c); }
            });
            out.push({ name: (g.name || "Sheet" + (out.length + 1)).slice(0, 31), cells: cells, w: w, rows: Math.max(100, maxR + 21), cols: Math.max(26, maxC + 1) });
          });
          if (!out.length) throw new Error("there were no sheets in it");
          return out;
        });
      });
    });
  }

  window.Xlsx = { build: build, read: read };
})();
