/* PowerPoint files (.pptx) for Launchpad Present, written from scratch on
   JSZip. A slide unit is 1 pt = 12700 EMU, so 960 x 540 is 16:9. */
(function () {
  var NSA = "http://schemas.openxmlformats.org/drawingml/2006/main";
  var NSR = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  var NSP = "http://schemas.openxmlformats.org/presentationml/2006/main";
  var EMU = 12700;
  var XMLNS = 'xmlns:a="' + NSA + '" xmlns:r="' + NSR + '" xmlns:p="' + NSP + '"';
  var HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
  function x(s) { return String(s).replace(/[&<>"]/g, function (c) { return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]; })
    .replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g, ""); }
  function hex(c) {
    if (!c) return null; c = String(c).trim();
    var m = c.match(/^#([0-9a-f]{3}|[0-9a-f]{6})$/i); if (m) { var h = m[1]; if (h.length === 3) h = h.replace(/./g, "$&$&"); return h.toUpperCase(); }
    m = c.match(/^rgba?\(\s*(\d+)[,\s]+(\d+)[,\s]+(\d+)/i);
    if (m) return [m[1], m[2], m[3]].map(function (n) { return ("0" + (+n).toString(16)).slice(-2); }).join("").toUpperCase();
    return null;
  }
  function font1(f) { return String(f || "Calibri").split(",")[0].replace(/["']/g, "").trim(); }
  function e(n) { return Math.round(n * EMU); }
  var GRP = '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>';

  var THEME = HEAD + '<a:theme xmlns:a="' + NSA + '" name="Launchpad"><a:themeElements>' +
    '<a:clrScheme name="Launchpad"><a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="1F2937"/></a:dk2><a:lt2><a:srgbClr val="F3F4F6"/></a:lt2>' +
    '<a:accent1><a:srgbClr val="2563EB"/></a:accent1><a:accent2><a:srgbClr val="EA580C"/></a:accent2><a:accent3><a:srgbClr val="16A34A"/></a:accent3><a:accent4><a:srgbClr val="7C3AED"/></a:accent4><a:accent5><a:srgbClr val="DB2777"/></a:accent5><a:accent6><a:srgbClr val="0891B2"/></a:accent6>' +
    '<a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme>' +
    '<a:fontScheme name="Launchpad"><a:majorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>' +
    '<a:fmtScheme name="Launchpad"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>' +
    '<a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>' +
    '<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>' +
    '<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme>' +
    '</a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>';
  var LVL = '<a:lvl1pPr><a:defRPr sz="1800"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/></a:defRPr></a:lvl1pPr>';
  var MASTER = HEAD + '<p:sldMaster ' + XMLNS + '><p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg><p:spTree>' + GRP + '</p:spTree></p:cSld>' +
    '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>' +
    '<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>' +
    '<p:txStyles><p:titleStyle>' + LVL + '</p:titleStyle><p:bodyStyle>' + LVL + '</p:bodyStyle><p:otherStyle>' + LVL + '</p:otherStyle></p:txStyles></p:sldMaster>';
  var LAYOUT = HEAD + '<p:sldLayout ' + XMLNS + ' type="blank" preserve="1"><p:cSld name="Blank"><p:spTree>' + GRP + '</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>';

  function textBody(el, def) {
    // html -> paragraphs of runs
    var box = document.createElement("div"); box.innerHTML = el.html || "";
    var paras = [];
    function newPara(bullet, align) { var p = { bullet: bullet, align: align, runs: [] }; paras.push(p); return p; }
    function runs(node, p, st) {
      for (var n = node.firstChild; n; n = n.nextSibling) {
        if (n.nodeType === 3) { if (n.nodeValue) p.runs.push({ t: n.nodeValue.replace(/ /g, " "), st: st }); continue; }
        if (n.nodeType !== 1) continue;
        if (n.tagName === "BR") { p.runs.push({ br: true }); continue; }
        if (/^(UL|OL|P|DIV|LI)$/.test(n.tagName)) continue;
        var s = Object.assign({}, st), cs = n.style;
        if (/^(B|STRONG)$/.test(n.tagName) || /bold|[6-9]00/.test(cs.fontWeight)) s.b = 1;
        if (/^(I|EM)$/.test(n.tagName) || cs.fontStyle === "italic") s.i = 1;
        if (n.tagName === "U" || /underline/.test(cs.textDecorationLine || cs.textDecoration || "")) s.u = 1;
        if (cs.color && hex(cs.color)) s.color = hex(cs.color);
        runs(n, p, s);
      }
    }
    function blocks(node, bullet) {
      var loose = null;
      for (var n = node.firstChild; n; n = n.nextSibling) {
        if (n.nodeType === 1 && (n.tagName === "UL" || n.tagName === "OL")) { loose = null; Array.prototype.forEach.call(n.children, function (li) { var p = newPara(n.tagName === "OL" ? "num" : "dot", li.style.textAlign); runs(li, p, {}); blocks2(li); }); continue; }
        if (n.nodeType === 1 && /^(P|DIV)$/.test(n.tagName)) { loose = null; if (n.querySelector("ul,ol,p,div")) blocks(n, bullet); else runs(n, newPara(bullet, n.style.textAlign), {}); continue; }
        if (!loose) loose = newPara(bullet, null);
        var tmp = document.createElement("span"); tmp.appendChild(n.cloneNode(true)); runs(tmp, loose, {});
      }
    }
    function blocks2(li) { Array.prototype.forEach.call(li.children, function (c) { if (c.tagName === "UL" || c.tagName === "OL") { var w = document.createElement("div"); w.appendChild(c.cloneNode(true)); blocks(w, "dot"); } }); }
    blocks(box, null);
    if (!paras.length) paras.push({ runs: [] });
    var algn = { left: "l", center: "ctr", right: "r", justify: "just" };
    var out = "";
    paras.forEach(function (p) {
      var a = algn[p.align || el.align] || "l";
      var pPr = '<a:pPr algn="' + a + '"' + (p.bullet ? ' marL="' + e(el.size * 0.9) + '" indent="-' + e(el.size * 0.9) + '"><a:buFont typeface="Arial"/>' +
        (p.bullet === "num" ? '<a:buAutoNum type="arabicPeriod"/>' : '<a:buChar char="•"/>') + "</a:pPr>" : "><a:buNone/></a:pPr>");
      out += "<a:p>" + pPr;
      p.runs.forEach(function (r) {
        if (r.br) { out += "<a:br/>"; return; }
        var st = r.st || {}, col = st.color || def.color;
        out += '<a:r><a:rPr lang="en-GB" sz="' + Math.round(el.size * 100) + '"' + (st.b || def.b ? ' b="1"' : "") + (st.i ? ' i="1"' : "") + (st.u ? ' u="sng"' : "") + ' dirty="0">' +
          '<a:solidFill><a:srgbClr val="' + col + '"/></a:solidFill><a:latin typeface="' + x(def.font) + '"/><a:cs typeface="' + x(def.font) + '"/></a:rPr><a:t>' + x(r.t) + "</a:t></a:r>";
      });
      out += '<a:endParaRPr lang="en-GB" sz="' + Math.round(el.size * 100) + '" dirty="0"/></a:p>';
    });
    return out;
  }

  var GEOM = { rect: "rect", round: "roundRect", ellipse: "ellipse", triangle: "triangle", star: "star5", arrow: "rightArrow", line: "line" };
  function natural(src) {
    return new Promise(function (ok) { var im = new Image(); im.onload = function () { ok({ w: im.naturalWidth, h: im.naturalHeight }); }; im.onerror = function () { ok(null); }; im.src = src; });
  }

  function build(deck, THEMES, title) {
    var t = THEMES[deck.theme] || THEMES.classic, z = new JSZip(), mediaN = 0, jobs = [];
    deck.slides.forEach(function (s) { s.els.forEach(function (el) { if (el.type === "img") jobs.push(natural(el.src).then(function (n) { el._nat = n; })); }); });
    return Promise.all(jobs).then(function () {
      var ctOver = "";
      deck.slides.forEach(function (s, si) {
        var rels = ['<Relationship Id="rId1" Type="' + NSR + '/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>'], id = 2, tree = "";
        s.els.forEach(function (el) {
          var nid = id++;
          var xfrm = '<a:xfrm><a:off x="' + e(el.x) + '" y="' + e(el.y) + '"/><a:ext cx="' + e(el.w) + '" cy="' + e(el.h) + '"/></a:xfrm>';
          if (el.type === "text") {
            if (!el.html) return;
            var def = { color: hex(el.color) || hex(el.role === "title" ? t.title : t.text) || "000000", font: font1(el.role === "title" ? t.tf : t.bf), b: el.role === "title" };
            tree += '<p:sp><p:nvSpPr><p:cNvPr id="' + nid + '" name="' + (el.role === "title" ? "Title " : "Text ") + nid + '"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>' +
              '<p:spPr>' + xfrm + '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>' +
              '<p:txBody><a:bodyPr wrap="square" lIns="127000" tIns="76200" rIns="127000" bIns="76200" rtlCol="0" anchor="t"><a:normAutofit/></a:bodyPr><a:lstStyle/>' + textBody(el, def) + "</p:txBody></p:sp>";
          } else if (el.type === "shape") {
            var c = hex(el.color) || hex(t.accent) || "2563EB", line = el.shape === "line";
            var lx = line ? '<a:xfrm><a:off x="' + e(el.x) + '" y="' + e(el.y + el.h / 2) + '"/><a:ext cx="' + e(el.w) + '" cy="0"/></a:xfrm>' : xfrm;
            tree += '<p:sp><p:nvSpPr><p:cNvPr id="' + nid + '" name="Shape ' + nid + '"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr>' + lx +
              '<a:prstGeom prst="' + (GEOM[el.shape] || "rect") + '"><a:avLst/></a:prstGeom>' +
              (line ? '<a:noFill/><a:ln w="50800"><a:solidFill><a:srgbClr val="' + c + '"/></a:solidFill></a:ln>' : '<a:solidFill><a:srgbClr val="' + c + '"/></a:solidFill><a:ln><a:noFill/></a:ln>') +
              "</p:spPr></p:sp>";
          } else if (el.type === "img") {
            var m = String(el.src).match(/^data:image\/(png|jpe?g|gif|webp);base64,(.+)$/i); if (!m || /webp/i.test(m[1])) return;
            var ext = /png/i.test(m[1]) ? "png" : /gif/i.test(m[1]) ? "gif" : "jpeg", name = "image" + (++mediaN) + "." + ext;
            z.file("ppt/media/" + name, m[2], { base64: true });
            var rid = "rId" + (rels.length + 1);
            rels.push('<Relationship Id="' + rid + '" Type="' + NSR + '/image" Target="../media/' + name + '"/>');
            var bx = el.x, by = el.y, bw = el.w, bh = el.h, n = el._nat; delete el._nat;
            if (n && n.w && n.h) { var k = Math.min(el.w / n.w, el.h / n.h); bw = n.w * k; bh = n.h * k; bx = el.x + (el.w - bw) / 2; by = el.y + (el.h - bh) / 2; }
            tree += '<p:pic><p:nvPicPr><p:cNvPr id="' + nid + '" name="Picture ' + nid + '"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>' +
              '<p:blipFill><a:blip r:embed="' + rid + '"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>' +
              '<p:spPr><a:xfrm><a:off x="' + e(bx) + '" y="' + e(by) + '"/><a:ext cx="' + e(bw) + '" cy="' + e(bh) + '"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>';
          }
        });
        var bg = t.bg.length > 1 ?
          '<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:srgbClr val="' + hex(t.bg[0]) + '"/></a:gs><a:gs pos="100000"><a:srgbClr val="' + hex(t.bg[1]) + '"/></a:gs></a:gsLst><a:lin ang="2700000" scaled="0"/></a:gradFill>' :
          '<a:solidFill><a:srgbClr val="' + hex(t.bg[0]) + '"/></a:solidFill>';
        var tr = { fade: "<p:fade/>", push: '<p:push dir="l"/>', zoom: "<p:zoom/>" }[s.trans];
        z.file("ppt/slides/slide" + (si + 1) + ".xml", HEAD + '<p:sld ' + XMLNS + '><p:cSld><p:bg><p:bgPr>' + bg + '<a:effectLst/></p:bgPr></p:bg><p:spTree>' + GRP + tree + '</p:spTree></p:cSld>' +
          '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>' + (tr ? '<p:transition spd="med">' + tr + "</p:transition>" : "") + "</p:sld>");
        z.file("ppt/slides/_rels/slide" + (si + 1) + ".xml.rels", HEAD + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + rels.join("") + "</Relationships>");
        ctOver += '<Override PartName="/ppt/slides/slide' + (si + 1) + '.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>';
      });
      var n = deck.slides.length;
      var presRels = ['<Relationship Id="rId1" Type="' + NSR + '/slideMaster" Target="slideMasters/slideMaster1.xml"/>'];
      var sldIds = "";
      for (var i = 0; i < n; i++) { presRels.push('<Relationship Id="rId' + (i + 2) + '" Type="' + NSR + '/slide" Target="slides/slide' + (i + 1) + '.xml"/>'); sldIds += '<p:sldId id="' + (256 + i) + '" r:id="rId' + (i + 2) + '"/>'; }
      presRels.push('<Relationship Id="rId' + (n + 2) + '" Type="' + NSR + '/presProps" Target="presProps.xml"/>');
      presRels.push('<Relationship Id="rId' + (n + 3) + '" Type="' + NSR + '/viewProps" Target="viewProps.xml"/>');
      presRels.push('<Relationship Id="rId' + (n + 4) + '" Type="' + NSR + '/theme" Target="theme/theme1.xml"/>');
      presRels.push('<Relationship Id="rId' + (n + 5) + '" Type="' + NSR + '/tableStyles" Target="tableStyles.xml"/>');
      z.file("ppt/presentation.xml", HEAD + '<p:presentation ' + XMLNS + ' saveSubsetFonts="1"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>' +
        "<p:sldIdLst>" + sldIds + '</p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>');
      z.file("ppt/_rels/presentation.xml.rels", HEAD + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + presRels.join("") + "</Relationships>");
      z.file("ppt/presProps.xml", HEAD + '<p:presentationPr ' + XMLNS + '/>');
      z.file("ppt/viewProps.xml", HEAD + '<p:viewPr ' + XMLNS + '><p:normalViewPr><p:restoredLeft sz="15620"/><p:restoredTop sz="94660"/></p:normalViewPr><p:gridSpacing cx="76200" cy="76200"/></p:viewPr>');
      z.file("ppt/tableStyles.xml", HEAD + '<a:tblStyleLst xmlns:a="' + NSA + '" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>');
      z.file("ppt/theme/theme1.xml", THEME);
      z.file("ppt/slideMasters/slideMaster1.xml", MASTER);
      z.file("ppt/slideMasters/_rels/slideMaster1.xml.rels", HEAD + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="' + NSR + '/slideLayout" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rId2" Type="' + NSR + '/theme" Target="../theme/theme1.xml"/></Relationships>');
      z.file("ppt/slideLayouts/slideLayout1.xml", LAYOUT);
      z.file("ppt/slideLayouts/_rels/slideLayout1.xml.rels", HEAD + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="' + NSR + '/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>');
      var now = new Date().toISOString().replace(/\.\d+Z$/, "Z");
      z.file("docProps/core.xml", HEAD + '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>' + x(title) + '</dc:title><dc:creator>Homework Launchpad</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">' + now + '</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">' + now + "</dcterms:modified></cp:coreProperties>");
      z.file("docProps/app.xml", HEAD + '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Launchpad Present</Application><Slides>' + n + "</Slides></Properties>");
      z.file("_rels/.rels", HEAD + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="' + NSR + '/officeDocument" Target="ppt/presentation.xml"/>' +
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="' + NSR + '/extended-properties" Target="docProps/app.xml"/></Relationships>');
      z.file("[Content_Types].xml", HEAD + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' +
        '<Default Extension="png" ContentType="image/png"/><Default Extension="jpeg" ContentType="image/jpeg"/><Default Extension="gif" ContentType="image/gif"/>' +
        '<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>' +
        '<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>' +
        '<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>' +
        '<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>' +
        '<Override PartName="/ppt/presProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presProps+xml"/>' +
        '<Override PartName="/ppt/viewProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml"/>' +
        '<Override PartName="/ppt/tableStyles.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml"/>' +
        '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' +
        '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>' + ctOver + "</Types>");
      return z.generateAsync({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.presentationml.presentation", compression: "DEFLATE" });
    });
  }

  /* ---------- reading: text, pictures and shapes, where they sit ---------- */
  function q(el, ns, name) { return el ? Array.prototype.slice.call(el.getElementsByTagNameNS(ns, name)) : []; }
  function kids(el) { var o = []; for (var n = el && el.firstChild; n; n = n.nextSibling) if (n.nodeType === 1) o.push(n); return o; }
  function esc(s) { return String(s).replace(/[&<>"]/g, function (c) { return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]; }); }
  function read(file) {
    return JSZip.loadAsync(file).then(function (z) {
      var P = new DOMParser(), xml = function (s) { return P.parseFromString(s, "application/xml"); };
      var get = function (p) { var f = z.file(p); return f ? f.async("string") : Promise.resolve(null); };
      return Promise.all([get("ppt/presentation.xml"), get("ppt/_rels/presentation.xml.rels")]).then(function (a) {
        if (!a[0]) throw new Error("this doesn't look like a PowerPoint file");
        var pres = xml(a[0]), rels = {};
        Array.prototype.forEach.call(xml(a[1] || "<r/>").getElementsByTagName("Relationship"), function (r) { rels[r.getAttribute("Id")] = r.getAttribute("Target"); });
        var sz = q(pres, NSP, "sldSz")[0], k = sz ? 960 / +sz.getAttribute("cx") : 1 / EMU;
        var ids = q(pres, NSP, "sldId").slice(0, 200);
        return Promise.all(ids.map(function (sid) {
          var path = "ppt/" + String(rels[sid.getAttributeNS(NSR, "id")] || "").replace(/^\/?ppt\//, "");
          var rp = path.replace(/slides\/(slide\d+\.xml)$/, "slides/_rels/$1.rels");
          return Promise.all([get(path), get(rp)]).then(function (b) { return slide(z, b[0], b[1], k); });
        })).then(function (slides) {
          slides = slides.filter(Boolean); if (!slides.length) throw new Error("there were no slides in it");
          return { theme: "classic", slides: slides };
        });
      });
    });
    function slide(z, body, relsXml, k) {
      if (!body) return null;
      var d = new DOMParser().parseFromString(body, "application/xml"), rels = {};
      Array.prototype.forEach.call(new DOMParser().parseFromString(relsXml || "<r/>", "application/xml").getElementsByTagName("Relationship"), function (r) { rels[r.getAttribute("Id")] = r.getAttribute("Target"); });
      var els = [], jobs = [], id = 0, uid = function () { return "i" + (++id) + Math.random().toString(36).slice(2, 7); };
      var DEF = { title: [60, 36, 840, 84, 40], ctrTitle: [80, 160, 800, 120, 54], subTitle: [120, 295, 720, 70, 24], body: [60, 135, 840, 370, 24] };
      function box(sp) {
        var off = q(sp, NSA, "off")[0], ext = q(sp, NSA, "ext")[0];
        if (!off || !ext) return null;
        return { x: Math.round(+off.getAttribute("x") * k), y: Math.round(+off.getAttribute("y") * k), w: Math.round(+ext.getAttribute("cx") * k), h: Math.round(+ext.getAttribute("cy") * k) };
      }
      q(d, NSP, "sp").forEach(function (sp) {
        var ph = q(sp, NSP, "ph")[0], phType = ph ? (ph.getAttribute("type") || "body") : null;
        var b = box(sp), tb = q(sp, NSP, "txBody")[0];
        if (!b && phType) { var df = DEF[phType] || DEF.body; b = { x: df[0], y: df[1], w: df[2], h: df[3] }; }
        if (!b) return;
        var html = "", size = 0, inList = false;
        if (tb) q(tb, NSA, "p").forEach(function (p) {
          var line = "", bullet = q(p, NSA, "buChar").length || q(p, NSA, "buAutoNum").length || (phType === "body" && !q(p, NSA, "buNone").length);
          kids(p).forEach(function (n) {
            if (n.localName === "br") { line += "<br>"; return; }
            if (n.localName !== "r" && n.localName !== "fld") return;
            var rp = q(n, NSA, "rPr")[0], t = esc((q(n, NSA, "t")[0] || {}).textContent || "");
            if (rp) {
              if (+rp.getAttribute("sz")) size = size || +rp.getAttribute("sz") / 100;
              if (rp.getAttribute("b") === "1") t = "<b>" + t + "</b>"; if (rp.getAttribute("i") === "1") t = "<i>" + t + "</i>";
              if (rp.getAttribute("u") && rp.getAttribute("u") !== "none") t = "<u>" + t + "</u>";
              var c = q(rp, NSA, "srgbClr")[0]; if (c && /^[0-9a-f]{6}$/i.test(c.getAttribute("val"))) t = '<span style="color:#' + c.getAttribute("val") + '">' + t + "</span>";
            }
            line += t;
          });
          if (bullet && line) { if (!inList) { html += "<ul>"; inList = true; } html += "<li>" + line + "</li>"; }
          else { if (inList) { html += "</ul>"; inList = false; } html += "<div>" + (line || "<br>") + "</div>"; }
        });
        if (inList) html += "</ul>";
        var hasText = html.replace(/<[^>]+>/g, "").trim();
        if (hasText) {
          var title = phType === "title" || phType === "ctrTitle";
          els.push({ id: uid(), type: "text", x: b.x, y: b.y, w: b.w, h: b.h, size: size || (DEF[phType] || [0, 0, 0, 0, 24])[4], role: title ? "title" : "body", html: html, align: phType === "ctrTitle" || phType === "subTitle" ? "center" : "left" });
          return;
        }
        var geom = q(sp, NSA, "prstGeom")[0], fill = q(q(sp, NSP, "spPr")[0], NSA, "srgbClr")[0], map = { rect: "rect", roundRect: "round", ellipse: "ellipse", triangle: "triangle", star5: "star", rightArrow: "arrow", line: "line" };
        if (geom && map[geom.getAttribute("prst")] && fill) els.push({ id: uid(), type: "shape", shape: map[geom.getAttribute("prst")], x: b.x, y: b.y, w: Math.max(b.w, 4), h: Math.max(b.h, 4), color: "#" + fill.getAttribute("val") });
      });
      q(d, NSP, "pic").forEach(function (pic) {
        var b = box(pic), bl = q(pic, NSA, "blip")[0], rid = bl && bl.getAttributeNS(NSR, "embed"), t = rid && rels[rid];
        if (!b || !t) return;
        var path = "ppt/" + t.replace(/^\.\.\//, ""), ext = (path.match(/\.(png|jpe?g|gif)$/i) || [])[1], f = z.file(path);
        if (!f || !ext) return;
        var el = { id: uid(), type: "img", x: b.x, y: b.y, w: b.w, h: b.h, src: "" }; els.push(el);
        jobs.push(f.async("base64").then(function (s) { el.src = "data:image/" + (/png/i.test(ext) ? "png" : /gif/i.test(ext) ? "gif" : "jpeg") + ";base64," + s; }));
      });
      return Promise.all(jobs).then(function () { return { id: uid(), layout: "blank", trans: "fade", notes: "", els: els }; });
    }
  }

  window.Pptx = { build: build, read: read };
})();
