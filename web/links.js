// 🔗 Restyles the hyperlinks of a slide: the text keeps its own color and font, turns bold, and
// its underline turns blue — theme Text 2 on a light slide, a lighter blue (LIGHT_BLUE) on a dark
// one, where a navy line would sink into the page. PowerPoint paints a link in the theme's hyperlink color and Office.js
// has no underline color, so the pane exports the slide and this edits each linked run's
// <a:rPr> exactly the way PowerPoint's own dialog does: b="1", u="sng", <a:uLnTx/>, an <a:uFill>,
// and the 2018 hlinkClr extension with val="tx" ("use the text color, not the link color").
// Pure DOM, no Office.js, so test/links.test.js runs it in Node.
(function (root) {
  const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
  const AHYP = "http://schemas.microsoft.com/office/drawing/2018/hyperlinkcolor";
  const HLINK_COLOR_EXT = "{A12FA001-AC4F-418D-AE19-62706E023703}";
  const LIGHT_BLUE = "58A6FF"; // 7.5:1 on black: reads as blue, not as one more white line
  const UNDERLINE = new Set(["uLnTx", "uLn", "uFillTx", "uFill"]);
  // <a:rPr> children that come after the underline ones (CT_TextCharacterProperties order).
  const AFTER_UNDERLINE = new Set(["latin", "ea", "cs", "sym", "hlinkClick", "hlinkMouseOver", "rtl", "extLst"]);

  const elements = (node) => [...node.childNodes].filter((n) => n.nodeType === 1);
  const child = (node, ns, name) => elements(node).find((n) => n.namespaceURI === ns && n.localName === name);

  function make(doc, ns, qname, attrs = {}, children = []) {
    const el = doc.createElementNS(ns, qname);
    for (const [k, v] of Object.entries(attrs)) el.setAttribute(k, v);
    for (const c of children) el.appendChild(c);
    return el;
  }

  const colorOf = (dark) => (dark ? ["srgbClr", LIGHT_BLUE] : ["schemeClr", "tx2"]);

  function underline(doc, dark) {
    const [tag, val] = colorOf(dark);
    const color = make(doc, A, `a:${tag}`, { val });
    return [make(doc, A, "a:uLnTx"), make(doc, A, "a:uFill", {}, [make(doc, A, "a:solidFill", {}, [color])])];
  }

  // Makes the link take its run's text color: <a:extLst><a:ext uri=…><ahyp:hlinkClr val="tx"/>.
  function useTextColor(doc, hlink) {
    let extLst = child(hlink, A, "extLst");
    if (!extLst) extLst = hlink.appendChild(make(doc, A, "a:extLst"));
    let ext = elements(extLst).find((e) => e.localName === "ext" && e.getAttribute("uri") === HLINK_COLOR_EXT);
    if (!ext) ext = extLst.appendChild(make(doc, A, "a:ext", { uri: HLINK_COLOR_EXT }));
    let clr = child(ext, AHYP, "hlinkClr");
    if (!clr) {
      clr = ext.appendChild(make(doc, AHYP, "ahyp:hlinkClr"));
      clr.setAttributeNS("http://www.w3.org/2000/xmlns/", "xmlns:ahyp", AHYP);
    }
    clr.setAttribute("val", "tx");
  }

  function linkTextColor(hlink) {
    const ext = elements(child(hlink, A, "extLst") || hlink).find((e) => e.getAttribute("uri") === HLINK_COLOR_EXT);
    const clr = ext && child(ext, AHYP, "hlinkClr");
    return clr && clr.getAttribute("val");
  }

  function isDone(rPr, hlink, dark) {
    const [tag, val] = colorOf(dark);
    const fill = child(rPr, A, "uFill");
    const solid = fill && child(fill, A, "solidFill");
    const color = solid && child(solid, A, tag);
    return rPr.getAttribute("b") === "1" && rPr.getAttribute("u") === "sng" && child(rPr, A, "uLnTx")
      && color && color.getAttribute("val").toUpperCase() === val.toUpperCase() && linkTextColor(hlink) === "tx";
  }

  // Every linked run (and field) on the slide turns bold with a blue underline; returns how many changed.
  // `dark`: the slide's page is dark, so the underline takes the lighter blue.
  function underlineLinks(doc, { dark = false } = {}) {
    let changed = 0;
    for (const rPr of [...doc.getElementsByTagNameNS(A, "rPr")]) {
      const hlink = child(rPr, A, "hlinkClick");
      if (!hlink || isDone(rPr, hlink, dark)) continue;
      rPr.setAttribute("b", "1");
      rPr.setAttribute("u", "sng");
      elements(rPr).filter((e) => e.namespaceURI === A && UNDERLINE.has(e.localName)).forEach((e) => rPr.removeChild(e));
      const next = elements(rPr).find((e) => e.namespaceURI === A && AFTER_UNDERLINE.has(e.localName)) || null;
      for (const el of underline(doc, dark)) rPr.insertBefore(el, next);
      useTextColor(doc, hlink);
      changed++;
    }
    return changed;
  }

  const api = { underlineLinks, LIGHT_BLUE };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.Links = api;
})(this);
