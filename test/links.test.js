// node --test — runs web/links.js on hand-written runs, outside PowerPoint.
const test = require("node:test");
const assert = require("node:assert/strict");
const { DOMParser, XMLSerializer } = require("@xmldom/xmldom");
const { underlineLinks } = require("../web/links.js");

const NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
  + 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';

// What PowerPoint writes after Font › Underline color on a link, copied from a hand-made slide.
const HLINK_TX = '<a:extLst><a:ext uri="{A12FA001-AC4F-418D-AE19-62706E023703}">'
  + '<ahyp:hlinkClr xmlns:ahyp="http://schemas.microsoft.com/office/drawing/2018/hyperlinkcolor" val="tx"/></a:ext></a:extLst>';
const DONE = '<a:rPr lang="en-US" sz="2400" u="sng" noProof="1"><a:uLnTx/>'
  + '<a:uFill><a:solidFill><a:schemeClr val="tx2"/></a:solidFill></a:uFill>'
  + `<a:hlinkClick r:id="rId4">${HLINK_TX}</a:hlinkClick></a:rPr>`;

const run = (rPr) => `<a:r>${rPr}<a:t>link</a:t></a:r>`;
const body = (...runs) => `<p:txBody xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" ${NS}><a:p>${runs.join("")}</a:p></p:txBody>`;

function apply(xml) {
  const doc = new DOMParser().parseFromString(xml, "application/xml");
  const changed = underlineLinks(doc);
  return { changed, xml: new XMLSerializer().serializeToString(doc) };
}

// The rPr of the first run, namespace declarations dropped, for comparing against DONE.
const firstRPr = (xml) => xml.match(/<a:rPr[\s\S]*?<\/a:rPr>/)[0];

test("a plain link gets exactly the hand-made formatting", () => {
  // u="none" only to pin the attribute order; XML doesn't care, the string compare does.
  const { changed, xml } = apply(body(run('<a:rPr lang="en-US" sz="2400" u="none" noProof="1"><a:hlinkClick r:id="rId4"/></a:rPr>')));
  assert.equal(changed, 1);
  assert.equal(firstRPr(xml), DONE);
});

test("a link already done is left alone", () => {
  const xml = body(run(DONE));
  const out = apply(xml);
  assert.equal(out.changed, 0);
  assert.equal(firstRPr(out.xml), DONE);
});

test("text color and font stay; the underline goes after the fill, before the font", () => {
  const rPr = '<a:rPr lang="en-US" dirty="0"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>'
    + '<a:highlight><a:srgbClr val="FFEBEB"/></a:highlight>'
    + '<a:latin typeface="Calibri"/><a:cs typeface="Calibri"/><a:hlinkClick r:id="rId2"/></a:rPr>';
  const { changed, xml } = apply(body(run(rPr)));
  assert.equal(changed, 1);
  const names = [...firstRPr(xml).matchAll(/<a:(\w+)/g)].map((m) => m[1]).filter((n) => n !== "rPr");
  assert.deepEqual(names.slice(0, 9),
    ["solidFill", "srgbClr", "highlight", "srgbClr", "uLnTx", "uFill", "solidFill", "schemeClr", "latin"]);
  assert.match(xml, /<a:srgbClr val="FF0000"\/>/);
  assert.match(xml, /<a:latin typeface="Calibri"\/>/);
});

test("an old underline style is replaced, not duplicated", () => {
  const rPr = '<a:rPr u="dbl"><a:uLnTx/><a:uFillTx/><a:hlinkClick r:id="rId2"/></a:rPr>';
  const { xml } = apply(body(run(rPr)));
  assert.equal((xml.match(/<a:uLnTx\/>/g) || []).length, 1);
  assert.doesNotMatch(xml, /uFillTx/);
  assert.match(xml, /u="sng"/);
});

test("runs without a link are untouched", () => {
  const plain = '<a:rPr lang="en-US" b="1"><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill></a:rPr>';
  const { changed, xml } = apply(body(run(plain)));
  assert.equal(changed, 0);
  assert.equal(firstRPr(xml), plain);
});
