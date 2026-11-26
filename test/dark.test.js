// node --test — runs web/dark.js on hand-written slides, outside PowerPoint.
const test = require("node:test");
const assert = require("node:assert/strict");
const { DOMParser, XMLSerializer } = require("@xmldom/xmldom");
const Dark = require("../web/dark.js");

const NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
  + 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
  + 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
const parse = (s) => new DOMParser().parseFromString(s, "application/xml");
const serialize = (n) => new XMLSerializer().serializeToString(n);

const THEME = parse(`<a:theme ${NS}><a:themeElements><a:clrScheme name="t">`
  + '<a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>'
  + '<a:dk2><a:srgbClr val="061DC8"/></a:dk2><a:lt2><a:srgbClr val="727272"/></a:lt2>'
  + '<a:accent1><a:srgbClr val="7889FB"/></a:accent1><a:accent2><a:srgbClr val="C7CDFD"/></a:accent2>'
  + '<a:accent3><a:srgbClr val="FFFFFF"/></a:accent3><a:accent4><a:srgbClr val="000000"/></a:accent4>'
  + '<a:accent5><a:srgbClr val="BEC4FD"/></a:accent5><a:accent6><a:srgbClr val="B4BAE5"/></a:accent6>'
  + '<a:hlink><a:srgbClr val="0432FF"/></a:hlink><a:folHlink><a:srgbClr val="0432FF"/></a:folHlink>'
  + "</a:clrScheme></a:themeElements></a:theme>");
const MASTER = parse(`<p:sldMaster ${NS}><p:cSld><p:bg><p:bgPr><a:solidFill><a:schemeClr val="bg1"/></a:solidFill></p:bgPr></p:bg><p:spTree/></p:cSld>`
  + '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/></p:sldMaster>');
const LAYOUT = parse(`<p:sldLayout ${NS}><p:cSld><p:spTree/></p:cSld></p:sldLayout>`);

const box = (id, spPr, runs, style = "") => `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="s${id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>`
  + `<p:spPr><a:xfrm><a:off x="${id * 1000}" y="0"/><a:ext cx="900" cy="900"/></a:xfrm>${spPr}</p:spPr>${style}`
  + `<p:txBody><a:bodyPr/><a:lstStyle/><a:p>${runs}</a:p></p:txBody></p:sp>`;
const run = (rPr, text = "x") => `<a:r>${rPr}<a:t>${text}</a:t></a:r>`;
const slideXml = (shapes, bg = "") => `<p:sld ${NS}><p:cSld>${bg}<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>${shapes}</p:spTree></p:cSld>`
  + "<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>";

function toggle(doc, tag) {
  return Dark.invertSlide(doc, { layout: LAYOUT, master: MASTER, theme: THEME, tag, parse, serialize });
}
const colors = (doc) => [...serialize(doc).matchAll(/srgbClr val="(\w+)"/g)].map((m) => m[1]);
const rgb = (hex) => [0, 2, 4].map((i) => parseInt(hex.slice(i, i + 2), 16));

const SAMPLE = slideXml(
  box(2, '<a:solidFill><a:srgbClr val="FFFF99"/></a:solidFill><a:ln><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln>',
    run('<a:rPr lang="en-US"/>') + run('<a:rPr><a:solidFill><a:srgbClr val="C00000"/></a:solidFill></a:rPr>'))
  + box(3, "<a:noFill/>", run('<a:rPr b="1"><a:highlight><a:srgbClr val="FFFF00"/></a:highlight></a:rPr>')
    + run('<a:rPr><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="50000"/><a:lumOff val="50000"/></a:schemeClr></a:solidFill></a:rPr>')
    + run('<a:rPr><a:solidFill><a:schemeClr val="bg1"/></a:solidFill></a:rPr>'),
  '<p:style><a:lnRef idx="0"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="0"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="dk1"/></a:fontRef></p:style>')
  + box(4, '<a:solidFill><a:srgbClr val="1E1E1E"/></a:solidFill>', run('<a:rPr><a:solidFill><a:srgbClr val="D4D4D4"/></a:solidFill></a:rPr>')));

test("flipping lightness twice gives every pixel back", () => {
  const px = new Uint8Array(4 * 4096);
  for (let i = 0; i < px.length; i++) px[i] = (i * 7919 + (i >> 3) * 31) & 255;
  const copy = px.slice();
  Dark.invertPixels(px);
  assert.notDeepEqual(px, copy);
  Dark.invertPixels(px);
  assert.deepEqual(px, copy);
  assert.deepEqual(Dark.flipLightness([255, 255, 255]), [0, 0, 0]);
  assert.deepEqual(Dark.flipLightness([200, 0, 0]), [255, 55, 55]); // dark red → light red, same hue
});

test("surfaces: white ↔ black, a pastel note becomes a deep tint, a light-gray panel stays visible", () => {
  assert.deepEqual(Dark.surface([255, 255, 255], true), [0, 0, 0]);
  assert.deepEqual(Dark.surface([0, 0, 0], false), [255, 255, 255]);
  const note = Dark.surface([255, 255, 153], true);
  assert.ok(Dark.contrast(note, [255, 255, 255]) >= 7 && note[0] > note[2] + 40, `deep gold: ${note}`);
  assert.deepEqual(Dark.surface([30, 30, 30], true), [30, 30, 30], "a dark surface stays dark");
  assert.ok(Dark.surface([191, 191, 191], true)[0] >= 0x22, "a gray panel does not sink into the page");
});

test("the yellow highlighter becomes an orange-amber that black text reads on", () => {
  const amber = Dark.marker([255, 255, 0], true);
  assert.ok(amber[0] > 240 && amber[1] > 140 && amber[1] < 200 && amber[2] < 90, `orange, not yellow, not brown: ${amber}`);
  assert.ok(Dark.contrast(amber, [0, 0, 0]) >= 9, `black reads on it: ${amber}`);
  assert.deepEqual(Dark.marker([204, 255, 204], true), Dark.surface([204, 255, 204], true), "a pastel is not a marker");
});

test("ink keeps its contrast: black turns white, a readable red stays, a dark red gets lighter", () => {
  const white = [255, 255, 255], black = [0, 0, 0];
  assert.deepEqual(Dark.ink(black, white, black), white);
  assert.deepEqual(Dark.ink([146, 208, 80], white, black), [146, 208, 80], "a bright green line is not touched");
  const red = Dark.ink([192, 0, 0], white, black);
  assert.ok(Dark.contrast(red, black) >= Dark.contrast([192, 0, 0], white) * 0.97, `${red}`);
  const faint = Dark.ink([191, 191, 191], white, black); // a faint gray stays faint
  assert.ok(Dark.contrast(faint, black) < 2.2, `${faint}`);
});

test("a slide turned dark and back is the same XML, and no tag is left", () => {
  const doc = parse(SAMPLE), before = serialize(doc);
  const there = toggle(doc, null);
  assert.equal(there.toDark, true);
  assert.notEqual(serialize(doc), before);
  assert.match(serialize(doc), /overrideClrMapping bg1="dk1" tx1="lt1"/);
  const back = toggle(doc, JSON.parse(JSON.stringify(there.tag)));
  assert.equal(serialize(doc), before);
  assert.equal(back.tag, null);
});

test("on the dark slide every text reads on what is behind it", () => {
  const doc = parse(SAMPLE);
  toggle(doc, null);
  const xml = serialize(doc);
  const fillOf = (id) => rgb(xml.match(new RegExp(`id="${id}".*?<p:spPr>.*?</a:xfrm><a:solidFill><a:srgbClr val="(\\w+)"`))[1]);
  const note = fillOf(2);
  assert.ok(Dark.luminance(note) < 0.18, "the pastel note is dark");
  assert.ok(Dark.contrast([255, 255, 255], note) >= 7, "white (inherited) text reads on it");
  const red = rgb(xml.match(/id="2".*?<a:rPr><a:solidFill><a:srgbClr val="(\w+)"/)[1]);
  assert.ok(Dark.contrast(red, note) >= 4.5, `the red text reads on the note: ${red}`);
  const outline = rgb(xml.match(/<a:ln><a:solidFill><a:srgbClr val="(\w+)"/)[1]);
  assert.ok(outline[0] === 255 && outline[1] < 90 && Dark.contrast(outline, [0, 0, 0]) >= 5.9, `the red outline stays red, a touch lighter: ${outline}`);
  assert.match(xml, /<a:highlight><a:srgbClr val="FFB32E"\/>/, "the yellow highlight is amber");
  assert.match(xml, /<a:rPr b="1"><a:solidFill><a:srgbClr val="0001\w\w"\/><\/a:solidFill><a:highlight>/, "and the text on it got a black of its own");
  assert.match(xml, /id="4".*?<a:srgbClr val="1E1E1E"\/>.*?<a:srgbClr val="D4D4D4"\/>/, "a dark code box stays as it is");
  assert.doesNotMatch(xml, /fontRef idx="minor"><a:schemeClr val="dk1"/, "text that was fixed black is not left black on black");
});

test("a color added while dark survives both ways", () => {
  const doc = parse(SAMPLE), before = serialize(doc);
  const there = toggle(doc, null);
  // The user paints the first run of shape 3 light green while the slide is dark.
  const rPr = [...doc.getElementsByTagName("a:rPr")].find((r) => r.getAttribute("b") === "1");
  const fill = doc.createElementNS(rPr.namespaceURI, "a:solidFill"), clr = doc.createElementNS(rPr.namespaceURI, "a:srgbClr");
  clr.setAttribute("val", "A0FFA0");
  rPr.insertBefore(fill, rPr.firstChild).appendChild(clr);
  const dark = serialize(doc);
  const back = toggle(doc, there.tag);
  assert.ok(back.tag && !back.tag.o && Object.values(back.tag.m).includes("#A0FFA0"), "the light slide remembers the added color");
  const green = colors(doc).find((c) => !colors(parse(before)).includes(c));
  assert.ok(Dark.contrast(rgb(green), [255, 255, 255]) > 3, `on white the green is dark enough to read: ${green}`);
  toggle(doc, back.tag);
  assert.equal(serialize(doc), dark, "dark again: the same dark slide, added color included");
});

test("a slide with a picture background is left alone", () => {
  const doc = parse(slideXml("", '<p:bg><p:bgPr><a:blipFill><a:blip r:embed="rId2"/></a:blipFill></p:bgPr></p:bg>'));
  const before = serialize(doc);
  assert.ok(toggle(doc, null).skipped);
  assert.equal(serialize(doc), before);
});

test("pictures: line art and light diagrams are flipped, photos and dark pictures are not", () => {
  const image = (w, h, paint) => { const px = new Uint8Array(w * h * 4); for (let y = 0; y < h; y++) for (let x = 0; x < w; x++) px.set(paint(x, y), (y * w + x) * 4); return Dark.imageStats(px, w, h); };
  const lineArt = image(100, 100, (x, y) => (Math.abs(x - y) < 6 ? [0, 0, 0, 255] : [0, 0, 0, 0]));
  const diagram = image(100, 100, (x, y) => (x > 20 && x < 60 && y > 30 && y < 50 ? [220, 40, 40, 255] : [255, 255, 255, 255]));
  const photo = image(100, 100, (x, y) => [(x * 13 + y * 7) % 256, (x * 5 + y * 11) % 256, (x * 3 + y * 17) % 256, 255]);
  const darkShot = image(100, 100, (x, y) => (y % 10 === 0 && x > 10 ? [212, 212, 212, 255] : [30, 30, 30, 255]));
  assert.equal(Dark.shouldInvert(lineArt, true), true);
  assert.equal(Dark.shouldInvert(diagram, true), true);
  assert.equal(Dark.shouldInvert(photo, true), false);
  assert.equal(Dark.shouldInvert(darkShot, true), false);
  assert.equal(Dark.shouldInvert(darkShot, false), true, "going light, the dark screenshot is the one to flip");
});
