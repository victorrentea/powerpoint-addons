// node --test — runs web/regroup.js on a hand-written slide, outside PowerPoint.
const test = require("node:test");
const assert = require("node:assert/strict");
const { DOMParser, XMLSerializer } = require("@xmldom/xmldom");
const { addToGroup } = require("../web/regroup.js");

const NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
  + 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';

const sp = (id, name, x, y, w, h) => `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>`
  + `<p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${w}" cy="${h}"/></a:xfrm></p:spPr></p:sp>`;

// One click step per shape id; `with` adds a with-previous effect to the last click step.
function timing(clicks, withPrevious) {
  let n = 3;
  const effect = (id, type) => `<p:par><p:cTn id="${n++}" presetID="1" presetClass="entr" nodeType="${type}"><p:childTnLst>`
    + `<p:set><p:cBhvr><p:cTn id="${n++}" dur="1"/><p:tgtEl><p:spTgt spid="${id}"/></p:tgtEl></p:cBhvr></p:set></p:childTnLst></p:cTn></p:par>`;
  const steps = clicks.map((id, i) => {
    const head = `<p:par><p:cTn id="${n++}"><p:childTnLst><p:par><p:cTn id="${n++}"><p:childTnLst>`;
    const body = effect(id, "clickEffect") + (i === clicks.length - 1 && withPrevious ? effect(withPrevious, "withEffect") : "");
    return head + body + "</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>";
  }).join("");
  return `<p:timing><p:tnLst><p:par><p:cTn id="1" nodeType="tmRoot"><p:childTnLst><p:seq><p:cTn id="2" nodeType="mainSeq"><p:childTnLst>`
    + steps + `</p:childTnLst></p:cTn></p:seq></p:childTnLst></p:cTn></p:par></p:tnLst>`
    + `<p:bldLst>${clicks.concat(withPrevious || []).map((id) => `<p:bldP spid="${id}" grpId="0"/>`).join("")}</p:bldLst></p:timing>`;
}

// Group 10 sits at (1000,1000) 2000×1000 on the slide, its children in a space scaled ×2.
function slide({ shapes, anim = "" }) {
  return `<p:sld ${NS}><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>`
    + shapes + `</p:spTree></p:cSld>${anim}</p:sld>`;
}
const group = `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="10" name="Group 9"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>`
  + `<p:grpSpPr><a:xfrm><a:off x="1000" y="1000"/><a:ext cx="2000" cy="1000"/><a:chOff x="0" y="0"/><a:chExt cx="4000" cy="2000"/></a:xfrm></p:grpSpPr>`
  + sp(11, "Inside", 0, 0, 4000, 2000) + `</p:grpSp>`;

// Marks the shapes with these ids the way the pane does, then runs addToGroup.
function run(xml, ids) {
  const doc = new DOMParser().parseFromString(xml, "application/xml");
  const marks = new Map();
  for (const pr of [...doc.getElementsByTagName("p:cNvPr")]) {
    if (ids.includes(pr.getAttribute("id"))) {
      marks.set(`mark-${pr.getAttribute("id")}`, pr.getAttribute("name"));
      pr.setAttribute("name", `mark-${pr.getAttribute("id")}`);
    }
  }
  const result = addToGroup(doc, marks);
  return { result, doc, xml: new XMLSerializer().serializeToString(doc) };
}

const attrs = (doc, tag, id) => {
  const pr = [...doc.getElementsByTagName("p:cNvPr")].find((c) => c.getAttribute("id") === id);
  const el = pr.parentNode.parentNode;
  const f = el.getElementsByTagName("a:xfrm")[0];
  const pt = (n, a, b) => { const e = f.getElementsByTagName(n)[0]; return e && [Number(e.getAttribute(a)), Number(e.getAttribute(b))]; };
  return { parent: el.parentNode.localName, off: pt("a:off", "x", "y"), ext: pt("a:ext", "cx", "cy"),
    chOff: pt("a:chOff", "x", "y"), chExt: pt("a:chExt", "cx", "cy") };
};

test("moves the shape into the animated group, in the group's scaled space, and grows the group", () => {
  const { result, doc } = run(slide({ shapes: group + sp(20, "Newcomer", 3500, 1500, 500, 500), anim: timing(["10"]) }), ["10", "20"]);
  assert.deepEqual(result, { group: "Group 9", animated: true, added: ["Newcomer"], droppedAnimations: 0 });
  const moved = attrs(doc, "p:sp", "20");
  assert.equal(moved.parent, "grpSp");
  assert.deepEqual(moved.off, [5000, 1000]); // (3500-1000)×2, (1500-1000)×2
  assert.deepEqual(moved.ext, [1000, 1000]);
  const g = attrs(doc, "p:grpSp", "10");
  assert.deepEqual(g.off, [1000, 1000]);
  assert.deepEqual(g.ext, [3000, 1000]); // now reaches x=4000
  assert.deepEqual(g.chOff, [0, 0]);
  assert.deepEqual(g.chExt, [6000, 2000]); // same ×2 scale, so "Inside" doesn't move
});

test("the group keeps its animation; the newcomer's own effect and build entry go", () => {
  const { xml } = run(slide({ shapes: group + sp(20, "Newcomer", 0, 0, 100, 100) + sp(30, "Other", 0, 0, 100, 100),
    anim: timing(["30", "10"], "20") }), ["10", "20"]);
  assert.deepEqual([...xml.matchAll(/spTgt spid="(\d+)"/g)].map((m) => m[1]), ["30", "10"]);
  assert.deepEqual([...xml.matchAll(/bldP spid="(\d+)"/g)].map((m) => m[1]), ["30", "10"]);
  const ids = [...xml.matchAll(/cTn id="(\d+)"/g)].map((m) => Number(m[1]));
  assert.deepEqual(ids, ids.map((_, i) => i + 1), "cTn ids renumbered densely");
});

test("a click step left empty disappears with its effect", () => {
  const { xml } = run(slide({ shapes: group + sp(20, "Newcomer", 0, 0, 100, 100), anim: timing(["20", "10"]) }), ["10", "20"]);
  assert.deepEqual([...xml.matchAll(/spTgt spid="(\d+)"/g)].map((m) => m[1]), ["10"]);
  assert.equal([...xml.matchAll(/nodeType="clickEffect"/g)].length, 1);
});

test("restores the original names", () => {
  const { xml } = run(slide({ shapes: group + sp(20, "Newcomer", 0, 0, 100, 100), anim: timing(["10"]) }), ["10", "20"]);
  assert.match(xml, /name="Group 9"/);
  assert.match(xml, /name="Newcomer"/);
  assert.doesNotMatch(xml, /mark-/);
});

test("a shape below the group goes to the bottom of the group, one above to the top", () => {
  const { doc } = run(slide({ shapes: sp(20, "Below", 0, 0, 100, 100) + group + sp(30, "Above", 0, 0, 100, 100), anim: timing(["10"]) }),
    ["10", "20", "30"]);
  const g = doc.getElementsByTagName("p:grpSp")[0];
  const order = [...g.childNodes].filter((n) => n.localName === "sp").map((n) => n.getElementsByTagName("p:cNvPr")[0].getAttribute("name"));
  assert.deepEqual(order, ["Below", "Inside", "Above"]);
});

test("an all-zero group frame counts as unscaled, sized by its children", () => {
  const zero = group.replace(/<a:xfrm>.*?<\/a:xfrm>/, '<a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm>');
  const { doc } = run(slide({ shapes: zero + sp(20, "Newcomer", 5000, 0, 1000, 1000), anim: timing(["10"]) }), ["10", "20"]);
  assert.deepEqual(attrs(doc, "p:sp", "20").off, [5000, 0]);
  const g = attrs(doc, "p:grpSp", "10");
  assert.deepEqual([g.off, g.ext, g.chOff, g.chExt], [[0, 0], [6000, 2000], [0, 0], [6000, 2000]]);
});

test("removes the timing altogether when the newcomer had the only effect", () => {
  const { xml } = run(slide({ shapes: group + sp(20, "Newcomer", 0, 0, 100, 100), anim: timing(["20"]) }), ["10", "20"]);
  assert.doesNotMatch(xml, /p:timing/);
});

test("refuses what PowerPoint can't group or what's ambiguous", () => {
  const group2 = group.replace('id="10" name="Group 9"', 'id="12" name="Group 11"').replace('id="11"', 'id="13"');
  const cases = [
    [slide({ shapes: group + group2, anim: timing(["10", "12"]) }), ["10", "12"], /2 animated groups/],
    [slide({ shapes: sp(20, "A", 0, 0, 1, 1) + sp(30, "B", 0, 0, 1, 1) }), ["20", "30"], /No group/],
    [slide({ shapes: group }), ["10"], /Select the animated group and the shapes/],
    [slide({ shapes: group + sp(20, "Title", 0, 0, 1, 1).replace("<p:nvPr/>", '<p:nvPr><p:ph type="title"/></p:nvPr>') }), ["10", "20"], /placeholder/],
    [slide({ shapes: group.replace("<a:xfrm>", '<a:xfrm rot="60000">') + sp(20, "A", 0, 0, 1, 1) }), ["10", "20"], /rotated/],
  ];
  for (const [xml, ids, error] of cases) assert.throws(() => run(xml, ids), error);
});
