#!/usr/bin/env node
// 🌗 Runs web/dark.js on a .pptx file, outside PowerPoint: to look at a whole deck turned dark
// (render the result with LibreOffice) and to prove that turning it back gives the same slides.
//
//   node tools/dark-deck.js IN.pptx OUT.pptx [--slides 3,7-9]   toggles the slides, tags in OUT.pptx.tags.json
//   node tools/dark-deck.js IN.pptx --check [--slides …]         toggles twice, compares with the original
//
// In PowerPoint the tag of a slide lives in the slide (Slide.tags); here it lives in the side file.
const fs = require("fs");
const JSZip = require("jszip");
const UPNG = require("upng-js");
const jpeg = require("jpeg-js");
const { DOMParser, XMLSerializer } = require("@xmldom/xmldom");
const Dark = require("../web/dark.js");

const io = {
  parse: (s) => new DOMParser().parseFromString(s, "application/xml"),
  serialize: (n) => new XMLSerializer().serializeToString(n),
  decodeImage: async (bytes, ext) => {
    if (ext === "png") {
      const img = UPNG.decode(bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength));
      return { rgba: new Uint8Array(UPNG.toRGBA8(img)[0]), width: img.width, height: img.height };
    }
    const img = jpeg.decode(bytes, { useTArray: true, maxMemoryUsageInMB: 2048 });
    return { rgba: img.data, width: img.width, height: img.height };
  },
  encodePng: (rgba, w, h) => new Uint8Array(UPNG.encode([rgba.buffer.slice(rgba.byteOffset, rgba.byteOffset + rgba.byteLength)], w, h, 0)),
};

function slidePaths(zip, pick) {
  // In presentation order: presentation.xml lists the rIds, its rels name the files.
  return (async () => {
    const pres = io.parse(await zip.file("ppt/presentation.xml").async("string"));
    const rels = io.parse(await zip.file("ppt/_rels/presentation.xml.rels").async("string"));
    const target = new Map([...rels.getElementsByTagName("Relationship")].map((r) => [r.getAttribute("Id"), r.getAttribute("Target")]));
    const all = [...pres.getElementsByTagName("p:sldId")].map((s) => `ppt/${target.get(s.getAttribute("r:id")).replace(/^\/?(ppt\/)?/, "")}`);
    return all.map((path, i) => ({ path, number: i + 1 })).filter((s) => !pick || pick.has(s.number));
  })();
}

function parseSlides(text) {
  const set = new Set();
  for (const part of text.split(",")) {
    const [a, b] = part.split("-").map(Number);
    for (let n = a; n <= (b || a); n++) set.add(n);
  }
  return set;
}

async function toggle(zip, slides, tags) {
  const report = [];
  for (const s of slides) {
    try {
      const r = await Dark.invertSlideInPackage(zip, s.path, tags[s.path] || null, io);
      if (!r.skipped) { if (r.tag) tags[s.path] = r.tag; else delete tags[s.path]; }
      report.push({ ...s, skipped: r.skipped, toDark: r.toDark, changed: r.changed, kept: r.kept, pictures: r.pictures });
    } catch (e) {
      report.push({ ...s, error: e.message });
    }
  }
  return report;
}

// The pictures a slide shows, as pixels: [shape id, hash of the decoded RGBA].
async function pictureHashes(zip, path) {
  const crypto = require("crypto");
  const slide = await zip.file(path).async("string");
  const rels = io.parse(await zip.file(path.replace(/([^/]+)$/, "_rels/$1.rels")).async("string"));
  const target = new Map([...rels.getElementsByTagName("Relationship")].map((r) => [r.getAttribute("Id"), r.getAttribute("Target")]));
  const out = [];
  for (const m of slide.matchAll(/r:embed="([^"]+)"/g)) {
    const file = `ppt/${(target.get(m[1]) || "").replace("../", "")}`, ext = file.split(".").pop().toLowerCase();
    if (!zip.file(file)) continue;
    const bytes = await zip.file(file).async("uint8array");
    let pixels = bytes;
    if (["png", "jpg", "jpeg"].includes(ext)) { try { pixels = (await io.decodeImage(bytes, ext)).rgba; } catch { /* compare the bytes */ } }
    out.push(crypto.createHash("sha1").update(pixels).digest("hex"));
  }
  return out;
}

// The slide XML with the relationship ids of its pictures blanked: a flipped-back picture is a
// new part with a new id, and its pixels are compared separately.
const comparable = (xml) => io.serialize(io.parse(xml)).replace(/r:embed="[^"]+"/g, 'r:embed=""');

async function main() {
  const args = process.argv.slice(2);
  const input = args[0], check = args.includes("--check");
  const pickAt = args.indexOf("--slides");
  const pick = pickAt > 0 ? parseSlides(args[pickAt + 1]) : null;
  const zip = await JSZip.loadAsync(fs.readFileSync(input));
  const slides = await slidePaths(zip, pick);
  const tags = !check && fs.existsSync(`${input}.tags.json`) ? JSON.parse(fs.readFileSync(`${input}.tags.json`, "utf8")) : {};

  if (!check) {
    const report = await toggle(zip, slides, tags);
    fs.writeFileSync(args[1], await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }));
    fs.writeFileSync(`${args[1]}.tags.json`, JSON.stringify(tags));
    for (const r of report) console.log(`${String(r.number).padStart(4)} ${r.error ? `ERROR ${r.error}` : r.skipped ? `skipped: ${r.skipped}` : `${r.toDark ? "→ dark " : "→ light"} ${r.changed} changed, ${r.kept} kept, ${r.pictures} pictures`}`);
    return;
  }

  const before = new Map();
  for (const s of slides) before.set(s.path, { xml: comparable(await zip.file(s.path).async("string")), pics: await pictureHashes(zip, s.path) });
  const there = await toggle(zip, slides, tags);
  const mid = new Map();
  for (const s of slides) mid.set(s.path, comparable(await zip.file(s.path).async("string")));
  const back = await toggle(zip, slides, tags);
  let bad = 0, skipped = 0, pictures = 0, changed = 0;
  for (const [i, s] of slides.entries()) {
    if (there[i].skipped) { skipped++; continue; }
    const problems = [];
    if (there[i].error || back[i].error) problems.push(there[i].error || back[i].error);
    else {
      if (comparable(await zip.file(s.path).async("string")) !== before.get(s.path).xml) problems.push("XML differs after the round trip");
      if (JSON.stringify(await pictureHashes(zip, s.path)) !== JSON.stringify(before.get(s.path).pics)) problems.push("picture pixels differ after the round trip");
      if (tags[s.path]) problems.push("a tag is left on the restored slide");
      if (mid.get(s.path) === before.get(s.path).xml) problems.push("nothing changed on the way there");
    }
    pictures += there[i].pictures || 0; changed += there[i].changed || 0;
    if (problems.length) { bad++; console.log(`  slide ${s.number}: ${problems.join("; ")}`); }
  }
  console.log(`${bad ? "FAIL" : "OK"} ${input}: ${slides.length} slides, ${bad} not identical, ${skipped} skipped, ${changed} colors and ${pictures} pictures changed and restored`);
  process.exitCode = bad ? 1 : 0;
}

main().catch((e) => { console.error(e); process.exit(2); });
