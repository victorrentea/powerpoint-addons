// Colors per highlight.js token class, VS Code Dark+ / Light+ palettes.
const THEMES = {
  dark: {
    background: "#1E1E1E", text: "#D4D4D4",
    keyword: "#569CD6", built_in: "#4EC9B0", type: "#4EC9B0", title: "#DCDCAA",
    "title.class": "#4EC9B0", "title.function": "#DCDCAA", string: "#CE9178",
    number: "#B5CEA8", literal: "#569CD6", comment: "#6A9955", meta: "#C586C0",
    attr: "#9CDCFE", variable: "#9CDCFE", property: "#9CDCFE",
    tag: "#569CD6", name: "#569CD6", "attribute": "#9CDCFE", symbol: "#D7BA7D",
    regexp: "#D16969", subst: "#D4D4D4", doctag: "#608B4E",
  },
  light: {
    background: "#FFFFFF", text: "#1F1F1F",
    keyword: "#0000FF", built_in: "#267F99", type: "#267F99", title: "#795E26",
    "title.class": "#267F99", "title.function": "#795E26", string: "#A31515",
    number: "#098658", literal: "#0000FF", comment: "#008000", meta: "#AF00DB",
    attr: "#E50000", variable: "#001080", property: "#001080",
    tag: "#800000", name: "#800000", "attribute": "#E50000", symbol: "#811F3F",
    regexp: "#811F3F", subst: "#1F1F1F", doctag: "#008000",
  },
};

const $ = (id) => document.getElementById(id);

function status(msg, isError = false) {
  $("status").textContent = msg;
  $("status").className = isError ? "error" : "";
}

// Maps an element's "hljs-title class_" classes to the most specific palette key.
function colorFor(el, palette) {
  const kinds = [...el.classList]
    .filter((c) => c.startsWith("hljs-"))
    .map((c) => c.slice(5));
  const modifiers = [...el.classList].filter((c) => c.endsWith("_")).map((c) => c.slice(0, -1));
  for (const k of kinds) {
    for (const m of modifiers) if (palette[`${k}.${m}`]) return palette[`${k}.${m}`];
    if (palette[k]) return palette[k];
  }
  return null;
}

// Returns [{start, length, color}] runs over `code`, merging adjacent runs of the same color.
function colorRuns(code, language, palette) {
  const html = language === "plaintext"
    ? hljs.highlight(code, { language: "plaintext" }).value
    : hljs.highlight(code, { language, ignoreIllegals: true }).value;
  const root = new DOMParser().parseFromString(`<pre>${html}</pre>`, "text/html").body.firstChild;
  const runs = [];
  let offset = 0;
  (function walk(node, inherited) {
    for (const child of node.childNodes) {
      if (child.nodeType === Node.TEXT_NODE) {
        const length = child.textContent.length;
        if (inherited && length) {
          const last = runs[runs.length - 1];
          if (last && last.color === inherited && last.start + last.length === offset) last.length += length;
          else runs.push({ start: offset, length, color: inherited });
        }
        offset += length;
      } else {
        walk(child, colorFor(child, palette) || inherited);
      }
    }
  })(root, null);
  return runs;
}

async function insertCode() {
  const code = $("src").value.replace(/\r\n?/g, "\n").replace(/\t/g, "    ").replace(/\s+$/, "");
  if (!code) return status("Paste some code first.", true);
  const palette = THEMES[$("theme").value];
  const runs = colorRuns(code, $("lang").value, palette);
  const size = Number($("size").value) || 16;
  const lines = code.split("\n");
  const width = Math.min(900, Math.max(...lines.map((l) => l.length)) * size * 0.62 + 30);
  const height = lines.length * size * 1.25 + 20;

  await PowerPoint.run(async (ctx) => {
    const slide = ctx.presentation.getSelectedSlides().getItemAt(0);
    const shape = slide.shapes.addTextBox(code, { left: 40, top: 110, width, height });
    shape.name = "Code";
    shape.fill.setSolidColor(palette.background);
    shape.textFrame.wordWrap = false;
    shape.textFrame.autoSizeSetting = PowerPoint.ShapeAutoSize.autoSizeShapeToFitText;
    shape.textFrame.leftMargin = shape.textFrame.rightMargin = 12;
    shape.textFrame.topMargin = shape.textFrame.bottomMargin = 10;
    const all = shape.textFrame.textRange;
    all.font.name = "JetBrains Mono";
    all.font.size = size;
    all.font.color = palette.text;
    for (const r of runs) all.getSubstring(r.start, r.length).font.color = r.color;
    await ctx.sync();
  });
  status(`Inserted ${lines.length} lines, ${runs.length} colored runs.`);
}

const TEXT_SHAPES = new Set(["TextBox", "GeometricShape", "Placeholder"]);

async function scanDeck() {
  const ranking = await PowerPoint.run(async (ctx) => {
    const slides = ctx.presentation.slides;
    slides.load("items/id");
    await ctx.sync();
    const shapesPerSlide = slides.items.map((s) => s.shapes.load("items/type"));
    await ctx.sync();
    const texts = shapesPerSlide.map((shapes) =>
      shapes.items.filter((sh) => TEXT_SHAPES.has(sh.type)).map((sh) => sh.textFrame.textRange.load("text")));
    await ctx.sync();
    return texts.map((ranges, i) => {
      const all = ranges.map((r) => r.text).join(" ").trim();
      return { index: i + 1, words: all ? all.split(/\s+/).length : 0, title: all.split(/[\r\n]/)[0].slice(0, 50) };
    }).sort((a, b) => b.words - a.words);
  });
  const list = $("ranking");
  list.innerHTML = "";
  for (const s of ranking.slice(0, 15)) {
    const li = document.createElement("li");
    li.innerHTML = `<span class="n">${s.words} words</span> · slide ${s.index} <span class="t"></span>`;
    li.querySelector(".t").textContent = s.title;
    li.onclick = () => Office.context.document.goToByIdAsync(s.index, Office.GoToType.Index);
    list.appendChild(li);
  }
  status(`Scanned ${ranking.length} slides.`);
}

// Pixels per point when rendering the picture, so the negative keeps detail when the slide is projected.
const RENDER_SCALE = 3;

// Flips R, G and B of an RGBA buffer in place; alpha stays, so transparent areas remain transparent.
function invertPixels(rgba) {
  for (let i = 0; i < rgba.length; i += 4) {
    rgba[i] = 255 - rgba[i];
    rgba[i + 1] = 255 - rgba[i + 1];
    rgba[i + 2] = 255 - rgba[i + 2];
  }
  return rgba;
}

function invertPng(base64) {
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = () => {
      const canvas = document.createElement("canvas");
      canvas.width = img.naturalWidth;
      canvas.height = img.naturalHeight;
      const g = canvas.getContext("2d");
      g.drawImage(img, 0, 0);
      const pixels = g.getImageData(0, 0, canvas.width, canvas.height);
      invertPixels(pixels.data);
      g.putImageData(pixels, 0, 0);
      resolve(canvas.toDataURL("image/png").split(",")[1]);
    };
    img.onerror = () => reject(new Error("PowerPoint returned an image the pane can't decode."));
    img.src = `data:image/png;base64,${base64}`;
  });
}

// ShapeCollection.addPicture is still preview-only, so pictures go in through the Common API.
function insertPicture(base64, { left, top, width, height }) {
  return new Promise((resolve, reject) => {
    Office.context.document.setSelectedDataAsync(base64, {
      coercionType: Office.CoercionType.Image,
      imageLeft: left, imageTop: top, imageWidth: width, imageHeight: height,
    }, (r) => (r.status === Office.AsyncResultStatus.Failed ? reject(new Error(r.error.message)) : resolve()));
  });
}

// The selected picture, rendered at RENDER_SCALE: its id, name, frame and PNG.
async function selectedPicture(action) {
  if (!Office.context.requirements.isSetSupported("PowerPointApi", "1.10")) {
    throw new Error(`${action} needs PowerPointApi 1.10 (PowerPoint for Mac 16.105 or newer).`);
  }
  return PowerPoint.run(async (ctx) => {
    const selected = ctx.presentation.getSelectedShapes();
    selected.load("items/id,items/type,items/name,items/left,items/top,items/width,items/height");
    await ctx.sync();
    const shape = selected.items.find((s) => s.type === PowerPoint.ShapeType.image);
    if (!shape) throw new Error("Select a picture on the slide first.");
    const png = shape.getImageAsBase64({ width: Math.round(shape.width * RENDER_SCALE) });
    await ctx.sync();
    const { id, name, left, top, width, height } = shape;
    return { id, name, left, top, width, height, png: png.value };
  });
}

// Puts `base64` at `frame` on the current slide and deletes the original picture.
async function replacePicture(original, base64, frame, name) {
  // Insert first, delete after: if the insert fails, the original picture is still there.
  await insertPicture(base64, frame);
  await PowerPoint.run(async (ctx) => {
    const shapes = ctx.presentation.getSelectedSlides().getItemAt(0).shapes;
    const old = shapes.getItemOrNullObject(original.id);
    await ctx.sync();
    if (!old.isNullObject) old.delete();
    shapes.load("items/name");
    await ctx.sync();
    shapes.items[shapes.items.length - 1].name = name;
    await ctx.sync();
  });
}

async function invertSelectedImage() {
  const original = await selectedPicture("Inverting");
  await replacePicture(original, await invertPng(original.png), original, `${original.name} (inverted)`);
  status(`Inverted "${original.name}".`);
}

// ✂️ BiRefNet runs next to the pane, in server.js — a few GB of torch can't live in a browser.
function prewarmBackgroundRemoval() {
  fetch("/bg-remove/prewarm", { method: "POST" }).catch(() => {});
}

async function removeBackground() {
  const original = await selectedPicture("Removing the background");
  status("Removing the background… (the first one loads the model, ~10 s)");
  const bytes = Uint8Array.from(atob(original.png), (c) => c.charCodeAt(0));
  const res = await fetch("/bg-remove", { method: "POST", headers: { "Content-Type": "image/png" }, body: bytes });
  const out = await res.json();
  if (!res.ok) throw new Error(out.error || `Background removal failed (${res.status}).`);
  // The cut-out comes back trimmed to the subject: put it exactly where the subject was.
  const [x0, y0, x1, y1] = out.box, [w, h] = out.size;
  const sx = original.width / w, sy = original.height / h;
  const frame = { left: original.left + x0 * sx, top: original.top + y0 * sy,
    width: (x1 - x0) * sx, height: (y1 - y0) * sy };
  await replacePicture(original, out.png, frame, `${original.name} (no bg)`);
  status(`Removed the background of "${original.name}" (${(out.ms / 1000).toFixed(1)} s).`);
}

// 🧩 Prefix of the temporary names that find the selected shapes again in the exported slide.
const MARK = "\u2063victor-tools-regroup-";

// The selected shapes, each one replaced by its top-level group when it was clicked inside one.
async function selectedTopLevelShapes(ctx) {
  const selected = ctx.presentation.getSelectedShapes();
  selected.load("items/id,items/name,items/level");
  await ctx.sync();
  let shapes = selected.items;
  while (shapes.some((s) => s.level > 0)) {
    shapes = shapes.map((s) => (s.level > 0 ? s.parentGroup : s));
    shapes.forEach((s) => s.load("id,name,level"));
    await ctx.sync();
  }
  return [...new Map(shapes.map((s) => [s.id, s])).values()];
}

// Exports the current slide with the selected shapes renamed to MARK markers (names restored at once).
async function exportSlideWithMarks() {
  return PowerPoint.run(async (ctx) => {
    const slide = ctx.presentation.getSelectedSlides().getItemAt(0);
    slide.load("id");
    const shapes = await selectedTopLevelShapes(ctx);
    if (shapes.length < 2) throw new Error("Select the animated group and the shapes to add to it.");
    const marks = new Map(shapes.map((s, i) => [`${MARK}${i}`, s.name]));
    shapes.forEach((s, i) => { s.name = `${MARK}${i}`; });
    await ctx.sync();
    try {
      const pptx = slide.exportAsBase64();
      await ctx.sync();
      return { slideId: slide.id, marks, pptx: pptx.value };
    } finally {
      shapes.forEach((s, i) => { s.name = marks.get(`${MARK}${i}`); });
      await ctx.sync();
    }
  });
}

async function addToAnimatedGroup() {
  if (!Office.context.requirements.isSetSupported("PowerPointApi", "1.8")) {
    throw new Error("Needs PowerPointApi 1.8 (PowerPoint for Mac 16.96 or newer).");
  }
  const { slideId, marks, pptx } = await exportSlideWithMarks();
  const zip = await JSZip.loadAsync(pptx, { base64: true });
  const slides = Object.keys(zip.files).filter((f) => /^ppt\/slides\/slide\d+\.xml$/.test(f));
  if (slides.length !== 1) throw new Error(`The exported slide came back as ${slides.length} slides.`);
  const doc = new DOMParser().parseFromString(await zip.file(slides[0]).async("string"), "application/xml");
  const result = Regroup.addToGroup(doc, marks);
  const xml = new XMLSerializer().serializeToString(doc);
  zip.file(slides[0], xml.startsWith("<?xml") ? xml : `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${xml}`);
  const rebuilt = await zip.generateAsync({ type: "base64", compression: "DEFLATE" });

  // Insert the rebuilt slide right after the original, then drop the original.
  await PowerPoint.run(async (ctx) => {
    const all = ctx.presentation.slides;
    all.load("items/id");
    await ctx.sync();
    const before = new Set(all.items.map((s) => s.id));
    ctx.presentation.insertSlidesFromBase64(rebuilt, {
      formatting: PowerPoint.InsertSlideFormatting.useDestinationTheme, targetSlideId: slideId,
    });
    await ctx.sync();
    all.load("items/id");
    await ctx.sync();
    const inserted = all.items.find((s) => !before.has(s.id));
    if (!inserted) throw new Error("PowerPoint didn't insert the rebuilt slide; the original is untouched.");
    ctx.presentation.slides.getItem(slideId).delete();
    await ctx.sync();
    ctx.presentation.setSelectedSlides([inserted.id]);
    await ctx.sync();
  });
  const dropped = result.droppedAnimations ? ` Dropped ${result.droppedAnimations} effect(s) of the added shapes.` : "";
  const kept = result.animated ? "keeps its animation" : "had no animation";
  status(`Added ${result.added.map((n) => `"${n}"`).join(", ")} to "${result.group}", which ${kept}.${dropped}`);
}

// Disables the clicked button while `fn` runs: a second click would replace the same picture twice.
function guarded(fn) {
  return async (event) => {
    const button = event && event.currentTarget;
    if (button) button.disabled = true;
    try { await fn(); } catch (e) { status(e.message || String(e), true); console.error(e); }
    finally { if (button) button.disabled = false; }
  };
}

Office.onReady(() => {
  document.querySelectorAll(".tab").forEach((tab) => {
    tab.onclick = () => {
      document.querySelectorAll(".tab, .panel").forEach((el) => el.classList.remove("active"));
      tab.classList.add("active");
      $(tab.dataset.tab).classList.add("active");
      if (tab.dataset.tab === "image") prewarmBackgroundRemoval();
    };
  });
  $("insert").onclick = guarded(insertCode);
  $("scan").onclick = guarded(scanDeck);
  $("invert").onclick = guarded(invertSelectedImage);
  $("remove-bg").onclick = guarded(removeBackground);
  $("add-to-group").onclick = guarded(addToAnimatedGroup);
  if (!Office.context.requirements.isSetSupported("PowerPointApi", "1.5")) {
    status("This PowerPoint is too old for the add-in (needs PowerPointApi 1.5).", true);
  }
});
