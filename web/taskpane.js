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

function guarded(fn) {
  return async () => {
    try { await fn(); } catch (e) { status(e.message || String(e), true); console.error(e); }
  };
}

Office.onReady(() => {
  document.querySelectorAll(".tab").forEach((tab) => {
    tab.onclick = () => {
      document.querySelectorAll(".tab, .panel").forEach((el) => el.classList.remove("active"));
      tab.classList.add("active");
      $(tab.dataset.tab).classList.add("active");
    };
  });
  $("insert").onclick = guarded(insertCode);
  $("scan").onclick = guarded(scanDeck);
  if (!Office.context.requirements.isSetSupported("PowerPointApi", "1.5")) {
    status("This PowerPoint is too old for the add-in (needs PowerPointApi 1.5).", true);
  }
});
