// 🌗 Turns a light slide dark (or a dark one light) and back, exactly.
//
// What moves and how:
// - the slide swaps bg1 ↔ tx1 in its own color map (<p:clrMapOvr>), so everything it inherits from
//   the master and layout (background, placeholder text) flips by itself;
// - every color written on the slide is judged by its role: a SURFACE (fill, highlight, background)
//   goes to the other end of the lightness scale and keeps a tint of its hue; INK (text, lines,
//   bullets) keeps the contrast it had against whatever is behind it, and stays untouched when it
//   already has it (a red line is as red on black as on white);
// - a picture that is ink on paper (line art, diagrams, light screenshots) gets its lightness
//   flipped per pixel, hue kept; photos stay.
// The way back is not computed, it is looked up: every color this writes is a unique sRGB value,
// and the slide's tag remembers what stood there before. Colors added while dark are computed.
// Pure DOM + arrays, no Office.js, so test/dark.test.js and tools/dark-deck.js run it in Node.
(function (root) {
  const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
  const P = "http://schemas.openxmlformats.org/presentationml/2006/main";
  const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  const COLOR_TAGS = new Set(["srgbClr", "schemeClr", "sysClr", "prstClr", "scrgbClr"]);
  const PRESET = { black: "000000", white: "FFFFFF", red: "FF0000", green: "008000", blue: "0000FF", yellow: "FFFF00",
    gray: "808080", ltGray: "D3D3D3", dkGray: "A9A9A9", orange: "FFA500", lime: "00FF00", cyan: "00FFFF", magenta: "FF00FF" };
  // What the tag stores instead of a color when this added an element that wasn't there.
  const ADDED_FILL = "!fill", ADDED_RPR = "!rPr", ADDED_BG = "!bg";

  // ---------- color math ----------
  const clamp01 = (x) => Math.min(1, Math.max(0, x));
  const hexOf = (rgb) => rgb.map((c) => Math.round(c).toString(16).padStart(2, "0")).join("").toUpperCase();
  const rgbOf = (hex) => [0, 2, 4].map((i) => parseInt(hex.slice(i, i + 2), 16));
  const toLinear = (c) => { c /= 255; return c <= 0.04045 ? c / 12.92 : Math.pow((c + 0.055) / 1.055, 2.4); };
  const fromLinear = (c) => 255 * (c <= 0.0031308 ? 12.92 * c : 1.055 * Math.pow(c, 1 / 2.4) - 0.055);
  // WCAG relative luminance and contrast ratio.
  const luminance = (rgb) => 0.2126 * toLinear(rgb[0]) + 0.7152 * toLinear(rgb[1]) + 0.0722 * toLinear(rgb[2]);
  const contrast = (a, b) => { const x = luminance(a), y = luminance(b); return (Math.max(x, y) + 0.05) / (Math.min(x, y) + 0.05); };
  const isDark = (rgb) => luminance(rgb) < 0.179; // below it, white text reads better than black

  function toOklch(rgb) {
    const [r, g, b] = rgb.map(toLinear);
    const l = Math.cbrt(0.4122214708 * r + 0.5363325363 * g + 0.0514459929 * b);
    const m = Math.cbrt(0.2119034982 * r + 0.6806995451 * g + 0.1073969566 * b);
    const s = Math.cbrt(0.0883024619 * r + 0.2817188376 * g + 0.6299787005 * b);
    const L = 0.2104542553 * l + 0.793617785 * m - 0.0040720468 * s;
    const a = 1.9779984951 * l - 2.428592205 * m + 0.4505937099 * s;
    const bb = 0.0259040371 * l + 0.7827717662 * m - 0.808675766 * s;
    return [L, Math.hypot(a, bb), (Math.atan2(bb, a) * 180 / Math.PI + 360) % 360];
  }
  function oklchToLinear(L, C, h) {
    const a = C * Math.cos(h * Math.PI / 180), b = C * Math.sin(h * Math.PI / 180);
    const l = Math.pow(L + 0.3963377774 * a + 0.2158037573 * b, 3);
    const m = Math.pow(L - 0.1055613458 * a - 0.0638541728 * b, 3);
    const s = Math.pow(L - 0.0894841775 * a - 1.291485548 * b, 3);
    return [4.0767416621 * l - 3.3077115913 * m + 0.2309699292 * s,
      -1.2684380046 * l + 2.6097574011 * m - 0.3413193965 * s,
      -0.0041960863 * l - 0.7034186147 * m + 1.707614701 * s];
  }
  // The sRGB color at this lightness and hue with as much of the chroma as the gamut holds.
  function fromOklch(L, C, h) {
    L = clamp01(L);
    const inGamut = (c) => oklchToLinear(L, c, h).every((x) => x >= -0.0005 && x <= 1.0005);
    if (!inGamut(C)) {
      let lo = 0, hi = C;
      for (let i = 0; i < 20; i++) { const mid = (lo + hi) / 2; if (inGamut(mid)) lo = mid; else hi = mid; }
      C = lo;
    }
    return oklchToLinear(L, C, h).map((x) => Math.round(fromLinear(clamp01(x))));
  }
  const smooth = (a, b, x) => { const t = clamp01((x - a) / (b - a)); return t * t * (3 - 2 * t); };

  // A yellow made darker turns olive-brown. Dark yellow that still reads as yellow is an orange,
  // so the hue slides toward 70° by how much lightness was lost.
  const inYellow = (h) => smooth(72, 90, h) * (1 - smooth(112, 128, h));
  const yellowToOrange = (h, lost) => h - 34 * inYellow(h) * clamp01(lost / 0.45);

  // SURFACE: white ↔ black, grays mirror, and a colored surface lands on a tinted dark (or light)
  // that still shows its hue: the more chroma, the further from the black it stays.
  function surface(rgb, toDark) {
    let [L, C, h] = toOklch(rgb);
    if (!toDark) L = 1 - L;
    const floor = smooth(0, 0.04, C) * Math.min(0.47, 0.28 + 0.8 * C);
    let L2 = Math.min(L, Math.max(1 - L, floor));
    // A light-gray panel was visible on the white page; mirrored it would be a black no projector
    // tells from the page. It stops at #242424. Pure white still becomes the page.
    if (C < 0.04 && L > 0.8 && L < 0.995) L2 = Math.max(L2, 0.26);
    // A surface takes only a little of the orange: a deep gold sits better next to greens than a brown.
    if (toDark) h -= 15 * inYellow(h) * clamp01((L - L2) / 0.45);
    return fromOklch(toDark ? L2 : 1 - L2, C, h);
  }

  // MARKER: a highlighter stroke, or a loud fill under text (a yellow star, a red badge). Dimmed
  // to a deep tint it would stop doing its job, so on a dark slide it stays a light, loud color
  // with dark text on it — and a yellow becomes an orange-amber: pure yellow on black is as
  // bright as the white text around it and projectors push it toward lime.
  function marker(rgb, toDark) {
    const [L, C, h] = toOklch(rgb);
    if (!toDark || C < 0.15 || L < 0.6) return surface(rgb, toDark);
    return fromOklch(Math.min(L, 0.82), C, Math.max(Math.min(h, 70), h - 34 * inYellow(h)));
  }

  // The PAGE of a dark slide turned light is paper, not the mirror of its near-black (a dusty mauve).
  function pageColor(rgb, toDark) {
    if (toDark) return surface(rgb, toDark);
    const [L, C, h] = toOklch(rgb);
    return fromOklch(Math.max(1 - L, 0.975), Math.min(C, 0.008), h);
  }

  // INK: keeps the contrast ratio it had against what was behind it. Chromatic ink that already
  // has it on the new backdrop is not touched; gray ink mirrors, so a faint gray stays faint.
  function ink(rgb, backOld, backNew) {
    let want = contrast(rgb, backOld);
    const have = contrast(rgb, backNew), [L, C, h] = toOklch(rgb);
    // Colored ink aims for 6–7:1. Below 6 a thin colored stroke on black smears in a screen share
    // (video carries color at half resolution); above 7 it has to give up its color to get there.
    if (C >= 0.06) want = Math.min(Math.max(want, 6), 7);
    const flipped = isDark(backOld) !== isDark(backNew);
    // The mirror is for grays in between: pure white or black text was never meant to be faint,
    // so when it reads better than before (the backdrop was guessed lighter than it is), it stays.
    // (Unless it was invisible — white on white, an overflow nobody saw: that stays invisible.)
    const faintGray = C < 0.04 && ((L > 0.1 && L < 0.94) || want < 1.25);
    if (have >= want * 0.98 && !(faintGray && flipped && have > want * 1.15)) return rgb;
    const yb = luminance(backNew), lighter = isDark(backNew);
    const enough = (c) => (lighter ? (luminance(c) + 0.05) / (yb + 0.05) : (yb + 0.05) / (luminance(c) + 0.05)) >= want;
    const at = (l) => fromOklch(l, C, lighter ? h : yellowToOrange(h, L - l));
    // Walk the lightness from the backdrop's side to the far end; stop at the first that has the contrast.
    let lo = lighter ? 0 : 1, hi = lighter ? 1 : 0;
    if (!enough(at(hi))) return at(hi);
    for (let i = 0; i < 22; i++) { const mid = (lo + hi) / 2; if (enough(at(mid))) hi = mid; else lo = mid; }
    return at(hi);
  }

  // Lightness flipped, hue and saturation kept (HSL): every channel moves by 255 - max - min.
  // Exact on integers and its own inverse, which is what makes the pictures reversible.
  function flipLightness(rgb) {
    const shift = 255 - Math.max(rgb[0], rgb[1], rgb[2]) - Math.min(rgb[0], rgb[1], rgb[2]);
    return [rgb[0] + shift, rgb[1] + shift, rgb[2] + shift];
  }
  function invertPixels(rgba) {
    for (let i = 0; i < rgba.length; i += 4) {
      const r = rgba[i], g = rgba[i + 1], b = rgba[i + 2];
      const shift = 255 - Math.max(r, g, b) - Math.min(r, g, b);
      rgba[i] = r + shift; rgba[i + 1] = g + shift; rgba[i + 2] = b + shift;
    }
    return rgba;
  }

  // ---------- pictures ----------
  // What a picture is made of, from ~40k sampled pixels: how much is see-through, paper-white,
  // ink-black, gray, and how many colors it takes to paint 90% of what is neither white nor black
  // (few = drawn, many = photo).
  // `opts.transparent` is the color a "Set transparent color" (<a:clrChange>) removes;
  // `opts.duotone` = [dark, light] recolors the picture by luminance before it is judged.
  function imageStats(rgba, width, height, opts = {}) {
    const step = Math.max(1, Math.floor(Math.sqrt((width * height) / 40000)));
    const bins = new Map();
    let n = 0, clear = 0, seen = 0, white = 0, black = 0, gray = 0, edge = 0, edgeClear = 0, edgeWhite = 0, edgeBlack = 0;
    const sum = [0, 0, 0];
    const key = opts.transparent;
    for (let y = 0; y < height; y += step) {
      for (let x = 0; x < width; x += step) {
        const i = (y * width + x) * 4;
        let r = rgba[i], g = rgba[i + 1], b = rgba[i + 2];
        const a = rgba[i + 3];
        const onEdge = x < step * 2 || y < step * 2 || x >= width - step * 2 || y >= height - step * 2;
        n++; if (onEdge) edge++;
        if (a < 24 || (key && Math.abs(r - key[0]) + Math.abs(g - key[1]) + Math.abs(b - key[2]) < 12)) {
          clear++; if (onEdge) edgeClear++;
          continue;
        }
        if (opts.duotone) {
          const t = (0.299 * r + 0.587 * g + 0.114 * b) / 255, [d, l] = opts.duotone;
          r = d[0] + (l[0] - d[0]) * t; g = d[1] + (l[1] - d[1]) * t; b = d[2] + (l[2] - d[2]) * t;
        }
        seen++;
        sum[0] += r; sum[1] += g; sum[2] += b;
        const max = Math.max(r, g, b), min = Math.min(r, g, b);
        const isWhite = min > 222, isBlack = max < 56;
        if (isWhite) { white++; if (onEdge) edgeWhite++; } else if (isBlack) { black++; if (onEdge) edgeBlack++; } else if (max - min < 28) gray++;
        if (isWhite || isBlack) continue;
        const bin = ((r >> 4) << 8) | ((g >> 4) << 4) | (b >> 4);
        bins.set(bin, (bins.get(bin) || 0) + 1);
      }
    }
    const counts = [...bins.values()].sort((p, q) => q - p);
    let covered = 0, colors = 0;
    while (colors < counts.length && covered < (seen - white - black) * 0.9) covered += counts[colors++];
    const of = (k, total) => (total ? k / total : 0);
    return { clear: of(clear, n), white: of(white, seen), black: of(black, seen), gray: of(gray, seen), colors,
      edgeClear: of(edgeClear, edge), edgeWhite: of(edgeWhite, edge - edgeClear), edgeBlack: of(edgeBlack, edge - edgeClear),
      mean: seen ? sum.map((s) => s / seen) : [128, 128, 128] };
  }

  // Flip the picture when it is ink on paper and the paper is about to change sides:
  // - a cut-out (see-through around it) that is black/white/gray drawing with ink that would sink
  //   into the new background;
  // - an opaque picture on a paper-colored field (diagram, chart, screenshot of a light window),
  //   drawn with few colors. Photos, dark screenshots and colorful art are left as they are.
  function shouldInvert(s, toDark) {
    const paper = toDark ? s.white : s.black, inkShare = toDark ? s.black : s.white;
    const edgePaper = toDark ? s.edgeWhite : s.edgeBlack;
    const drawn = s.colors <= 48;
    if (s.clear >= 0.2 || s.edgeClear >= 0.5) {
      return inkShare >= 0.3 && inkShare + paper + s.gray >= 0.8 && s.colors <= 40;
    }
    if (edgePaper >= 0.9 && paper >= 0.3 && s.colors <= 8) return true; // a flat icon on a paper card
    return edgePaper >= 0.7 && paper >= 0.5 && drawn;
  }

  // ---------- DOM helpers ----------
  const elements = (node) => (node ? [...node.childNodes].filter((n) => n.nodeType === 1) : []);
  const child = (node, ns, name) => elements(node).find((n) => n.namespaceURI === ns && n.localName === name);
  const kids = (node, ns, name) => elements(node).filter((n) => n.namespaceURI === ns && n.localName === name);
  const path = (node, ...steps) => steps.reduce((at, [ns, name]) => at && child(at, ns, name), node);
  const colorIn = (node) => elements(node).find((n) => n.namespaceURI === A && COLOR_TAGS.has(n.localName));
  const descendants = (node, ns, name) => (node ? [...node.getElementsByTagNameNS(ns, name)] : []);

  // "<a:srgbClr val=X/>" → "#X", "<a:schemeClr val=n/>" → "@n", anything else as XML without xmlns noise.
  function encode(el, serialize) {
    if (!elements(el).length && el.attributes.length === 1 && el.getAttribute("val") !== null) {
      if (el.localName === "srgbClr") return `#${el.getAttribute("val")}`;
      if (el.localName === "schemeClr") return `@${el.getAttribute("val")}`;
    }
    return serialize(el).replace(/ xmlns(:\w+)?="[^"]*"/g, "");
  }
  function decode(text, doc, parse) {
    if (text[0] === "#" || text[0] === "@") {
      const el = doc.createElementNS(A, text[0] === "#" ? "a:srgbClr" : "a:schemeClr");
      el.setAttribute("val", text.slice(1));
      return el;
    }
    const wrapped = parse(`<x xmlns:a="${A}" xmlns:p="${P}" xmlns:r="${R}">${text}</x>`);
    return doc.importNode(elements(wrapped.documentElement)[0], true);
  }

  // ---------- resolving a color element to sRGB ----------
  function hsl(rgb) {
    const [r, g, b] = rgb.map((c) => c / 255), max = Math.max(r, g, b), min = Math.min(r, g, b), l = (max + min) / 2;
    if (max === min) return [0, 0, l];
    const d = max - min, s = d / (1 - Math.abs(2 * l - 1));
    const h = max === r ? ((g - b) / d + 6) % 6 : max === g ? (b - r) / d + 2 : (r - g) / d + 4;
    return [h * 60, s, l];
  }
  function unhsl([h, s, l]) {
    const c = (1 - Math.abs(2 * l - 1)) * s, x = c * (1 - Math.abs(((h / 60) % 2) - 1)), m = l - c / 2;
    const [r, g, b] = h < 60 ? [c, x, 0] : h < 120 ? [x, c, 0] : h < 180 ? [0, c, x] : h < 240 ? [0, x, c] : h < 300 ? [x, 0, c] : [c, 0, x];
    return [r, g, b].map((v) => clamp01(v + m) * 255);
  }
  // → { rgb, alpha } or null when the color can't be known (placeholder color, unknown preset).
  function resolve(el, env, map) {
    let rgb;
    const val = el.getAttribute("val");
    if (el.localName === "srgbClr") rgb = rgbOf(val);
    else if (el.localName === "schemeClr") rgb = env.theme[map[val] || val];
    else if (el.localName === "sysClr") rgb = el.getAttribute("lastClr") ? rgbOf(el.getAttribute("lastClr")) : val === "window" ? [255, 255, 255] : val === "windowText" ? [0, 0, 0] : null;
    else if (el.localName === "prstClr") rgb = PRESET[val] && rgbOf(PRESET[val]);
    else if (el.localName === "scrgbClr") rgb = ["r", "g", "b"].map((k) => fromLinear(clamp01(Number(el.getAttribute(k)) / 100000)));
    if (!rgb) return null;
    let alpha = 1;
    for (const mod of elements(el)) {
      const v = Number(mod.getAttribute("val")) / 100000;
      if (mod.localName === "alpha") alpha = v;
      else if (mod.localName === "lumMod") { const c = hsl(rgb); c[2] = clamp01(c[2] * v); rgb = unhsl(c); }
      else if (mod.localName === "lumOff") { const c = hsl(rgb); c[2] = clamp01(c[2] + v); rgb = unhsl(c); }
      else if (mod.localName === "satMod") { const c = hsl(rgb); c[1] = clamp01(c[1] * v); rgb = unhsl(c); }
      else if (mod.localName === "tint") rgb = rgb.map((c) => c * v + 255 * (1 - v));
      else if (mod.localName === "shade") rgb = rgb.map((c) => c * v);
    }
    return { rgb: rgb.map((c) => Math.round(c)), alpha };
  }
  const over = (top, alpha, under) => top.map((c, i) => c * alpha + under[i] * (1 - alpha));
  const average = (list) => [0, 1, 2].map((i) => list.reduce((s, c) => s + c[i], 0) / list.length);
  const same = (a, b) => Math.max(Math.abs(a[0] - b[0]), Math.abs(a[1] - b[1]), Math.abs(a[2] - b[2])) <= 2;

  // ---------- the theme and the color map a slide lives in ----------
  const MAP_KEYS = ["bg1", "tx1", "bg2", "tx2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"];
  function themeColors(themeDoc) {
    const colors = {};
    for (const el of elements(descendants(themeDoc, A, "clrScheme")[0])) {
      const c = elements(el)[0];
      const hex = c && (c.getAttribute("lastClr") || (c.localName === "srgbClr" ? c.getAttribute("val") : c.getAttribute("val") === "window" ? "FFFFFF" : "000000"));
      if (hex) colors[el.localName] = rgbOf(hex);
    }
    return colors;
  }
  const mapOf = (el) => Object.fromEntries(MAP_KEYS.map((k) => [k, el.getAttribute(k)]));

  // ---------- geometry: where a shape is, to know what is behind it ----------
  function frameOf(xfrm) {
    const off = xfrm && child(xfrm, A, "off"), ext = xfrm && child(xfrm, A, "ext");
    if (!off || !ext) return null;
    return { x: Number(off.getAttribute("x")), y: Number(off.getAttribute("y")), w: Number(ext.getAttribute("cx")), h: Number(ext.getAttribute("cy")) };
  }
  // A group draws its children in its own coordinates: chOff/chExt stretched onto off/ext.
  function groupTransform(grpSp, outer) {
    const xfrm = path(grpSp, [P, "grpSpPr"], [A, "xfrm"]);
    const frame = frameOf(xfrm), chOff = xfrm && child(xfrm, A, "chOff"), chExt = xfrm && child(xfrm, A, "chExt");
    if (!frame || !chOff || !chExt) return outer;
    const cx = Number(chOff.getAttribute("x")), cy = Number(chOff.getAttribute("y"));
    const sx = frame.w / (Number(chExt.getAttribute("cx")) || 1), sy = frame.h / (Number(chExt.getAttribute("cy")) || 1);
    return (f) => outer({ x: frame.x + (f.x - cx) * sx, y: frame.y + (f.y - cy) * sy, w: f.w * sx, h: f.h * sy });
  }
  // The placeholder of a layout or master that a slide's placeholder takes its place and style from.
  function matchingPlaceholder(doc, ph) {
    if (!doc || !ph) return null;
    const type = ph.getAttribute("type") || "body", idx = ph.getAttribute("idx");
    const all = descendants(doc, P, "ph");
    const found = (idx !== null && all.find((p) => p.getAttribute("idx") === idx))
      || all.find((p) => (p.getAttribute("type") || "body") === type)
      || (type === "ctrTitle" && all.find((p) => p.getAttribute("type") === "title"))
      || (type === "subTitle" && all.find((p) => (p.getAttribute("type") || "body") === "body"));
    return found ? found.parentNode.parentNode.parentNode : null; // ph → nvPr → nvSpPr → sp
  }

  // The background a slide shows: its own, else the layout's, else the master's.
  function backgroundOf(slide, env) {
    const owner = [slide, env.layout, env.master].find((d) => d && descendants(d, P, "bg")[0]);
    const el = owner && descendants(owner, P, "bg")[0];
    const colors = el ? descendants(el, "*", "*").filter((e) => e.namespaceURI === A && COLOR_TAGS.has(e.localName)) : [];
    return { owner, el, colors, picture: !!(el && descendants(el, A, "blipFill").length) };
  }

  // True when the slide's page reads dark (white text reads better on it); false for a picture background.
  function isSlideDark(slide, env) {
    const theme = themeColors(env.theme);
    const ovr = descendants(slide, P, "clrMapOvr")[0];
    const override = ovr && child(ovr, A, "overrideClrMapping");
    const map = override ? mapOf(override) : mapOf(descendants(env.master, P, "clrMap")[0]);
    const bg = backgroundOf(slide, env);
    if (bg.picture) return false;
    const seen = bg.colors.map((el) => resolve(el, { theme, map }, map)).filter(Boolean).map((c) => c.rgb);
    return isDark(seen.length ? average(seen) : theme[map.bg1]);
  }

  // ---------- the slide ----------
  // env: { layout, master, theme (docs), tag (what the last toggle stored, or null),
  //        images: Map blip rId → imageStats(...) raw result plus {width,height} — or absent,
  //        serialize(node) → string, parse(string) → Document }
  // Changes `slide` in place. Returns { tag, pixels: [blip elements whose picture must be flipped],
  // toDark, changed, kept, pictures, skipped? }.
  function invertSlide(slide, env) {
    const cSld = descendants(slide, P, "cSld")[0];
    const theme = themeColors(env.theme);
    const tag = env.tag || {};
    const table = tag.m || {};          // color this wrote last time → what stood there before
    const derived = !!tag.o;            // true: the slide is in its toggled state, on its way home
    const out = {};                     // the table for the next toggle
    const used = new Set(descendants(slide, A, "srgbClr").map((e) => e.getAttribute("val").toUpperCase()));
    for (const v of Object.values(table)) for (const m of v.matchAll(/#([0-9A-Fa-f]{6})|srgbClr val="([0-9A-Fa-f]{6})"/g)) used.add((m[1] || m[2]).toUpperCase());

    // The color map: the slide's own, else the master's; the new one has bg1 and tx1 swapped.
    const masterMap = mapOf(descendants(env.master, P, "clrMap")[0]);
    let ovr = descendants(slide, P, "clrMapOvr")[0];
    if (!ovr) ovr = slide.documentElement.insertBefore(slide.createElementNS(P, "p:clrMapOvr"), cSld.nextSibling);
    const override = ovr && child(ovr, A, "overrideClrMapping");
    const mapOld = override ? mapOf(override) : masterMap;
    const mapNew = { ...mapOld, bg1: mapOld.tx1, tx1: mapOld.bg1 };
    const R_old = { theme, map: mapOld }, R_new = { theme, map: mapNew };
    const was = (el) => resolve(el, R_old, mapOld), willBe = (el) => resolve(el, R_new, mapNew);

    const { owner: bgOwner, el: bgEl, colors: bgColors, picture: bgPicture } = backgroundOf(slide, env);
    if (bgPicture) return { skipped: "its background is a picture" };
    const bgWas = bgColors.map(was).filter(Boolean).map((c) => c.rgb);
    const pageOld = bgWas.length ? average(bgWas) : theme[mapOld.bg1];
    const toDark = !isDark(pageOld);
    const stats = { changed: 0, kept: 0, pictures: 0 };

    // A new sRGB value that is the only one standing for `original`.
    function unique(rgb, original) {
      for (let d = 0; d < 48; d++) {
        for (const sign of d ? [1, -1] : [1]) {
          const c = rgb.slice();
          const i = d % 3; // nudge blue, green, red in turn: invisible at ±16 steps of 1
          c[2 - i] = Math.min(255, Math.max(0, c[2 - i] + sign * Math.ceil(d / 3)));
          const hex = hexOf(c);
          if (out[hex] === original || (out[hex] === undefined && !used.has(hex))) { out[hex] = original; used.add(hex); return hex; }
        }
      }
      throw new Error("Too many near-identical colors on one slide.");
    }

    // Decides one color element. role "surface" or "ink"; back = { old, new } behind an ink.
    // Returns { old, new } as sRGB (with alpha), or null when the color is unknowable.
    const decided = new Map();
    function decide(el, role, back, opaque) {
      const d = decide0(el, role, back, opaque);
      if (d) decided.set(d.el, d);
      return d;
    }
    function decide0(el, role, back, opaque) {
      const old = was(el);
      if (!old) return null;
      const val = el.localName === "srgbClr" && el.getAttribute("val").toUpperCase();
      if (val && table[val] !== undefined && table[val][0] !== "!") {         // written by the last toggle: put back
        const original = decode(table[val], slide, env.parse);
        el.parentNode.replaceChild(original, el);
        stats.changed++;
        return { old, new: willBe(original) || old, el: original };
      }
      const natural = willBe(el);
      if (derived && !val) return { old, new: natural, el };                 // theme colors travel with the map
      let wanted;
      if (role === "ink") wanted = ink(old.rgb, back.old, back.new);
      else if (old.alpha < 1 && opaque) {
        // A see-through background is painted on white paper, which no color can darken: it becomes solid.
        return rewrite(el, pageColor(over(old.rgb, old.alpha, [255, 255, 255]), toDark), encode(el, env.serialize), old, true);
      } else if (old.alpha < 1 && old.alpha >= 0.12 && back) {
        // A see-through surface: what counts is how it looks over what is behind it, before and after.
        const seen = surface(over(old.rgb, old.alpha, back.old), toDark);
        wanted = seen.map((c, i) => Math.min(255, Math.max(0, Math.round((c - back.new[i] * (1 - old.alpha)) / old.alpha))));
      } else if (role === "marker") wanted = marker(old.rgb, toDark);
      else if (role === "graphic" && toOklch(old.rgb)[1] >= 0.12 && toOklch(old.rgb)[0] <= 0.82 && old.alpha >= 0.99 && contrast(old.rgb, page.new) >= 3) {
        wanted = old.rgb; // a colored shape with no text (an arrow, a tick) that reads on the new page keeps its color
      } else wanted = role === "page" ? pageColor(old.rgb, toDark) : surface(old.rgb, toDark);
      const text = encode(el, env.serialize);
      if (same(wanted, natural.rgb)) {
        // Stays as written. A plain sRGB still goes into the table, so the way back knows it is not new.
        if (val && (out[val] === undefined || out[val] === text) && elements(el).every((m) => m.localName === "alpha")) out[val] = text;
        else if (val) return rewrite(el, natural.rgb, text, old);
        stats.kept++;
        return { old, new: natural, el };
      }
      return rewrite(el, wanted, text, old);
    }
    function rewrite(el, rgb, original, old, opaque) {
      const fresh = slide.createElementNS(A, "a:srgbClr");
      fresh.setAttribute("val", unique(rgb, original));
      if (!opaque) for (const mod of kids(el, A, "alpha")) fresh.appendChild(mod.cloneNode(true));
      el.parentNode.replaceChild(fresh, el);
      stats.changed++;
      return { old, new: { rgb: rgbOf(fresh.getAttribute("val")), alpha: opaque ? 1 : old.alpha }, el: fresh };
    }

    // --- background ---
    let page = { old: pageOld, new: null };
    if (bgOwner === slide) {
      const marker = bgColors.length === 1 && bgColors[0].localName === "srgbClr" && table[bgColors[0].getAttribute("val").toUpperCase()];
      if (marker === ADDED_BG) {
        bgEl.parentNode.removeChild(bgEl);
        const inherited = [env.layout, env.master].map((d) => d && descendants(d, P, "bg")[0]).find(Boolean);
        const colors = inherited ? descendants(inherited, "*", "*").filter((e) => e.namespaceURI === A && COLOR_TAGS.has(e.localName)).map(willBe).filter(Boolean) : [];
        page.new = colors.length ? average(colors.map((c) => c.rgb)) : theme[mapNew.bg1];
      } else {
        const done = bgColors.map((el) => decide(el, "page", null, true)).filter(Boolean);
        page.new = done.length ? average(done.map((d) => d.new.rgb)) : theme[mapNew.bg1];
      }
    } else {
      const natural = bgColors.map(willBe).filter(Boolean).map((c) => c.rgb);
      page.new = natural.length ? average(natural) : theme[mapNew.bg1];
      {
        // The slide gets a background of its own: the color the map gives the inherited one, or,
        // when that one is written in fixed colors, its surface color. (PowerPoint repaints the
        // master's background through the slide's map; LibreOffice and Keynote don't.)
        const rgb = isDark(page.new) === isDark(pageOld) ? pageColor(pageOld, toDark) : page.new.map(Math.round);
        const bg = slide.createElementNS(P, "p:bg"), pr = bg.appendChild(slide.createElementNS(P, "p:bgPr"));
        const clr = pr.appendChild(slide.createElementNS(A, "a:solidFill")).appendChild(slide.createElementNS(A, "a:srgbClr"));
        clr.setAttribute("val", unique(rgb, ADDED_BG));
        pr.appendChild(slide.createElementNS(A, "a:effectLst"));
        cSld.insertBefore(bg, cSld.firstChild);
        page.new = rgbOf(clr.getAttribute("val"));
        stats.changed++;
      }
    }

    // --- links ---
    // A link is painted in the theme's hyperlink color, whatever its run says. When that color
    // sinks into the new background, the map points "hlink" at the theme color nearest in hue
    // that reads on it (a slide can't change the theme, but it can choose among its colors).
    if (derived && tag.h) [mapNew.hlink, mapNew.folHlink] = tag.h;
    else if (!derived) {
      for (const k of ["hlink", "folHlink"]) {
        const now = theme[mapOld[k]];
        if (!now || contrast(now, page.new) >= 4.5) continue;
        const hue = toOklch(now)[2];
        const readable = ["accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "dk2", "lt2", "hlink", "folHlink"]
          .filter((slot) => theme[slot] && contrast(theme[slot], page.new) >= 4.5 && toOklch(theme[slot])[1] > 0.05)
          .sort((a, b) => Math.abs(((toOklch(theme[a])[2] - hue + 540) % 360) - 180) - Math.abs(((toOklch(theme[b])[2] - hue + 540) % 360) - 180));
        mapNew[k] = readable[0] || mapNew.tx1;
      }
    }
    const linksMoved = !derived && (mapNew.hlink !== mapOld.hlink || mapNew.folHlink !== mapOld.folHlink);

    // --- shapes, back to front; `painted` is what has been drawn so far, to know what is behind ---
    const painted = [];
    const behind = (frame) => {
      if (frame) {
        const x = frame.x + frame.w / 2, y = frame.y + frame.h / 2;
        for (let i = painted.length - 1; i >= 0; i--) {
          const p = painted[i];
          if (x >= p.x && x <= p.x + p.w && y >= p.y && y <= p.y + p.h) return p.back;
        }
      }
      return page;
    };
    const blend = (done, back) => ({ old: over(done.old.rgb, done.old.alpha, back.old), new: over(done.new.rgb, done.new.alpha, back.new) });

    // A fill element (solid / gradient / pattern) as one backdrop, or null for no fill.
    function fill(holder, back, role = "surface") {
      const solid = child(holder, A, "solidFill"), grad = child(holder, A, "gradFill"), patt = child(holder, A, "pattFill");
      const colors = solid ? [colorIn(solid)] : grad ? descendants(grad, A, "gs").map(colorIn) : patt ? [colorIn(child(patt, A, "bgClr")), colorIn(child(patt, A, "fgClr"))] : [];
      const done = colors.filter(Boolean).map((el) => decide(el, role, back)).filter(Boolean);
      if (!done.length) return null;
      const alpha = average(done.map((d) => [d.old.alpha, 0, 0]))[0];
      if (alpha < 0.12) return null;
      return blend({ old: { rgb: average(done.map((d) => d.old.rgb)), alpha }, new: { rgb: average(done.map((d) => d.new.rgb)), alpha } }, back);
    }
    const hasFill = (holder) => ["solidFill", "gradFill", "pattFill", "noFill", "blipFill", "grpFill"].some((n) => child(holder, A, n));
    function line(ln, back) {
      const solid = ln && child(ln, A, "solidFill"), grad = ln && child(ln, A, "gradFill");
      for (const el of solid ? [colorIn(solid)] : grad ? descendants(grad, A, "gs").map(colorIn) : []) if (el) decide(el, "ink", back);
    }

    // The color a run takes when it has none of its own: the shape's list style, its style's
    // fontRef, the layout's and master's placeholder, the master's text styles, else tx1.
    function inheritedInk(sp, level, isTable) {
      const lvl = `lvl${level + 1}pPr`;
      const fromList = (lst) => { const c = lst && path(lst, [A, lvl], [A, "defRPr"], [A, "solidFill"]); return c && colorIn(c); };
      const own = fromList(path(sp, [P, "txBody"], [A, "lstStyle"]));
      if (own) return { el: own, local: true };
      const fontRef = path(sp, [P, "style"], [A, "fontRef"]);
      if (fontRef && colorIn(fontRef)) return { el: colorIn(fontRef), local: true };
      const ph = path(sp, [P, "nvSpPr"], [P, "nvPr"], [P, "ph"]);
      if (ph && !isTable) {
        for (const doc of [env.layout, env.master]) {
          const c = fromList(path(matchingPlaceholder(doc, ph), [P, "txBody"], [A, "lstStyle"]));
          if (c) return { el: c };
        }
        const type = ph.getAttribute("type") || "body";
        const style = /title/i.test(type) ? "titleStyle" : ["body", "subTitle", "obj"].includes(type) ? "bodyStyle" : "otherStyle";
        const c = fromList(path(env.master.documentElement, [P, "txStyles"], [P, style]));
        if (c) return { el: c };
      }
      return null;
    }

    function text(txBody, sp, back, isTable) {
      if (!txBody) return;
      for (const lst of kids(txBody, A, "lstStyle")) {
        for (const el of descendants(lst, "*", "*").filter((e) => e.namespaceURI === A && COLOR_TAGS.has(e.localName))) decide(el, "ink", back);
      }
      for (const p of kids(txBody, A, "p")) {
        const pPr = child(p, A, "pPr"), level = Number((pPr && pPr.getAttribute("lvl")) || 0);
        const bu = pPr && child(pPr, A, "buClr");
        if (bu && colorIn(bu)) decide(colorIn(bu), "ink", back);
        for (const run of elements(p).filter((e) => ["r", "fld", "br", "endParaRPr"].includes(e.localName))) {
          const rPr = run.localName === "endParaRPr" ? run : child(run, A, "rPr");
          let behindRun = back;
          const hl = rPr && child(rPr, A, "highlight");
          if (hl && colorIn(hl)) { const d = decide(colorIn(hl), "marker"); if (d) behindRun = blend(d, back); }
          const solid = rPr && child(rPr, A, "solidFill");
          if (solid && colorIn(solid)) {
            const c = colorIn(solid), marker = c.localName === "srgbClr" && table[c.getAttribute("val").toUpperCase()];
            if (marker === ADDED_FILL) { rPr.removeChild(solid); stats.changed++; }
            else if (marker === ADDED_RPR) { run.removeChild(rPr); stats.changed++; continue; }
            else decide(c, "ink", behindRun);
          } else if (rPr && child(rPr, A, "gradFill")) {
            for (const gs of descendants(child(rPr, A, "gradFill"), A, "gs")) if (colorIn(gs)) decide(colorIn(gs), "ink", behindRun);
          } else if (!derived && run.localName !== "br" && run.localName !== "endParaRPr" && !(rPr && (child(rPr, A, "noFill") || child(rPr, A, "pattFill") || child(rPr, A, "blipFill")))) {
            // No color of its own: it follows the map. When that leaves it unreadable here
            // (text over a photo that stays, or on a highlight), it gets a color of its own.
            const source = inheritedInk(sp, level, isTable);
            const known = source && decided.get(source.el);
            const old = known ? known.old : source ? was(source.el) : { rgb: theme[mapOld.tx1] };
            const natural = known ? known.new : source ? willBe(source.el) : { rgb: theme[mapNew.tx1] };
            if (old && natural) {
              const wanted = ink(old.rgb, behindRun.old, behindRun.new);
              if (!same(wanted, natural.rgb) && contrast(natural.rgb, behindRun.new) < Math.min(4.5, contrast(old.rgb, behindRun.old)) * 0.9) {
                const holder = rPr || run.insertBefore(slide.createElementNS(A, "a:rPr"), run.firstChild);
                const fresh = slide.createElementNS(A, "a:solidFill"), clr = fresh.appendChild(slide.createElementNS(A, "a:srgbClr"));
                clr.setAttribute("val", unique(wanted, rPr ? ADDED_FILL : ADDED_RPR));
                const ln = child(holder, A, "ln");
                holder.insertBefore(fresh, ln ? ln.nextSibling : holder.firstChild);
                stats.changed++;
              }
            }
          }
          if (!rPr || !rPr.parentNode) continue;
          line(child(rPr, A, "ln"), behindRun);
          const uFill = child(rPr, A, "uFill");
          if (uFill && child(uFill, A, "solidFill") && colorIn(child(uFill, A, "solidFill"))) decide(colorIn(child(uFill, A, "solidFill")), "ink", behindRun);
          for (const glow of descendants(child(rPr, A, "effectLst"), A, "glow")) if (colorIn(glow)) decide(colorIn(glow), "surface");
        }
      }
    }

    const pixels = [], pictures = {};
    const listed = tag.p || {};
    function picture(pic, frame, back) {
      const spPr = child(pic, P, "spPr");
      line(spPr && child(spPr, A, "ln"), back);
      const blip = path(pic, [P, "blipFill"], [A, "blip"]);
      const rId = blip && blip.getAttributeNS(R, "embed");
      const id = path(pic, [P, "nvPicPr"], [P, "cNvPr"]).getAttribute("id");
      const duo = blip && child(blip, A, "duotone"), key = blip && child(blip, A, "clrChange");
      const raw = env.images && env.images.get(rId);
      let s = raw && raw.stats, op = null;
      if (raw && raw.rgba) {
        const duoColors = duo && elements(duo).map(was);
        const keyColor = key && colorIn(child(key, A, "clrFrom")) && was(colorIn(child(key, A, "clrFrom")));
        s = imageStats(raw.rgba, raw.width, raw.height, {
          duotone: duoColors && duoColors.length === 2 && duoColors.every(Boolean) ? duoColors.map((c) => c.rgb) : null,
          transparent: keyColor ? keyColor.rgb : null });
      }
      if (derived) op = listed[id] || null;
      else if (s && shouldInvert(s, toDark) && !descendants(blip, "*", "svgBlip").length) op = duo && elements(duo).length === 2 ? "duo" : "px";
      if (op === "duo" && duo && elements(duo).length === 2) duo.appendChild(elements(duo)[0]);   // swap dark and light
      else if (op === "px") {
        pixels.push(blip);
        const from = key && colorIn(child(key, A, "clrFrom"));
        if (from && from.localName === "srgbClr") from.setAttribute("val", hexOf(flipLightness(rgbOf(from.getAttribute("val")))));
      } else op = null;
      if (op) { stats.pictures++; if (!derived) pictures[id] = op; }
      if (frame && s && s.clear < 0.5) painted.push({ ...frame, back: { old: s.mean, new: op ? flipLightness(s.mean.map(Math.round)) : s.mean } });
    }

    function shape(sp, frameIn) {
      const spPr = child(sp, P, "spPr"), style = child(sp, P, "style");
      const own = spPr && frameOf(child(spPr, A, "xfrm"));
      let frame = own ? frameIn(own) : null;
      if (!frame) { // a placeholder with no place of its own stands where the layout or the master puts it
        const ph = path(sp, [P, "nvSpPr"], [P, "nvPr"], [P, "ph"]);
        for (const doc of ph ? [env.layout, env.master] : []) {
          const f = frameOf(path(matchingPlaceholder(doc, ph), [P, "spPr"], [A, "xfrm"]));
          if (f) { frame = f; break; }
        }
      }
      const back = behind(frame);
      const hasText = descendants(child(sp, P, "txBody"), A, "t").some((t) => t.textContent.trim());
      const role = hasText ? "marker" : "graphic";
      let inside = spPr && fill(spPr, back, role);
      const fillRef = style && child(style, A, "fillRef");
      if (!inside && fillRef && !(spPr && hasFill(spPr)) && fillRef.getAttribute("idx") !== "0" && colorIn(fillRef)) {
        const d = decide(colorIn(fillRef), role);
        if (d && d.old.alpha >= 0.12) inside = blend(d, back);
      }
      const ln = spPr && child(spPr, A, "ln"), lnRef = style && child(style, A, "lnRef");
      if (ln && (child(ln, A, "solidFill") || child(ln, A, "gradFill"))) line(ln, back);
      else if (lnRef && !(ln && child(ln, A, "noFill")) && lnRef.getAttribute("idx") !== "0" && colorIn(lnRef)) decide(colorIn(lnRef), "ink", back);
      for (const glow of descendants(spPr && child(spPr, A, "effectLst"), A, "glow")) if (colorIn(glow)) decide(colorIn(glow), "surface");
      const fontRef = style && child(style, A, "fontRef");
      if (fontRef && colorIn(fontRef)) decide(colorIn(fontRef), "ink", inside || back);
      text(child(sp, P, "txBody"), sp, inside || back, false);
      if (inside && frame) painted.push({ ...frame, back: inside });
    }

    function table_(frameEl, frame, back) {
      for (const tc of descendants(frameEl, A, "tc")) {
        const tcPr = child(tc, A, "tcPr");
        const inside = tcPr && fill(tcPr, back);
        for (const side of elements(tcPr).filter((e) => /^ln/.test(e.localName))) line(side, back);
        const body = child(tc, A, "txBody");
        if (body) text(body, tc, inside || back, true);
      }
    }

    function tree(parent, frameIn) {
      for (const el of elements(parent)) {
        if (el.localName === "sp" || el.localName === "cxnSp") shape(el, frameIn);
        else if (el.localName === "pic") {
          const own = frameOf(path(el, [P, "spPr"], [A, "xfrm"]));
          const frame = own && frameIn(own);
          picture(el, frame, behind(frame));
        } else if (el.localName === "grpSp") tree(el, groupTransform(el, frameIn));
        else if (el.localName === "graphicFrame") {
          const own = frameOf(child(el, P, "xfrm"));
          const frame = own && frameIn(own);
          table_(el, frame, behind(frame));
        } else if (el.localName === "AlternateContent" || el.localName === "Choice" || el.localName === "Fallback") tree(el, frameIn);
      }
    }
    tree(child(cSld, P, "spTree"), (f) => f);

    // --- the color map ---
    {
      const backToMaster = MAP_KEYS.every((k) => mapNew[k] === masterMap[k]) && !tag.ovr;
      const next = slide.createElementNS(A, backToMaster ? "a:masterClrMapping" : "a:overrideClrMapping");
      if (!backToMaster) for (const k of MAP_KEYS) next.setAttribute(k, mapNew[k]);
      for (const old of elements(ovr)) ovr.removeChild(old);
      ovr.appendChild(next);
    }

    const next = { v: 1 };
    if (!derived) { next.o = 1; if (override) next.ovr = 1; if (linksMoved) next.h = [mapOld.hlink, mapOld.folHlink]; }
    if (Object.keys(out).length) next.m = out;
    if (Object.keys(pictures).length) next.p = pictures;
    return { tag: next.o || next.m ? next : null, pixels, toDark, ...stats };
  }

  // ---------- the slide inside its .pptx ----------
  const REL = "http://schemas.openxmlformats.org/package/2006/relationships";
  const IMAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
  const relsOf = (part) => part.replace(/([^/]+)$/, "_rels/$1.rels");
  function resolveTarget(part, target) {
    const out = part.split("/").slice(0, -1);
    for (const step of target.split("/")) { if (step === "..") out.pop(); else if (step !== ".") out.push(step); }
    return out.join("/");
  }
  // Toggles the slide at `slidePath` of a JSZip package, pictures included.
  // io: { parse, serialize, decodeImage(bytes, ext) → {rgba,width,height} | null, encodePng(rgba, w, h) → bytes }
  // The layout, master and theme a slide of a JSZip package lives in.
  async function slideEnv(zip, slidePath, io) {
    const read = async (p) => io.parse(await zip.file(p).async("string"));
    const related = async (part, type) => {
      const rels = await read(relsOf(part));
      const rel = elements(rels.documentElement).find((r) => r.getAttribute("Type").endsWith(`/${type}`));
      return resolveTarget(part, rel.getAttribute("Target"));
    };
    const layoutPath = await related(slidePath, "slideLayout"), masterPath = await related(layoutPath, "slideMaster");
    return { layout: await read(layoutPath), master: await read(masterPath), theme: await read(await related(masterPath, "theme")) };
  }

  async function isSlideDarkInPackage(zip, slidePath, io) {
    return isSlideDark(io.parse(await zip.file(slidePath).async("string")), await slideEnv(zip, slidePath, io));
  }

  async function invertSlideInPackage(zip, slidePath, tag, io) {
    const read = async (p) => io.parse(await zip.file(p).async("string"));
    const slide = await read(slidePath), rels = await read(relsOf(slidePath));
    const relById = new Map(elements(rels.documentElement).map((r) => [r.getAttribute("Id"), r]));

    // Every picture of the slide, decoded once per file.
    const decoded = new Map(), images = new Map();
    for (const blip of descendants(slide, A, "blip")) {
      const rId = blip.getAttributeNS(R, "embed"), rel = relById.get(rId);
      if (!rel || rel.getAttribute("TargetMode") === "External" || images.has(rId)) continue;
      const file = resolveTarget(slidePath, rel.getAttribute("Target")), ext = file.split(".").pop().toLowerCase();
      if (!["png", "jpg", "jpeg"].includes(ext) || !zip.file(file)) continue;
      if (!decoded.has(file)) decoded.set(file, await io.decodeImage(await zip.file(file).async("uint8array"), ext).catch(() => null));
      if (decoded.get(file)) images.set(rId, { ...decoded.get(file), file });
    }

    const result = invertSlide(slide, { ...(await slideEnv(zip, slidePath, io)), tag, images, parse: io.parse, serialize: io.serialize });
    if (result.skipped) return result;

    // The flipped pictures become new PNG parts; the originals are not touched.
    const flipped = new Map();
    let nextId = Math.max(0, ...[...relById.keys()].map((id) => Number((id.match(/\d+$/) || [0])[0]))) + 1;
    for (const blip of result.pixels) {
      const rId = blip.getAttributeNS(R, "embed"), image = images.get(rId);
      if (!image) throw new Error("A picture to flip back is in a format the pane can't read.");
      if (!flipped.has(image.file)) {
        let name, n = 0;
        do name = `ppt/media/flipped-${slidePath.match(/(\d+)\.xml$/)[1]}-${++n}.png`; while (zip.file(name));
        zip.file(name, io.encodePng(invertPixels(new Uint8Array(image.rgba)), image.width, image.height));
        const rel = rels.createElementNS(REL, "Relationship");
        rel.setAttribute("Id", `rId${nextId++}`); rel.setAttribute("Type", IMAGE);
        rel.setAttribute("Target", `../media/${name.split("/").pop()}`);
        rels.documentElement.appendChild(rel);
        flipped.set(image.file, rel.getAttribute("Id"));
      }
      blip.setAttributeNS(R, "r:embed", flipped.get(image.file));
    }
    const xml = io.serialize(slide);
    for (const [id, rel] of relById) {
      if (rel.getAttribute("Type") === IMAGE && !xml.includes(`"${id}"`)) rels.documentElement.removeChild(rel);
    }
    if (result.pixels.length) {
      const types = await read("[Content_Types].xml");
      if (!elements(types.documentElement).some((e) => e.localName === "Default" && e.getAttribute("Extension").toLowerCase() === "png")) {
        const png = types.createElementNS(types.documentElement.namespaceURI, "Default");
        png.setAttribute("Extension", "png"); png.setAttribute("ContentType", "image/png");
        types.documentElement.insertBefore(png, types.documentElement.firstChild);
        zip.file("[Content_Types].xml", io.serialize(types));
      }
    }
    const declared = (s) => (s.startsWith("<?xml") ? s : `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${s}`);
    zip.file(slidePath, declared(xml));
    zip.file(relsOf(slidePath), declared(io.serialize(rels)));
    return result;
  }

  const api = { invertSlide, invertSlideInPackage, isSlideDark, isSlideDarkInPackage, marker, invertPixels, flipLightness, imageStats, shouldInvert, surface, ink, contrast, luminance, toOklch, fromOklch };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.Dark = api;
})(this);
