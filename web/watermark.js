// 🧽 Removes a stock watermark ("SVG VECTOR MONSTER", "shutterstock"…) from black-and-white
// clip art. Such a watermark is gray text laid over the art: gray on the black strokes, light
// gray on the white background. So every mid-gray pixel whose neighborhood is (almost) all
// dark turns dark, and one whose neighborhood is (almost) all light turns light, each taking
// the average color of those neighbors. Anti-aliased edges sit between a dark and a light
// pixel, so they are left alone and the art keeps its smooth outline.
// Colors are judged on RGB alone, alpha untouched: on a transparent cut-out the soft edges
// live in alpha (RGB stays black), so they are never mistaken for gray.
// Pure function on an RGBA buffer, no DOM, so test/watermark.test.js runs it in Node.
(function (root) {
  const DARK = 45;        // luminance below this is art
  const LIGHT = 245;      // above this is background
  const VISIBLE = 32;     // pixels more transparent than this don't vote
  const MAJORITY = 0.75;  // share of the voting neighbors that must agree
  const EDGE = 2;         // px from the other side within which a pixel is an anti-aliased edge

  // Summed-area table of `values` (w × h), padded with a zero row and column.
  function integral(values, w, h) {
    const t = new Float64Array((w + 1) * (h + 1));
    for (let y = 0; y < h; y++) {
      let row = 0;
      for (let x = 0; x < w; x++) {
        row += values[y * w + x];
        t[(y + 1) * (w + 1) + x + 1] = t[y * (w + 1) + x + 1] + row;
      }
    }
    return t;
  }

  function boxSum(t, w, h, x, y, r) {
    const x0 = Math.max(0, x - r), y0 = Math.max(0, y - r);
    const x1 = Math.min(w, x + r + 1), y1 = Math.min(h, y + r + 1);
    const W = w + 1;
    return t[y1 * W + x1] - t[y0 * W + x1] - t[y1 * W + x0] + t[y0 * W + x0];
  }

  // 4-connected runs of light pixels (kind 2), each as a list of pixel indexes.
  function lightBlobs(kind, w, h) {
    const seen = new Uint8Array(w * h), blobs = [];
    for (let start = 0; start < w * h; start++) {
      if (kind[start] !== 2 || seen[start]) continue;
      const blob = [], stack = [start];
      seen[start] = 1;
      while (stack.length) {
        const i = stack.pop(), x = i % w;
        blob.push(i);
        for (const j of [x > 0 ? i - 1 : -1, x < w - 1 ? i + 1 : -1, i - w, i + w]) {
          if (j >= 0 && j < w * h && kind[j] === 2 && !seen[j]) { seen[j] = 1; stack.push(j); }
        }
      }
      blobs.push(blob);
    }
    return blobs;
  }

  // Rewrites `rgba` in place; returns how many pixels changed. `radius` is the half-size of the
  // neighborhood, which must be wider than the watermark's letters are tall (~7 px at 3 px/pt).
  function removeWatermark(rgba, w, h, radius = 6) {
    const n = w * h;
    const kind = new Uint8Array(n);   // 0 none, 1 dark, 2 light, 3 mid-gray
    const planes = { dark: new Float64Array(n), light: new Float64Array(n) };
    const sums = { dark: [0, 1, 2].map(() => new Float64Array(n)), light: [0, 1, 2].map(() => new Float64Array(n)) };
    for (let i = 0; i < n; i++) {
      const p = i * 4;
      if (rgba[p + 3] <= VISIBLE) continue;
      const lum = 0.299 * rgba[p] + 0.587 * rgba[p + 1] + 0.114 * rgba[p + 2];
      const k = lum < DARK ? "dark" : lum > LIGHT ? "light" : null;
      if (!k) { kind[i] = 3; continue; }
      kind[i] = k === "dark" ? 1 : 2;
      planes[k][i] = 1;
      for (let c = 0; c < 3; c++) sums[k][c][i] = rgba[p + c];
    }
    // White specks of the watermark inside the art aren't background: a light blob smaller than
    // the neighborhood is gray to repaint, so it neither votes nor shields its neighbors.
    for (const blob of lightBlobs(kind, w, h)) {
      if (blob.length >= (2 * radius + 1) ** 2) continue;
      for (const i of blob) {
        kind[i] = 3;
        planes.light[i] = 0;
        for (let c = 0; c < 3; c++) sums.light[c][i] = 0;
      }
    }
    const count = { dark: integral(planes.dark, w, h), light: integral(planes.light, w, h) };
    const color = {
      dark: sums.dark.map((s) => integral(s, w, h)),
      light: sums.light.map((s) => integral(s, w, h)),
    };
    const paint = new Uint8Array(n);   // 1 dark, 2 light
    for (let y = 0; y < h; y++) {
      for (let x = 0; x < w; x++) {
        const i = y * w + x;
        if (kind[i] !== 3) continue;
        const dark = boxSum(count.dark, w, h, x, y, radius), light = boxSum(count.light, w, h, x, y, radius);
        if (dark + light === 0) continue;
        // A gray pixel close to the background is the art's anti-aliased outline: leave it.
        // Gray turns light only with no art at all around it: near the art, the gray between
        // its outline and the background may just as well be the art's own shading.
        if (dark / (dark + light) >= MAJORITY && boxSum(count.light, w, h, x, y, EDGE) === 0) paint[i] = 1;
        else if (dark === 0) paint[i] = 2;
      }
    }
    // The letters' own anti-aliasing is dark enough to pass for art: take the 1 px ring around
    // what was repainted dark along with it.
    const painted = integral(paint.map((p) => (p === 1 ? 1 : 0)), w, h);
    for (let y = 0; y < h; y++) {
      for (let x = 0; x < w; x++) {
        const i = y * w + x;
        if (kind[i] === 1 && !paint[i] && boxSum(painted, w, h, x, y, 1) > 0) paint[i] = 1;
      }
    }
    let changed = 0;
    for (let y = 0; y < h; y++) {
      for (let x = 0; x < w; x++) {
        const i = y * w + x;
        if (!paint[i]) continue;
        const k = paint[i] === 1 ? "dark" : "light";
        const total = boxSum(count[k], w, h, x, y, radius);
        for (let c = 0; c < 3; c++) rgba[i * 4 + c] = Math.round(boxSum(color[k][c], w, h, x, y, radius) / total);
        changed++;
      }
    }
    return changed;
  }

  const api = { removeWatermark };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.Watermark = api;
})(this);
