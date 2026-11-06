// 🖼️ Finds the useless border of a picture: per side, the flat band (transparent, white, black
// or any one color) between the edge and the content. Pure pixels, no Office.js and no canvas,
// so test/crop.test.js runs it in Node.
(function (root) {
  // Widest spread per channel (0–255) inside a line that still counts as one color: enough for
  // JPEG noise, and a border stopped short by ringing beats content cut away.
  const TOLERANCE = 16;
  // PowerPoint's render can blur the outermost lines. Up to this many lines of another color
  // don't count as a band, so the real border behind them still goes.
  const RIM = 2;

  // Average color of a line of pixels, or null when the line isn't one color. Premultiplied:
  // a transparent pixel has no color, whatever RGB it happens to hide.
  function flatColor(rgba, start, step, count) {
    const low = [255, 255, 255, 255], high = [0, 0, 0, 0];
    for (let n = 0, p = start; n < count; n++, p += step) {
      const alpha = rgba[p + 3];
      for (let c = 0; c < 4; c++) {
        const value = c === 3 ? alpha : (rgba[p + c] * alpha) / 255;
        if (value < low[c]) low[c] = value;
        if (value > high[c]) high[c] = value;
      }
    }
    if (high.some((value, c) => value - low[c] > TOLERANCE)) return null;
    return high.map((value, c) => (value + low[c]) / 2);
  }

  const near = (a, b) => a.every((value, c) => Math.abs(value - b[c]) <= TOLERANCE);

  // Takes one more line off a side, unless that line is already the picture. A second flat band
  // is content (a header bar under the margin), and so is a slow drift (a sky, a gradient):
  // both stay, and only a real border is left to take.
  function peel(side) {
    const color = side.color();
    if (!color) return false;
    const newBand = !side.last || !near(color, side.last);
    if (newBand && side.last && side.lines > RIM) return false;
    if (!newBand && !near(color, side.band)) return false;
    if (newBand) { side.band = color; side.lines = 0; }
    side.last = color;
    side.lines++;
    side.shrink();
    return true;
  }

  // [left, top, right, bottom] of what's left in pixels, right and bottom exclusive;
  // null when nothing would be left: the whole picture is flat.
  function borderBox(rgba, width, height) {
    let left = 0, top = 0, right = width, bottom = height;
    const row = (y) => flatColor(rgba, (y * width + left) * 4, 4, right - left);
    const column = (x) => flatColor(rgba, (top * width + x) * 4, width * 4, bottom - top);
    const sides = [
      { color: () => row(top), shrink: () => { top++; } },
      { color: () => row(bottom - 1), shrink: () => { bottom--; } },
      { color: () => column(left), shrink: () => { left++; } },
      { color: () => column(right - 1), shrink: () => { right--; } },
    ];
    // Rounds, because a side can only go once its neighbors went: the top row is flat only
    // after a left border of another color is gone.
    for (let peeled = true; peeled;) {
      peeled = false;
      for (const side of sides) {
        while (left < right && top < bottom && peel(side)) peeled = true;
        if (left >= right || top >= bottom) return null;
      }
    }
    return [left, top, right, bottom];
  }

  const api = { borderBox };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.Crop = api;
})(this);
