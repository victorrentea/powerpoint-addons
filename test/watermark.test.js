// node --test — runs web/watermark.js on hand-drawn pixel grids, outside PowerPoint.
const test = require("node:test");
const assert = require("node:assert/strict");
const { removeWatermark } = require("../web/watermark.js");

// Builds an RGBA buffer from rows of characters: '#' black art, '.' white background,
// 'g' mid gray (the watermark), 'w' pure white, ' ' fully transparent.
const SHADE = { "#": [0, 0, 0, 255], ".": [255, 255, 255, 255], g: [128, 128, 128, 255],
  w: [255, 255, 255, 255], e: [40, 40, 40, 128], " ": [0, 0, 0, 0] };
function image(rows) {
  const h = rows.length, w = rows[0].length, rgba = new Uint8ClampedArray(w * h * 4);
  rows.forEach((row, y) => [...row].forEach((ch, x) => rgba.set(SHADE[ch], (y * w + x) * 4)));
  return { rgba, w, h };
}
const pixel = ({ rgba, w }, x, y) => [...rgba.slice((y * w + x) * 4, (y * w + x) * 4 + 4)];

// A thick black band across a transparent picture, with a word of gray and white letters on it.
const band = (letters) => [
  ...Array(8).fill(" ".repeat(30)),
  ...Array(5).fill("#".repeat(30)),
  "#".repeat(8) + letters + "#".repeat(22 - letters.length),
  ...Array(5).fill("#".repeat(30)),
  ...Array(8).fill(" ".repeat(30)),
];

test("gray and white letters on the black art turn black, alpha kept", () => {
  const img = image(band("gwggwg"));
  assert.ok(removeWatermark(img.rgba, img.w, img.h) >= 6);
  for (let x = 8; x < 14; x++) assert.deepEqual(pixel(img, x, 13), [0, 0, 0, 255]);
});

test("the soft alpha edge of a transparent cut-out is left alone", () => {
  const rows = band("gggg");
  rows[7] = "e".repeat(30);
  const img = image(rows);
  removeWatermark(img.rgba, img.w, img.h);
  for (let x = 0; x < 30; x++) assert.deepEqual(pixel(img, x, 7), [40, 40, 40, 128]);
  assert.deepEqual(pixel(img, 0, 0), [0, 0, 0, 0]);
});

test("gray watermark on the white background, away from the art, turns white", () => {
  const img = image([...Array(10).fill("."), "...gggg...", ...Array(10).fill(".")].map((r) => r.padEnd(30, ".")));
  removeWatermark(img.rgba, img.w, img.h);
  for (let x = 3; x < 7; x++) assert.deepEqual(pixel(img, x, 10), [255, 255, 255, 255]);
});

test("the gray outline between black art and white background stays", () => {
  const rows = [...Array(10).fill("#".repeat(30)), "g".repeat(30), ...Array(10).fill(".".repeat(30))];
  const img = image(rows);
  removeWatermark(img.rgba, img.w, img.h);
  for (let x = 0; x < 30; x++) assert.deepEqual(pixel(img, x, 10), [128, 128, 128, 255]);
});

test("a big white area inside the art is background, not a speck", () => {
  const rows = Array.from({ length: 40 }, (_, y) => (y < 5 || y >= 35 ? "#".repeat(40) : "#####" + ".".repeat(30) + "#####"));
  const img = image(rows);
  removeWatermark(img.rgba, img.w, img.h);
  assert.deepEqual(pixel(img, 20, 20), [255, 255, 255, 255]);
});
