// node --test — runs web/crop.js on hand-painted pixels, outside PowerPoint.
const test = require("node:test");
const assert = require("node:assert/strict");
const { borderBox, frameOfBox } = require("../web/crop.js");

const WHITE = [255, 255, 255, 255], BLACK = [0, 0, 0, 255], RED = [200, 30, 30, 255];
const BLUE = [20, 40, 200, 255], CLEAR = [0, 0, 0, 0];

// A w×h picture whose pixels come from paint(x, y).
function picture(w, h, paint) {
  const rgba = new Uint8ClampedArray(w * h * 4);
  for (let y = 0; y < h; y++) for (let x = 0; x < w; x++) rgba.set(paint(x, y), (y * w + x) * 4);
  return { rgba, w, h };
}
const box = ({ rgba, w, h }) => borderBox(rgba, w, h);

// Content that can't be mistaken for a flat line: a checkerboard of red and blue.
const checker = (x, y) => ((x + y) % 2 ? RED : BLUE);
const framed = (w, h, [l, t, r, b], outside) => picture(w, h, (x, y) =>
  (x >= l && x < r && y >= t && y < b ? checker(x, y) : outside(x, y)));

for (const [name, color] of [["white", WHITE], ["black", BLACK], ["transparent", CLEAR], ["red", RED]]) {
  test(`crops a ${name} border, each side its own width`, () => {
    assert.deepEqual(box(framed(40, 30, [3, 5, 32, 27], () => color)), [3, 5, 32, 27]);
  });
}

test("a picture with no border stays whole", () => {
  assert.deepEqual(box(picture(20, 10, checker)), [0, 0, 20, 10]);
});

test("a border on one side only", () => {
  assert.deepEqual(box(framed(20, 10, [6, 0, 20, 10], () => WHITE)), [6, 0, 20, 10]);
});

test("each side can have its own color", () => {
  const p = framed(30, 30, [4, 4, 26, 26], (x, y) => (x < 4 ? BLACK : y < 4 ? WHITE : x >= 26 ? RED : BLUE));
  assert.deepEqual(box(p), [4, 4, 26, 26]);
});

test("transparent pixels count as one color, whatever RGB they hide", () => {
  const p = framed(20, 20, [2, 2, 18, 18], (x, y) => [(x * 37) % 256, (y * 91) % 256, x * y % 256, 0]);
  assert.deepEqual(box(p), [2, 2, 18, 18]);
});

test("JPEG noise on the border doesn't stop the crop", () => {
  const noisy = (x, y) => { const v = 248 + ((x * 7 + y * 13) % 8); return [v, v, v, 255]; };
  assert.deepEqual(box(framed(40, 40, [5, 6, 35, 34], noisy)), [5, 6, 35, 34]);
});

test("a flat picture has nothing left to keep", () => {
  assert.equal(box(picture(10, 10, () => WHITE)), null);
  assert.equal(box(picture(10, 10, () => CLEAR)), null);
});

test("a thin diagonal line touching no edge is kept whole", () => {
  const p = picture(40, 40, (x, y) => (x === y && x >= 8 && x < 30 ? BLACK : WHITE));
  assert.deepEqual(box(p), [8, 8, 30, 30]);
});

test("rounded corners are not eaten: the corner pixels of the frame just stay", () => {
  const p = picture(20, 20, (x, y) => {
    const corner = (x < 2 || x > 17) && (y < 2 || y > 17);
    return corner ? CLEAR : x < 2 || y < 2 || x > 17 || y > 17 ? RED : checker(x, y);
  });
  assert.deepEqual(box(p), [0, 0, 20, 20]);
});

test("a sky that slowly changes color is the picture, not a border", () => {
  const sky = (x, y) => (y < 60 ? [100, 150, 150 + y, 255] : checker(x, y));
  const p = picture(40, 80, sky);
  const [, top] = box(p);
  assert.ok(top < 20, `ate ${top} lines of sky`);
});

test("a header bar under the margin stays: only the outer band goes", () => {
  const p = picture(40, 40, (x, y) => (y < 4 ? WHITE : y < 10 ? BLUE : checker(x, y)));
  assert.deepEqual(box(p), [0, 4, 40, 40]);
});

test("a blurred rim doesn't hide the border behind it", () => {
  const half = [255, 255, 255, 128];
  const p = framed(30, 30, [6, 6, 24, 24], (x, y) =>
    (x === 0 || y === 0 || x === 29 || y === 29 ? half : WHITE));
  assert.deepEqual(box(p), [6, 6, 24, 24]);
});

const near = (frame, expected) => Object.keys(expected).forEach((k) =>
  assert.ok(Math.abs(frame[k] - expected[k]) < 1e-9, `${k}: ${frame[k]} ≠ ${expected[k]}`));
const slab = { left: 100, top: 50, width: 200, height: 100 };

test("an upright picture keeps the part where it was", () => {
  near(frameOfBox(slab, [10, 20, 110, 70], [400, 200]), { left: 105, top: 60, width: 50, height: 25 });
});

test("a rotated picture keeps the part where it shows: its center turns around the frame's", () => {
  // Left half of a 200×100 frame centered at (200, 100), turned 90° clockwise: it now shows on top.
  near(frameOfBox({ ...slab, rotation: 90 }, [0, 0, 200, 200], [400, 200]),
    { left: 150, top: 0, width: 100, height: 100 });
  // The whole picture stays exactly the original frame, at any angle.
  near(frameOfBox({ ...slab, rotation: -17 }, [0, 0, 400, 200], [400, 200]), slab);
});
