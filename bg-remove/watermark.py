"""🧽 Watermark removal for the pane's "Remove watermark": Claude finds it, LaMa paints it out.

    uv run --project bg-remove python bg-remove/watermark.py IN.png OUT.png
    → stdout: {"ok": true, "found": [31, 4], "ms": …} | {"ok": false, "error": "…"}

Claude can see a picture but can't draw one, so the work is split:
1. Locate — the picture goes to Claude (`claude -p`, the subscription, never an API key) in
   overlapping tiles, each scaled up to 1024 px: on the whole picture at once it found the long
   runs of text and missed most of the small fragments on the sides; 2×2 tiles found nearly all.
   macOS's own OCR (Vision) was tried first and read nothing on gray text over black art.
2. Mask — Claude's boxes are loose, and painting a whole box out erased the art under it (a long
   box over a barbed wire left a gap LaMa never closed). So inside each box the mask keeps only
   thin strokes that stand out from what's around them (a grayscale top-hat, lighter or darker,
   whichever the box shows more of than its surroundings), dilated over the letters' halo.
   Boxes Claude calls a logo are masked whole.
3. Inpaint — big-lama (TorchScript, MPS), one crop per cluster of holes. Color and alpha are
   painted separately, so a watermark over a transparent cut-out heals in both.
Then once more on the result: what a second look still finds is what the first one missed.
"""
import base64
import concurrent.futures as cf
import hashlib
import io
import json
import os
import shutil
import subprocess
import sys
import time
import urllib.request
from pathlib import Path

import numpy as np
import torch
import torch.nn.functional as F
from PIL import Image
from scipy import ndimage

PASSES = 2
TILE = 1024            # px Claude gets per tile
MODEL = "opus"
EFFORT = "low"         # 9 s a tile against 15 s at medium, with the same boxes (2026-11-03)
CLAUDE_TIMEOUT_S = 180
LAMA_URL = "https://github.com/enesmsahin/simple-lama-inpainting/releases/download/v0.1.0/big-lama.pt"
LAMA_SHA256 = "7ba7aa7ac37a4d41fdbbeba3a2af7ead18058552997e3a3cd1a3b2210c9e6b4c"
DEBUG = Path(os.environ["WATERMARK_DEBUG"]) if os.environ.get("WATERMARK_DEBUG") else None
LAMA_PATH = Path.home() / "Library/Caches/powerpoint-addons/big-lama.pt"

SCHEMA = {
    "type": "object",
    "properties": {"watermarks": {"type": "array", "items": {
        "type": "object",
        "properties": {
            "kind": {"type": "string", "enum": ["text", "logo"]},
            "letter_height": {"type": "number"},
            "box": {"type": "array", "items": {"type": "number"}, "minItems": 4, "maxItems": 4},
        },
        "required": ["kind", "box", "letter_height"], "additionalProperties": False,
    }}},
    "required": ["watermarks"], "additionalProperties": False,
}
SYSTEM = "You locate watermarks in pictures so they can be painted out."
PROMPT = (
    "Find every watermark in this picture: text or logos a stock site or its author laid over the art "
    "(e.g. 'shutterstock', 'SVG VECTOR MONSTER', 'Dreamstime', semi-transparent tiled marks), even partial "
    "fragments of a few letters, at any angle. Not the picture's own content, not text that belongs to it. "
    "Return one tight box per watermark instance, in pixels of this {w}x{h} image, as [x0,y0,x1,y1]. "
    "kind is 'text' for lettering, 'logo' for a symbol or badge; letter_height is the height of its capital "
    "letters in pixels (for rotated text, measured across the line, not the box). If there is none, return an empty list."
)


def claude_binary():
    for candidate in (Path.home() / ".local/bin/claude", shutil.which("claude")):
        if candidate and os.access(candidate, os.X_OK):
            return str(candidate)
    raise RuntimeError("claude CLI not found (~/.local/bin/claude)")


def ask_claude(tile):
    buf = io.BytesIO()
    tile.save(buf, "PNG")
    message = {"type": "user", "message": {"role": "user", "content": [
        {"type": "image", "source": {"type": "base64", "media_type": "image/png",
                                     "data": base64.b64encode(buf.getvalue()).decode()}},
        {"type": "text", "text": PROMPT.format(w=tile.width, h=tile.height)},
    ]}}
    # Lean flags: no tools, no MCP, no user settings — a default `claude -p` spends ~37k tokens
    # just loading them. The subscription's OAuth login still applies.
    cmd = [claude_binary(), "-p", "--model", MODEL, "--effort", EFFORT,
           "--input-format", "stream-json", "--output-format", "stream-json", "--verbose",
           "--json-schema", json.dumps(SCHEMA), "--tools", "", "--strict-mcp-config",
           "--setting-sources", "", "--no-session-persistence", "--system-prompt", SYSTEM]
    run = subprocess.run(cmd, input=json.dumps(message) + "\n", capture_output=True, text=True,
                         timeout=CLAUDE_TIMEOUT_S)
    for line in run.stdout.splitlines():
        try:
            event = json.loads(line)
        except json.JSONDecodeError:
            continue
        if event.get("type") == "result":
            if event.get("is_error") or not event.get("structured_output"):
                raise RuntimeError(f"Claude: {event.get('result') or 'no answer'}")
            return event["structured_output"]["watermarks"]
    raise RuntimeError(f"Claude gave no result: {(run.stderr or run.stdout)[-400:]}")


def locate(rgb):
    """Boxes (kind, x0, y0, x1, y1, letter height) in pixels of `rgb`, from overlapping tiles asked in parallel."""
    w, h = rgb.size
    n = 1 if max(w, h) <= 600 else 2 if max(w, h) <= 1600 else 3
    pad = int(0.05 * max(w, h)) if n > 1 else 0
    tiles = []
    for j in range(n):
        for i in range(n):
            x0, y0 = max(0, i * w // n - pad), max(0, j * h // n - pad)
            x1, y1 = min(w, (i + 1) * w // n + pad), min(h, (j + 1) * h // n + pad)
            crop = rgb.crop((x0, y0, x1, y1))
            s = TILE / max(crop.size)
            tiles.append((x0, y0, s, crop.resize((round(crop.width * s), round(crop.height * s)), Image.LANCZOS)))
    with cf.ThreadPoolExecutor(len(tiles)) as pool:
        answers = list(pool.map(lambda t: ask_claude(t[3]), tiles))
    boxes = []
    for (x0, y0, s, _), found in zip(tiles, answers):
        for mark in found:
            bx0, by0, bx1, by1 = mark["box"]
            letters = mark.get("letter_height")
            boxes.append((mark["kind"], x0 + min(bx0, bx1) / s, y0 + min(by0, by1) / s,
                          x0 + max(bx0, bx1) / s, y0 + max(by0, by1) / s, letters / s if letters else None))
    return boxes


def dilate(t, k):
    """Max over a k×k square, as two 1-D passes: one 31×31 pass took ~50 s on a tiled watermark."""
    t = F.max_pool2d(t, (1, k), 1, (0, k // 2))
    return F.max_pool2d(t, (k, 1), 1, (k // 2, 0))


def erode(t, k):
    return -dilate(-t, k)


def odd(x, lo, hi):
    return int(np.clip(round(x), lo, hi)) | 1


def hole_mask(gray, boxes):
    """True where the watermark's own pixels are, inside the boxes.

    Inside each text box: thin strokes lighter or darker than what's around them (grayscale
    top-hat), minus any that run on outside the box — that's the art passing under it. Of what's
    left, only the tone the watermark has in most boxes: a dark line of the art cut in pieces by
    white letters crossing it would otherwise pass for letters (a hand's outline under a diagonal
    "shutterstock" was erased that way). Logos are masked whole.
    """
    H, W = gray.shape
    g_all = torch.from_numpy(gray)[None, None]
    hole = np.zeros((H, W), bool)
    pieces = []   # (region slices, inside-box mask, pieces mask, halo) per text box, tone later
    tones = []    # (mean gray, pixel count) of every candidate piece, across all boxes
    for kind, x0, y0, x1, y1, letters in boxes:
        size = max(2.0, min(letters or min(x1 - x0, y1 - y0), x1 - x0, y1 - y0))
        m = int(max(3, 0.25 * size))
        X0, Y0, X1, Y1 = max(0, int(x0) - m), max(0, int(y0) - m), min(W, int(x1) + m + 1), min(H, int(y1) + m + 1)
        if X1 <= X0 or Y1 <= Y0:
            continue
        if kind == "logo":
            hole[Y0:Y1, X0:X1] = True
            continue
        r = int(max(size, 8))
        RX0, RY0, RX1, RY1 = max(0, X0 - r), max(0, Y0 - r), min(W, X1 + r), min(H, Y1 + r)
        g = g_all[..., RY0:RY1, RX0:RX1]
        k = odd(size / 2.5, 3, 31)
        lighter = (g - dilate(erode(g, k), k))[0, 0].numpy()
        darker = (erode(dilate(g, k), k) - g)[0, 0].numpy()
        inside = np.zeros(g.shape[2:], bool)
        inside[Y0 - RY0:Y1 - RY0, X0 - RX0:X1 - RX0] = True
        region = gray[RY0:RY1, RX0:RX1]
        labels = np.zeros(inside.shape, np.int32)
        count = 0
        for th in (lighter, darker):
            found = th > max(0.05, 0.3 * np.percentile(th[inside], 99))
            parts, n = ndimage.label(found, structure=np.ones((3, 3)))
            leaving = np.unique(parts[~inside & found])
            keep = found & inside & ~np.isin(parts, leaving)
            labels[keep] = parts[keep] + count
            count += n
        ids = np.unique(labels[labels > 0])
        if not len(ids):
            continue
        means = ndimage.mean(region, labels, ids)
        sizes = ndimage.sum(np.ones_like(region), labels, ids)
        tones += list(zip(means, sizes))
        pieces.append(((slice(RY0, RY1), slice(RX0, RX1)), inside, labels, ids, means, odd(size * 0.55, 5, 31)))
    if not tones:
        return hole
    # The watermark's tone: the most painted 0.05-wide band of gray across all candidate pieces.
    bands = np.zeros(21)
    for mean, n in tones:
        bands[int(round(mean * 20))] += n
    tone = np.argmax(bands) / 20
    for (ys, xs), inside, labels, ids, means, halo in pieces:
        strokes = np.isin(labels, ids[np.abs(means - tone) <= 0.12])
        # The letters' anti-aliased halo is wider than the strokes: a tight mask left ghosts of
        # every letter, which LaMa then copied back in (7 px at 12 px text was the cleanest).
        strokes = dilate(torch.from_numpy(strokes.astype(np.float32))[None, None], halo)[0, 0].numpy() > 0
        hole[ys, xs] |= strokes & inside
    return hole


def ensure_lama():
    if LAMA_PATH.exists():
        return LAMA_PATH
    LAMA_PATH.parent.mkdir(parents=True, exist_ok=True)
    tmp = LAMA_PATH.with_suffix(".part")
    urllib.request.urlretrieve(LAMA_URL, tmp)
    digest = hashlib.sha256(tmp.read_bytes()).hexdigest()
    if digest != LAMA_SHA256:
        tmp.unlink()
        raise RuntimeError(f"big-lama.pt checksum mismatch ({digest})")
    tmp.rename(LAMA_PATH)
    return LAMA_PATH


class Lama:
    def __init__(self):
        self.device = "mps" if torch.backends.mps.is_available() else "cpu"
        self.model = torch.jit.load(str(ensure_lama()), map_location="cpu").eval().to(self.device)

    def paint(self, img3, hole):
        """`img3` H×W×3 in 0..1, `hole` H×W bool → img3 with the holes painted, one crop per cluster."""
        out = img3.copy()
        H, W = hole.shape
        clusters, _ = ndimage.label(ndimage.binary_dilation(hole, iterations=24))
        for ys, xs in ndimage.find_objects(clusters):
            c = 48   # context LaMa sees around the holes
            Y0, Y1, X0, X1 = max(0, ys.start - c), min(H, ys.stop + c), max(0, xs.start - c), min(W, xs.stop + c)
            m = hole[Y0:Y1, X0:X1]
            if not m.any():
                continue
            x = torch.from_numpy(np.ascontiguousarray(img3[Y0:Y1, X0:X1])).permute(2, 0, 1)[None]
            mt = torch.from_numpy(m.astype(np.float32))[None, None]
            h, w = m.shape
            ph, pw = (-h) % 8, (-w) % 8
            x = F.pad(x, (0, pw, 0, ph), mode="reflect")
            mt = F.pad(mt, (0, pw, 0, ph))
            with torch.inference_mode():
                y = self.model(x.to(self.device), mt.to(self.device))[0, :, :h, :w].permute(1, 2, 0).cpu().numpy()
            # LaMa tints what it paints into a black-and-white picture (a green dot in a black knot):
            # where everything around the hole is gray, the paint stays gray.
            around = img3[Y0:Y1, X0:X1][~m]
            if np.percentile(around.max(axis=1) - around.min(axis=1), 99) < 0.06:
                y = np.repeat(y.mean(axis=2, keepdims=True), 3, axis=2)
            out[Y0:Y1, X0:X1] = np.where(m[..., None], np.clip(y, 0, 1), out[Y0:Y1, X0:X1])
        return out


def log(msg):
    print(f"{time.strftime('%H:%M:%S')} {msg}", file=sys.stderr, flush=True)


def remove(rgba, lama):
    """One pass: returns (new RGBA float array, boxes found)."""
    alpha = rgba[..., 3:4]
    comp = rgba[..., :3] * alpha + (1 - alpha)          # over white, as Claude and the slide see it
    shown = Image.fromarray((comp * 255).round().astype(np.uint8), "RGB")
    t = time.time()
    boxes = locate(shown)
    log(f"Claude: {len(boxes)} box(es) in {time.time() - t:.1f} s")
    if not boxes:
        return rgba, boxes
    hole = hole_mask((comp @ np.array([0.299, 0.587, 0.114], np.float32)).astype(np.float32), boxes)
    if DEBUG:   # WATERMARK_DEBUG=<dir>: the boxes and the mask of every pass, to see what went wrong
        n = len(list(DEBUG.glob("boxes-*.json")))
        (DEBUG / f"boxes-{n}.json").write_text(json.dumps(boxes))
        Image.fromarray((hole * 255).astype(np.uint8)).save(DEBUG / f"mask-{n}.png")
    if not hole.any():
        return rgba, boxes
    t = time.time()
    painted = lama.paint(comp, hole)
    if (alpha > 0.999).all():
        new_alpha = alpha                                   # opaque picture: nothing to heal in alpha
    else:
        a = lama.paint(np.repeat(alpha, 3, axis=2), hole).mean(axis=2, keepdims=True).clip(0, 1)
        new_alpha = np.where(hole[..., None], a, alpha)
    log(f"LaMa: {hole.sum()} px in {time.time() - t:.1f} s")
    # Undo the white backdrop: color = (shown − (1 − α)) / α.
    rgb = np.clip((painted - (1 - new_alpha)) / np.maximum(new_alpha, 1e-3), 0, 1)
    return np.dstack([np.where(hole[..., None], rgb, rgba[..., :3]), new_alpha]), boxes


def main():
    started = time.time()
    src, dst = sys.argv[1], sys.argv[2]
    rgba = np.asarray(Image.open(src).convert("RGBA")).astype(np.float32) / 255
    lama = Lama()
    found = []
    for _ in range(PASSES):
        rgba, boxes = remove(rgba, lama)
        found.append(len(boxes))
        if not boxes:
            break
    Image.fromarray((rgba * 255).round().astype(np.uint8), "RGBA").save(dst)
    print(json.dumps({"ok": True, "found": found, "ms": round((time.time() - started) * 1000)}))


if __name__ == "__main__":
    try:
        main()
    except Exception as e:  # the caller reads one JSON line, whatever happened
        print(json.dumps({"ok": False, "error": f"{type(e).__name__}: {e}"}))
        sys.exit(1)
