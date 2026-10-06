"""✂️ BiRefNet background removal, kept warm by server.js for the pane's
"Remove background" (moved here from Victor Addons' ⌘⇧V bezel, 2026-10-06).

Line protocol on stdin/stdout, one JSON object per line:
  startup  → {"ready": true, "load_ms": …}
  request  ← {"in": "/path/picture.png", "out": "/path/cutout.png"}
  answer   → {"ok": true, "ms": …, "box": [x0, y0, x1, y1], "size": [w, h]}
           | {"ok": false, "error": "…"}
  `box` is where the trimmed cut-out sat in the input (x1/y1 exclusive), so
  the caller can put it back exactly over the subject.

BiRefNet on the Apple GPU (MPS, fp16): ~0.6 s per image warm on the M1 Max,
against 11 s for the same weights on the CPU through onnxruntime, and a CoreML
provider that had not finished compiling the graph after 10 minutes
(2026-10-05). Six engines were compared on the same pictures; BiRefNet
general with the edge bleed below was the clear winner on the zoomed edges.
"""
import json
import sys
import time

import numpy as np
import torch
import torch.nn.functional as F
from PIL import Image
from transformers import AutoModelForImageSegmentation

MODEL = "ZhengPeng7/BiRefNet"
# The model's Python comes from the Hub (`trust_remote_code`): pin the commit
# that was read and measured, so a push upstream cannot change what runs here.
REVISION = "e2bf8e4460fc8fa32bba5ea4d94b3233d367b0e4"
SIZE = 1024
MEAN = torch.tensor([0.485, 0.456, 0.406]).view(3, 1, 1)
STD = torch.tensor([0.229, 0.224, 0.225]).view(3, 1, 1)


def say(obj):
    sys.stdout.write(json.dumps(obj) + "\n")
    sys.stdout.flush()


def bleed(rgb, alpha, solid=0.95):
    """Edge colour decontamination: every pixel that is not solid foreground
    takes the colour of the solid foreground nearest to it (a push-pull
    pyramid of alpha-weighted averages). The models copy the ORIGINAL pixel
    into the soft edge, background included — on a black background that is a
    dark rim around the cut-out on any light page. Alpha itself is untouched."""
    w = (alpha > solid).astype(np.float32)
    levels = [np.dstack([rgb * w[..., None], w])]
    while min(levels[-1].shape[:2]) > 2:
        x = levels[-1]
        h, wd = x.shape[:2]
        x = np.pad(x, ((0, h % 2), (0, wd % 2), (0, 0)), mode="edge")
        levels.append(x.reshape(x.shape[0] // 2, 2, x.shape[1] // 2, 2, 4).mean(axis=(1, 3)))
    filled = levels[-1]
    for finer in reversed(levels[:-1]):
        up = np.dstack([np.asarray(Image.fromarray(filled[..., c].astype(np.float32)).resize(
            (finer.shape[1], finer.shape[0]), Image.BILINEAR)) for c in range(4)])
        cover = np.clip(finer[..., 3:4], 0, 1)
        filled = finer + (1 - cover) * up
    colour = filled[..., :3] / np.maximum(filled[..., 3:4], 1e-6)
    out = np.where((alpha > solid)[..., None], rgb, colour)
    # No solid pixel anywhere: nothing to borrow from, keep the original.
    return out if w.any() else rgb


def trim(rgba, visible=8):
    """Crop to the bounding box of the pixels that are actually there: the
    cut-out sits on the original canvas, and a subject in a corner pasted
    with a screenful of transparent margin around it. `visible` (of 255)
    ignores the faint haze a model leaves where the background was.
    Returns the crop and its box in the input, [x0, y0, x1, y1]."""
    ys, xs = np.nonzero(rgba[..., 3] > visible)
    if ys.size == 0:
        return rgba, [0, 0, rgba.shape[1], rgba.shape[0]]
    box = [int(xs.min()), int(ys.min()), int(xs.max()) + 1, int(ys.max()) + 1]
    return rgba[box[1]:box[3], box[0]:box[2]], box


def main():
    started = time.time()
    device = "mps" if torch.backends.mps.is_available() else "cpu"
    dtype = torch.float16 if device == "mps" else torch.float32
    model = AutoModelForImageSegmentation.from_pretrained(
        MODEL, revision=REVISION, trust_remote_code=True).to(device).eval().to(dtype)
    # First inference on MPS compiles kernels (~4 s) — pay it now, not on the click.
    with torch.no_grad():
        model(torch.zeros(1, 3, SIZE, SIZE, device=device, dtype=dtype))
    say({"ready": True, "load_ms": int((time.time() - started) * 1000), "device": device})

    for line in sys.stdin:
        if not line.strip():
            continue
        t = time.time()
        try:
            req = json.loads(line)
            image = Image.open(req["in"]).convert("RGB")
            x = torch.from_numpy(np.array(image.resize((SIZE, SIZE), Image.BILINEAR))).permute(2, 0, 1)
            x = ((x.float() / 255 - MEAN) / STD).unsqueeze(0).to(device, dtype)
            with torch.no_grad():
                pred = model(x)[-1].sigmoid().float()
                pred = F.interpolate(pred, size=(image.height, image.width), mode="bilinear", align_corners=False)
            alpha = pred[0, 0].clamp(0, 1).cpu().numpy()
            rgb = np.asarray(image).astype(np.float32) / 255
            out = (np.dstack([bleed(rgb, alpha), alpha]) * 255).round().astype(np.uint8)
            cut, box = trim(out)
            Image.fromarray(cut, "RGBA").save(req["out"])
            say({"ok": True, "ms": int((time.time() - t) * 1000), "box": box, "size": [image.width, image.height]})
        except Exception as e:  # one bad image must not kill the warm model
            say({"ok": False, "error": f"{type(e).__name__}: {e}"})


if __name__ == "__main__":
    main()
