// 🧽 Runs bg-remove/watermark.py once per click: Claude (`claude -p`, the subscription) finds the
// watermark, LaMa paints it out — see the script's docstring for why it's split that way.
// Not kept warm like BiRefNet: torch and LaMa load in ~3 s, nothing next to the ~10 s per pass
// Claude takes to look, and a resident model would cost RAM all day for a rare click.
const { spawn } = require("child_process");
const fs = require("fs");
const os = require("os");
const path = require("path");

const UV = "/opt/homebrew/bin/uv";
const DIR = path.join(__dirname, "bg-remove");
const LOG = "/tmp/powerpoint-addons-watermark.log";
const TIMEOUT_MS = 6 * 60 * 1000;   // two passes of Claude + LaMa; a first run also downloads LaMa (200 MB)

let chain = Promise.resolve();

function run(input, output) {
  return new Promise((resolve, reject) => {
    const p = spawn(UV, ["run", "-q", "--project", DIR, "python", "-W", "ignore",
      path.join(DIR, "watermark.py"), input, output], {
      env: { ...process.env, PYTORCH_ENABLE_MPS_FALLBACK: "1",
        PATH: `${path.join(os.homedir(), ".local/bin")}:/opt/homebrew/bin:/usr/bin:/bin` },
      stdio: ["ignore", "pipe", fs.openSync(LOG, "a")],
    });
    let out = "";
    p.stdout.on("data", (c) => { out += c; });
    const timer = setTimeout(() => p.kill(), TIMEOUT_MS);
    p.on("error", reject);
    p.on("exit", (code) => {
      clearTimeout(timer);
      const line = out.trim().split("\n").pop() || "";
      let reply;
      try { reply = JSON.parse(line); } catch { return reject(new Error(`watermark.py exited (${code}) — see ${LOG}`)); }
      if (reply.ok !== true) return reject(new Error(reply.error || "watermark removal failed"));
      resolve(reply);
    });
  });
}

// PNG bytes in → { png, found: [boxes per pass], ms }. One at a time: each one runs Claude in parallel tiles.
function removeWatermark(pngBytes) {
  const job = async () => {
    const tmp = fs.mkdtempSync(path.join(os.tmpdir(), "watermark-"));
    try {
      const input = path.join(tmp, "in.png"), output = path.join(tmp, "out.png");
      fs.writeFileSync(input, pngBytes);
      const reply = await run(input, output);
      return { png: fs.readFileSync(output), found: reply.found, ms: reply.ms };
    } finally {
      fs.rmSync(tmp, { recursive: true, force: true });
    }
  };
  const result = chain.then(job, job);
  chain = result.catch(() => {});
  return result;
}

module.exports = { removeWatermark };
