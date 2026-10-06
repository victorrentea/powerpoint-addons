// ✂️ The BiRefNet sidecar: `uv run` of bg-remove/cutout_server.py, one JSON line
// in, one out, the model kept on the GPU between clicks.
//
// Warm, because cold is unusable: ~5 s to load and a ~4 s first inference (MPS
// compiling its kernels), then ~1 s per picture. So the pane asks for a
// `prewarm()` when the Image tab opens, and the process exits after IDLE_MS
// without a request — a few GB of torch and weights have no business sitting
// in memory all day for a button.
//
// Requests are chained one after another, so a click that lands while the
// model is still loading simply waits for it.
const { spawn } = require("child_process");
const fs = require("fs");
const os = require("os");
const path = require("path");
const readline = require("readline");

const UV = "/opt/homebrew/bin/uv";
const DIR = path.join(__dirname, "bg-remove");
const LOG = "/tmp/powerpoint-addons-bg-remove.log";
const IDLE_MS = 30 * 60 * 1000;
// A first-ever start also builds the venv (torch is ~1 GB of wheels).
const START_TIMEOUT_MS = 300 * 1000;
const REQUEST_TIMEOUT_MS = 60 * 1000;

let child = null;
let ready = null;      // Promise of the {"ready": true} hello, while `child` lives
let waiters = [];      // resolvers for the next JSON lines, in order
let chain = Promise.resolve();
let idleTimer = null;

function nextMessage(timeoutMs) {
  return new Promise((resolve, reject) => {
    const waiter = { resolve, reject };
    waiter.timer = setTimeout(() => {
      waiters = waiters.filter((w) => w !== waiter);
      reject(new Error(`BiRefNet did not answer in ${timeoutMs / 1000} s`));
    }, timeoutMs);
    waiters.push(waiter);
  });
}

function stop() {
  clearTimeout(idleTimer);
  if (child) {
    child.stdin.end();   // EOF ends the server's read loop
    child.kill();
  }
  child = null;
  ready = null;
}

function scheduleIdleStop() {
  clearTimeout(idleTimer);
  idleTimer = setTimeout(() => {
    if (!child) return;
    console.log(new Date().toISOString(), `✂️ BiRefNet stopped after ${IDLE_MS / 60000} min idle`);
    stop();
  }, IDLE_MS);
}

function ensureRunning() {
  if (ready) return ready;
  const p = spawn(UV, ["run", "-q", "--project", DIR, "python", "-u", path.join(DIR, "cutout_server.py")], {
    env: { ...process.env, PYTORCH_ENABLE_MPS_FALLBACK: "1" },
    // Warnings and tracebacks go to a file, not the void: a sidecar that dies
    // without saying why is the hardest bug to chase.
    stdio: ["pipe", "pipe", fs.openSync(LOG, "w")],
  });
  child = p;
  readline.createInterface({ input: p.stdout }).on("line", (line) => {
    let msg;
    try { msg = JSON.parse(line); } catch { return; }   // a library printing to stdout
    if (msg === null || typeof msg !== "object") return;
    const w = waiters.shift();
    if (w) { clearTimeout(w.timer); w.resolve(msg); }
  });
  p.on("exit", (code) => {
    if (child === p) { child = null; ready = null; }
    for (const w of waiters.splice(0)) {
      clearTimeout(w.timer);
      w.reject(new Error(`BiRefNet exited (${code}) — see ${LOG}`));
    }
  });
  p.on("error", (e) => console.error("✂️ could not start BiRefNet:", e.message));
  ready = nextMessage(START_TIMEOUT_MS).then((hello) => {
    if (hello.ready !== true) throw new Error(`BiRefNet failed to start — see ${LOG}`);
    console.log(new Date().toISOString(), `✂️ BiRefNet ready on ${hello.device} in ${Math.round(hello.load_ms / 1000)} s`);
    return hello;
  }, (e) => { stop(); throw e; });
  return ready;
}

function serial(job) {
  const run = chain.then(job, job);
  chain = run.catch(() => {});
  return run;
}

function prewarm() {
  return serial(async () => {
    try { await ensureRunning(); } finally { scheduleIdleStop(); }
  });
}

// PNG bytes in → { png, box: [x0, y0, x1, y1], size: [w, h], ms } out, `box`
// being where the trimmed cut-out sat in the input picture.
function cutOut(pngBytes) {
  return serial(async () => {
    try {
      await ensureRunning();
      const tmp = fs.mkdtempSync(path.join(os.tmpdir(), "cutout-"));
      try {
        const input = path.join(tmp, "in.png"), output = path.join(tmp, "out.png");
        fs.writeFileSync(input, pngBytes);
        const answer = nextMessage(REQUEST_TIMEOUT_MS);
        child.stdin.write(JSON.stringify({ in: input, out: output }) + "\n");
        let reply;
        try { reply = await answer; } catch (e) { stop(); throw e; }
        if (reply.ok !== true) throw new Error(`BiRefNet: ${reply.error || "failed"}`);
        return { png: fs.readFileSync(output), box: reply.box, size: reply.size, ms: reply.ms };
      } finally {
        fs.rmSync(tmp, { recursive: true, force: true });
      }
    } finally {
      scheduleIdleStop();
    }
  });
}

module.exports = { prewarm, cutOut, stop };
