// 🔤 Font bar: the font name and size of the selection, shown (and editable) in
// PowerPoint's title bar, right of the deck's name — Office for Mac lets only
// buttons into the Quick Access Toolbar, never the Font / Font Size boxes.
//
// Three pieces, this module is the hub between them:
//   pane (web/fontbar.js, inside PowerPoint)  — reads/sets the font via Office.js
//     POST /font/state  {pane, doc, name, size, sample}  whenever the selection's font or text changes
//     GET  /font/events?pane=…  (SSE)              ← {name} or {size} to apply
//   bar (font-bar/FontBar.swift, native panel)   — child process, JSON lines both ways
//     stdin  ← {panes: [{pane, doc, name, size, sample, at}]}
//     stdout → {pane, name} | {pane, size}
// Each deck with the pane open is one `pane`; the bar picks the one whose `doc`
// matches PowerPoint's front window, else the one that reported last.
const { spawn } = require("child_process");
const fs = require("fs");
const path = require("path");
const readline = require("readline");

const BIN = path.join(__dirname, "font-bar", ".build", "font-bar");
const LOG = "/tmp/powerpoint-addons-font-bar.log";
const RESPAWN_MS = 5000;

const panes = new Map();   // pane id → { pane, doc, name, size, at, stream }
let bar = null;

function broadcast() {
  if (!bar) return;
  const list = [...panes.values()].map(({ stream, ...state }) => state);
  bar.stdin.write(JSON.stringify({ panes: list }) + "\n");
}

function startBar() {
  if (!fs.existsSync(BIN)) {
    console.error(`🔤 ${BIN} missing — build it with font-bar/build.sh`);
    return;
  }
  const p = spawn(BIN, [], { stdio: ["pipe", "pipe", fs.openSync(LOG, "a")] });
  bar = p;
  p.stdin.on("error", () => {});   // the bar died mid-write; "exit" respawns it
  readline.createInterface({ input: p.stdout }).on("line", (line) => {
    let cmd;
    try { cmd = JSON.parse(line); } catch { return; }
    const target = panes.get(cmd.pane);
    if (!target) return;
    const { pane, ...change } = cmd;
    target.stream.write(`data: ${JSON.stringify(change)}\n\n`);
  });
  p.on("exit", (code, signal) => {
    if (bar === p) bar = null;
    console.error(new Date().toISOString(), `🔤 font bar exited (${code ?? signal}), restarting`);
    setTimeout(startBar, RESPAWN_MS).unref();
  });
  p.on("error", (e) => console.error("🔤 could not start the font bar:", e.message));
  broadcast();
}

// GET /font/events?pane=… — the pane's command stream; it also tells us the pane is alive.
function events(req, res, paneId) {
  res.writeHead(200, { "Content-Type": "text/event-stream", "Cache-Control": "no-store",
    Connection: "keep-alive" });
  res.write(": hello\n\n");
  const entry = panes.get(paneId) || { pane: paneId, doc: "", name: null, size: null, sample: null, at: Date.now() };
  entry.stream = res;
  panes.set(paneId, entry);
  // WKWebView drops an idle stream after a while; a comment every 20 s keeps it open.
  const keepAlive = setInterval(() => res.write(": ping\n\n"), 20000);
  req.on("close", () => {
    clearInterval(keepAlive);
    if (panes.get(paneId)?.stream === res) panes.delete(paneId);
    broadcast();
  });
}

// POST /font/state — {pane, doc, name, size, sample}
function state(body) {
  const { pane, doc, name, size, sample } = JSON.parse(body);
  const entry = panes.get(pane);
  if (!entry) return;   // no event stream yet: it registers the pane, and the pane re-reports on connect
  Object.assign(entry, { doc: doc || "", name: name ?? null, size: size ?? null, sample: sample ?? null, at: Date.now() });
  broadcast();
}

function stop() {
  if (bar) { const p = bar; bar = null; p.removeAllListeners("exit"); p.kill(); }
}

module.exports = { startBar, events, state, stop };
