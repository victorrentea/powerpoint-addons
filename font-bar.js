// 🔤 Font bar: the font name and size of the selection, shown (and editable) in
// PowerPoint's title bar, right of the deck's name — Office for Mac lets only
// buttons into the Quick Access Toolbar, never the Font / Font Size boxes.
//
// The bar (font-bar/FontBar.swift, a native panel) reads and sets the font by
// itself, over AppleScript, in any deck — the Victor Tools pane plays no part.
// This module only starts it and starts it again if it dies; the bar exits when
// its stdin closes, i.e. with the server.
const { spawn } = require("child_process");
const fs = require("fs");
const path = require("path");

const BIN = path.join(__dirname, "font-bar", ".build", "font-bar");
const LOG = "/tmp/powerpoint-addons-font-bar.log";
const RESPAWN_MS = 5000;

let bar = null;

function startBar() {
  if (!fs.existsSync(BIN)) {
    console.error(`🔤 ${BIN} missing — build it with font-bar/build.sh`);
    return;
  }
  const p = spawn(BIN, [], { stdio: ["pipe", "ignore", fs.openSync(LOG, "a")] });
  bar = p;
  p.on("exit", (code, signal) => {
    if (bar === p) bar = null;
    console.error(new Date().toISOString(), `🔤 font bar exited (${code ?? signal}), restarting`);
    setTimeout(startBar, RESPAWN_MS).unref();
  });
  p.on("error", (e) => console.error("🔤 could not start the font bar:", e.message));
}

function stop() {
  if (bar) { const p = bar; bar = null; p.removeAllListeners("exit"); p.kill(); }
}

module.exports = { startBar, stop };
