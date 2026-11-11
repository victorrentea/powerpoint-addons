// Static HTTPS server for the add-in: Office only loads task panes over https.
// Plus the ✂️ background removal (bg-remove.js) and 🧽 watermark removal (watermark.js)
// the pane can't do in the browser, and the 🔤 font bar's hub (font-bar.js).
const https = require("https");
const fs = require("fs");
const path = require("path");
const bgRemove = require("./bg-remove");
const watermark = require("./watermark");
const fontBar = require("./font-bar");

const PORT = 44344;
const ROOT = path.join(__dirname, "web");
const TYPES = { ".html": "text/html", ".js": "text/javascript", ".css": "text/css",
  ".png": "image/png", ".xml": "application/xml" };

const tls = {
  key: fs.readFileSync(path.join(__dirname, "certs/localhost.key")),
  cert: fs.readFileSync(path.join(__dirname, "certs/localhost.crt")),
};

function json(res, status, body) {
  res.writeHead(status, { "Content-Type": "application/json", "Cache-Control": "no-store" });
  res.end(JSON.stringify(body));
}

function readBody(req) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    req.on("data", (c) => chunks.push(c));
    req.on("end", () => resolve(Buffer.concat(chunks)));
    req.on("error", reject);
  });
}

// POST /bg-remove/prewarm → starts BiRefNet, answers at once.
// POST /bg-remove  body: PNG bytes → { png: base64, box, size, ms }.
async function api(req, res, urlPath) {
  if (urlPath === "/bg-remove/prewarm") {
    bgRemove.prewarm().catch((e) => console.error("✂️", e.message));
    return json(res, 202, { ok: true });
  }
  try {
    const started = Date.now();
    const out = await bgRemove.cutOut(await readBody(req));
    console.log(new Date().toISOString(), `✂️ cut-out in ${Date.now() - started} ms (model ${out.ms} ms)`);
    json(res, 200, { png: out.png.toString("base64"), box: out.box, size: out.size, ms: out.ms });
  } catch (e) {
    console.error("✂️", e.message);
    json(res, 500, { error: e.message });
  }
}

// POST /watermark  body: PNG bytes → { png: base64, found, ms }.
async function watermarkApi(req, res) {
  try {
    const out = await watermark.removeWatermark(await readBody(req));
    console.log(new Date().toISOString(), `🧽 watermark: ${out.found.join(" + ")} box(es) in ${out.ms} ms`);
    json(res, 200, { png: out.png.toString("base64"), found: out.found, ms: out.ms });
  } catch (e) {
    console.error("🧽", e.message);
    json(res, 500, { error: e.message });
  }
}

https.createServer(tls, (req, res) => {
  // 🔤 The pane reports the font on every selection change: too chatty for the log.
  if (!req.url.startsWith("/font/state")) {
    console.log(new Date().toISOString(), req.method, req.url, req.headers["user-agent"] || "");
  }
  const urlPath = decodeURIComponent(new URL(req.url, "https://localhost").pathname);
  if (req.method === "POST" && urlPath.startsWith("/bg-remove")) return api(req, res, urlPath);
  if (req.method === "POST" && urlPath === "/watermark") return watermarkApi(req, res);
  if (req.method === "GET" && urlPath === "/font/events") {
    return fontBar.events(req, res, new URL(req.url, "https://localhost").searchParams.get("pane"));
  }
  if (req.method === "POST" && urlPath === "/font/state") {
    return readBody(req).then((b) => { fontBar.state(b); res.writeHead(204).end(); })
      .catch((e) => json(res, 400, { error: e.message }));
  }
  const file = path.join(ROOT, urlPath === "/" ? "taskpane.html" : urlPath);
  if (!file.startsWith(ROOT)) { res.writeHead(403).end(); return; }
  fs.readFile(file, (err, data) => {
    if (err) { res.writeHead(404).end("not found"); return; }
    res.writeHead(200, { "Content-Type": TYPES[path.extname(file)] || "application/octet-stream",
      "Cache-Control": "no-store" });
    res.end(data);
  });
}).listen(PORT, "127.0.0.1", () => console.log(`powerpoint-addons on https://localhost:${PORT}`));

bgRemove.watchPowerPoint();
fontBar.startBar();

for (const signal of ["SIGTERM", "SIGINT"]) process.on(signal, () => { bgRemove.stop(); fontBar.stop(); process.exit(0); });
