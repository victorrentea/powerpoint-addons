// Static HTTPS server for the add-in: Office only loads task panes over https.
const https = require("https");
const fs = require("fs");
const path = require("path");

const PORT = 44344;
const ROOT = path.join(__dirname, "web");
const TYPES = { ".html": "text/html", ".js": "text/javascript", ".css": "text/css",
  ".png": "image/png", ".xml": "application/xml" };

const tls = {
  key: fs.readFileSync(path.join(__dirname, "certs/localhost.key")),
  cert: fs.readFileSync(path.join(__dirname, "certs/localhost.crt")),
};

https.createServer(tls, (req, res) => {
  console.log(new Date().toISOString(), req.method, req.url, req.headers["user-agent"] || "");
  const urlPath = decodeURIComponent(new URL(req.url, "https://localhost").pathname);
  const file = path.join(ROOT, urlPath === "/" ? "taskpane.html" : urlPath);
  if (!file.startsWith(ROOT)) { res.writeHead(403).end(); return; }
  fs.readFile(file, (err, data) => {
    if (err) { res.writeHead(404).end("not found"); return; }
    res.writeHead(200, { "Content-Type": TYPES[path.extname(file)] || "application/octet-stream",
      "Cache-Control": "no-store" });
    res.end(data);
  });
}).listen(PORT, "127.0.0.1", () => console.log(`victor-ppt-addin on https://localhost:${PORT}`));
