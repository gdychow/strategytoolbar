const http = require("http");
const https = require("https");
const fs = require("fs");
const path = require("path");
const os = require("os");

const PORT = process.env.PORT || 3000;
const ROOT = path.join(__dirname, "dist");

// Cloudflare Tunnel (or any reverse proxy) terminates TLS at the edge — the
// container only ever needs to speak plain HTTP internally. Local dev still
// needs real HTTPS, since Office requires it and there's no tunnel in front
// of localhost. Set USE_TLS=false (the Dockerfile does this) to skip cert
// loading entirely.
const useTls = process.env.USE_TLS !== "false";

const MIME = {
  ".html": "text/html",
  ".js": "text/javascript",
  ".css": "text/css",
  ".xml": "application/xml",
  ".png": "image/png",
  ".map": "application/json",
};

function requestHandler(req, res) {
  const urlPath = decodeURIComponent(req.url.split("?")[0]);
  const filePath = path.join(ROOT, urlPath === "/" ? "/taskpane.html" : urlPath);

  if (!filePath.startsWith(ROOT)) {
    res.writeHead(403);
    res.end("Forbidden");
    return;
  }

  fs.readFile(filePath, (err, data) => {
    if (err) {
      res.writeHead(404);
      res.end("Not found: " + urlPath);
      return;
    }
    const ext = path.extname(filePath);
    res.writeHead(200, {
      "Content-Type": MIME[ext] || "application/octet-stream",
      "Access-Control-Allow-Origin": "*",
    });
    res.end(data);
  });
}

if (useTls) {
  const certDir = path.join(os.homedir(), ".office-addin-dev-certs");
  const options = {
    key: fs.readFileSync(path.join(certDir, "localhost.key")),
    cert: fs.readFileSync(path.join(certDir, "localhost.crt")),
  };
  https.createServer(options, requestHandler).listen(PORT, () => {
    console.log(`Strategy Toolbar dev server running at https://localhost:${PORT}/taskpane.html`);
  });
} else {
  http.createServer(requestHandler).listen(PORT, () => {
    console.log(`Strategy Toolbar server running on plain HTTP, port ${PORT} (TLS expected to terminate upstream)`);
  });
}
