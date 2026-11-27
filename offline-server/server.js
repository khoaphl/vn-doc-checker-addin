const https = require("https");
const fs = require("fs");
const path = require("path");

const PORT = Number(process.env.PORT) || 8443;
const ROOT_DIR = path.resolve(__dirname, "..");
const CERTS_DIR = path.join(__dirname, "certs");

const MIME_TYPES = {
  ".html": "text/html; charset=utf-8",
  ".js": "application/javascript; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".png": "image/png",
  ".jpg": "image/jpeg",
  ".svg": "image/svg+xml",
  ".txt": "text/plain; charset=utf-8",
  ".xml": "application/xml; charset=utf-8"
};

function loadTlsOptions() {
  const pfxPath = path.join(CERTS_DIR, "localhost.pfx");
  const keyPath = path.join(CERTS_DIR, "localhost.key");
  const crtPath = path.join(CERTS_DIR, "localhost.crt");

  if (fs.existsSync(pfxPath)) {
    return {
      pfx: fs.readFileSync(pfxPath),
      passphrase: process.env.CERT_PASSWORD || "pcdt-offline"
    };
  }

  if (fs.existsSync(keyPath) && fs.existsSync(crtPath)) {
    return {
      key: fs.readFileSync(keyPath),
      cert: fs.readFileSync(crtPath)
    };
  }

  console.error("[LỖI] Chưa tìm thấy chứng chỉ SSL trong thư mục:", CERTS_DIR);
  console.error("Hãy chạy PowerShell: .\\offline-server\\certs\\create-cert.ps1");
  process.exit(1);
}

const server = https.createServer(loadTlsOptions(), (req, res) => {
  const rawUrl = (req.url || "/").split("?")[0];
  const relPath = rawUrl === "/" ? "taskpane.html" : rawUrl.replace(/^\/+/, "");
  const safePath = path.normalize(relPath).replace(/^(\.\.(\/|\\|$))+/, "");
  const fullPath = path.join(ROOT_DIR, safePath);

  if (!fullPath.startsWith(ROOT_DIR)) {
    res.writeHead(403);
    res.end("Forbidden");
    return;
  }

  fs.stat(fullPath, (err, stats) => {
    if (err || !stats.isFile()) {
      res.writeHead(404, { "Content-Type": "text/plain; charset=utf-8" });
      res.end("Not Found: " + safePath);
      return;
    }

    const ext = path.extname(fullPath).toLowerCase();
    const contentType = MIME_TYPES[ext] || "application/octet-stream";

    res.writeHead(200, {
      "Content-Type": contentType,
      "Cache-Control": "no-store, no-cache, must-revalidate, proxy-revalidate",
      "Pragma": "no-cache",
      "Expires": "0"
    });
    fs.createReadStream(fullPath).pipe(res);
  });
});

server.listen(PORT, "127.0.0.1", () => {
  console.log(`[PCĐT Offline Server] Đang chạy tại https://localhost:${PORT}/taskpane.html`);
});
