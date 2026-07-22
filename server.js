/* Minimal HTTPS static file server for the add-in (uses Office dev certs). */
const https = require("https");
const fs = require("fs");
const path = require("path");
const devCerts = require("office-addin-dev-certs");

const PORT = 3002;
const ROOT = __dirname;

const MIME = {
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".png": "image/png",
  ".xml": "text/xml; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".ico": "image/x-icon"
};

async function main() {
  const options = await devCerts.getHttpsServerOptions();

  const server = https.createServer(options, (req, res) => {
    const urlPath = decodeURIComponent(req.url.split("?")[0]);
    let filePath = path.normalize(
      path.join(ROOT, urlPath === "/" ? "/src/taskpane/taskpane.html" : urlPath)
    );

    if (!filePath.startsWith(ROOT)) {
      res.writeHead(403);
      return res.end("Forbidden");
    }

    fs.readFile(filePath, (err, data) => {
      if (err) {
        res.writeHead(404);
        return res.end("Not found: " + urlPath);
      }
      res.writeHead(200, {
        "Content-Type": MIME[path.extname(filePath).toLowerCase()] || "application/octet-stream",
        "Cache-Control": "no-cache"
      });
      res.end(data);
    });
  });

  server.listen(PORT, () => {
    console.log(`Excel Map add-in server running at https://localhost:${PORT}/`);
  });
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
