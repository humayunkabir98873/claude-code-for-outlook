// Serves the add-in's files over local HTTPS (Outlook requires HTTPS). Runs as `node server.js`, or as
// the packaged executable with the files embedded (see scripts/build.mjs). No API keys or mail pass through here.
const https = require("https");
const fs = require("fs");
const path = require("path");
const { loadCertificates } = require("./certs");

const PORT = Number(process.env.PORT) || 3000;
const PUBLIC = path.join(__dirname, "public");
const sea = (() => {
  try {
    const s = require("node:sea");
    return s.isSea() ? s : null;
  } catch {
    return null;
  }
})();

const TYPES = {
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".json": "application/json",
  ".png": "image/png",
  ".svg": "image/svg+xml",
  ".ico": "image/x-icon",
};

function readFile(rel) {
  try {
    return sea ? Buffer.from(sea.getAsset(`public/${rel}`)) : fs.readFileSync(path.join(PUBLIC, rel));
  } catch {
    return null;
  }
}

function handle(req, res) {
  let rel = "";
  try {
    rel = decodeURIComponent(new URL(req.url, "https://localhost").pathname).replace(/^\/+/, "");
  } catch {
    /* malformed URL: 404 below */
  }
  const allowed = (req.method === "GET" || req.method === "HEAD") && rel && !rel.split(/[\\/]/).includes("..");
  const body = allowed ? readFile(rel) : null;
  if (!body) {
    res.writeHead(404, { "Content-Type": "text/plain" });
    return res.end("Not found");
  }
  res.writeHead(200, {
    "Content-Type": TYPES[path.extname(rel).toLowerCase()] || "application/octet-stream",
    "Content-Length": body.length,
    "Cache-Control": "no-store",
    "X-Content-Type-Options": "nosniff",
  });
  res.end(req.method === "HEAD" ? undefined : body);
}

function fail(message) {
  console.error(`Couldn't start Claude Mail Assistant: ${message}`);
  // A double-clicked executable closes its window on exit, so keep the message readable.
  if (sea && process.stdin.isTTY) {
    console.error("Press Enter to close.");
    process.stdin.once("data", () => process.exit(1));
  } else {
    process.exit(1);
  }
}

(async () => {
  const { key, cert, ca } = await loadCertificates({ skipTrust: process.argv.includes("--skip-trust") });
  const server = https.createServer({ key, cert, ca }, handle);
  server.on("error", (err) =>
    fail(err.code === "EADDRINUSE" ? `port ${PORT} is already in use. Is the add-in already running in another window?` : err.message)
  );
  server.listen(PORT, "localhost", () => {
    console.log(`Claude Mail Assistant is running at https://localhost:${PORT}`);
    console.log("Leave this window open while you use the add-in in Outlook.");
  });
})().catch((err) => fail(err.message));
