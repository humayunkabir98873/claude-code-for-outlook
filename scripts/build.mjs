// Builds the local server as a single executable (Node.js SEA) for Windows and macOS, with public/ embedded.
//   npm run build              all targets
//   npm run build -- macos     only targets whose name contains "macos"
// Output: dist/<target>.zip containing the executable, manifest.xml and HOW-TO-USE.txt.
// macOS runtimes are downloaded from nodejs.org (same version as the Node running this script) and
// checked against the published SHA-256 sums. Rebuild after changing anything in public/.
import { build } from "esbuild";
import { inject } from "postject";
import { createHash } from "node:crypto";
import { execFileSync } from "node:child_process";
import fs from "node:fs";
import path from "node:path";

const ROOT = path.resolve(import.meta.dirname, "..");
const BUILD = path.join(ROOT, "build");
const CACHE = path.join(BUILD, "cache");
const DIST = path.join(ROOT, "dist");
const NAME = "ClaudeMailAssistant";
const VERSION = process.version; // the embedded blob must run on the same Node version that made it
const FUSE = "NODE_SEA_FUSE_fce680ab2cc467b6e072b8b5df1996b2";
// Windows' own bsdtar reads and writes zip; Git Bash's GNU tar (often first on PATH) does neither.
const TAR = process.platform === "win32" ? path.join(process.env.SystemRoot, "System32", "tar.exe") : "tar";

const TARGETS = [
  { name: "windows-x64", node: "win-x64", exe: `${NAME}.exe` },
  { name: "macos-arm64", node: "darwin-arm64", exe: NAME, mac: true },
  { name: "macos-x64", node: "darwin-x64", exe: NAME, mac: true },
];

const HOWTO_WINDOWS = `Claude Mail Assistant for classic Outlook on Windows

1. Double-click ${NAME}.exe. If Windows SmartScreen appears, choose "More info", then "Run anyway"
   (the program isn't code-signed).
2. The first time, Windows asks whether to install a certificate from "Developer CA for Microsoft Office
   Add-ins". Choose Yes. Outlook only loads add-ins over HTTPS, and this certificate is used only for
   https://localhost on this PC.
3. Leave the window open while you use the add-in. Close it to stop.
4. In Outlook, open an email, then Get Add-ins > My add-ins > Custom add-ins > Add from file, and pick
   manifest.xml from this folder. You only do this once.
5. Click Claude on the ribbon, open Settings, paste your Claude API key, and choose Save.
`;

const HOWTO_MAC = `Claude Mail Assistant for Outlook on Mac

1. One time: open Terminal, type "cd " (with a space), drag this folder onto the Terminal window, press
   Return, then paste this line and press Return:

   chmod +x ${NAME} && xattr -c ${NAME} && codesign --force --sign - ${NAME}

   (macOS only runs downloaded programs that are signed. This signs it for your Mac only.)
2. Start it by double-clicking ${NAME} (it opens in Terminal), or run ./${NAME} in Terminal.
3. The first time, macOS asks for your password to trust the localhost certificate
   ("Developer CA for Microsoft Office Add-ins"). Outlook only loads add-ins over HTTPS.
4. Leave the Terminal window open while you use the add-in. Press Control-C or close it to stop.
5. In Outlook, open an email, then Get Add-ins > My add-ins > Add a custom add-in > Add from file, and
   pick manifest.xml from this folder. You only do this once.
6. Click Claude, open Settings, paste your Claude API key, and choose Save.
`;

async function download(url, file) {
  const res = await fetch(url);
  if (!res.ok) throw new Error(`Download failed (${res.status}): ${url}`);
  fs.writeFileSync(file, Buffer.from(await res.arrayBuffer()));
}

// Returns the path of a Node runtime binary for the target.
async function nodeBinary(target) {
  if (target.node === `${process.platform === "win32" ? "win" : process.platform}-${process.arch}`) return process.execPath;
  if (!target.mac) throw new Error(`Build ${target.name} on that platform, or add a download step for it.`);

  const dir = `node-${VERSION}-${target.node}`;
  const binary = path.join(CACHE, dir, "bin", "node");
  if (fs.existsSync(binary)) return binary;

  const base = `https://nodejs.org/dist/${VERSION}`;
  const archive = path.join(CACHE, `${dir}.tar.gz`);
  console.log(`  downloading ${base}/${dir}.tar.gz`);
  await download(`${base}/${dir}.tar.gz`, archive);
  const sums = await (await fetch(`${base}/SHASUMS256.txt`)).text();
  const expected = sums.split("\n").find((l) => l.endsWith(`  ${dir}.tar.gz`))?.split(" ")[0];
  const actual = createHash("sha256").update(fs.readFileSync(archive)).digest("hex");
  if (!expected || expected !== actual) {
    fs.rmSync(archive);
    throw new Error(`Checksum mismatch for ${dir}.tar.gz`);
  }
  execFileSync(TAR, ["-xzf", archive, "-C", CACHE, `${dir}/bin/node`]);
  fs.rmSync(archive);
  return binary;
}

function listFiles(dir) {
  return fs.readdirSync(dir, { withFileTypes: true }).flatMap((e) => {
    const full = path.join(dir, e.name);
    return e.isDirectory() ? listFiles(full) : [full];
  });
}

async function prepareBlob() {
  fs.mkdirSync(CACHE, { recursive: true });
  await build({
    entryPoints: [path.join(ROOT, "server.js")],
    bundle: true,
    platform: "node",
    format: "cjs",
    target: `node${process.versions.node.split(".")[0]}`,
    outfile: path.join(BUILD, "server.cjs"),
    logLevel: "warning",
  });
  const assets = Object.fromEntries(
    listFiles(path.join(ROOT, "public")).map((f) => [path.relative(ROOT, f).split(path.sep).join("/"), f])
  );
  const config = {
    main: path.join(BUILD, "server.cjs"),
    output: path.join(BUILD, "sea-prep.blob"),
    disableExperimentalSEAWarning: true,
    useSnapshot: false, // both must stay off for the blob to run on other platforms
    useCodeCache: false,
    assets,
  };
  fs.writeFileSync(path.join(BUILD, "sea-config.json"), JSON.stringify(config, null, 2));
  execFileSync(process.execPath, ["--experimental-sea-config", path.join(BUILD, "sea-config.json")], { stdio: "inherit" });
  console.log(`Embedded ${Object.keys(assets).length} files from public/`);
  return fs.readFileSync(config.output);
}

async function buildTarget(target, blob) {
  console.log(`${target.name}:`);
  const out = path.join(DIST, `${NAME}-${target.name}`);
  fs.rmSync(out, { recursive: true, force: true });
  fs.mkdirSync(out, { recursive: true });

  const exe = path.join(out, target.exe);
  fs.copyFileSync(await nodeBinary(target), exe);
  fs.chmodSync(exe, 0o755);
  await inject(exe, "NODE_SEA_BLOB", blob, { sentinelFuse: FUSE, ...(target.mac && { machoSegmentName: "NODE_SEA" }) });
  fs.copyFileSync(path.join(ROOT, "manifest.xml"), path.join(out, "manifest.xml"));
  fs.writeFileSync(path.join(out, "HOW-TO-USE.txt"), target.mac ? HOWTO_MAC : HOWTO_WINDOWS);

  const zip = `${out}.zip`;
  fs.rmSync(zip, { force: true });
  execFileSync(TAR, ["-a", "-cf", zip, "-C", DIST, path.basename(out)]);
  console.log(`  ${path.relative(ROOT, zip)} (${(fs.statSync(zip).size / 1e6).toFixed(0)} MB)`);
}

const filter = process.argv[2];
const targets = TARGETS.filter((t) => !filter || t.name.includes(filter));
if (!targets.length) throw new Error(`No target matches "${filter}". Targets: ${TARGETS.map((t) => t.name).join(", ")}`);
const blob = await prepareBlob();
for (const t of targets) await buildTarget(t, blob);
