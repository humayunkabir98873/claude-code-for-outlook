// Builds the Outlook COM add-in and its installer (Windows only):
//   npm run build:outlook   ->  dist/ClaudeMailAssistantSetup.exe
// Needs only Windows' .NET Framework 4.8 and an Office install (for the interop assemblies in the GAC).
// The C# compiler and the WebView2 SDK come from nuget.org once (pinned versions, SHA-512 checked against
// nuget.org's catalog) into build/cache. Rebuild after changing anything in public/ or outlook-com/.
import { createHash } from "node:crypto";
import { execFileSync } from "node:child_process";
import fs from "node:fs";
import path from "node:path";

const ROOT = path.resolve(import.meta.dirname, "..");
const SRC = path.join(ROOT, "outlook-com");
const BUILD = path.join(ROOT, "build");
const CACHE = path.join(BUILD, "cache");
const OUT = path.join(BUILD, "outlook"); // the installed layout; the test harness runs from here too
const DIST = path.join(ROOT, "dist");
const WINDIR = process.env.SystemRoot || "C:\\Windows";
const TAR = path.join(WINDIR, "System32", "tar.exe");
const FRAMEWORK = path.join(WINDIR, "Microsoft.NET", "Framework64", "v4.0.30319");
const FRAMEWORK_REFS = ["mscorlib", "System", "System.Core", "System.Drawing", "System.Windows.Forms", "System.Web.Extensions", "Microsoft.CSharp"];

const COMPILER = ["microsoft.net.compilers.toolset", "5.9.0"];
const WEBVIEW2 = ["microsoft.web.webview2", "1.0.4258.31"];

async function nugetPackage([id, version]) {
  const dir = path.join(CACHE, `${id}.${version}`);
  if (fs.existsSync(dir)) return dir;
  console.log(`  downloading ${id} ${version} from nuget.org`);
  const res = await fetch(`https://api.nuget.org/v3-flatcontainer/${id}/${version}/${id}.${version}.nupkg`);
  if (!res.ok) throw new Error(`Download failed (${res.status}) for ${id} ${version}`);
  const bytes = Buffer.from(await res.arrayBuffer());
  const leaf = await (await fetch(`https://api.nuget.org/v3/registration5-gz-semver2/${id}/${version}.json`)).json();
  const catalog = await (await fetch(leaf.catalogEntry)).json();
  if (catalog.packageHash !== createHash("sha512").update(bytes).digest("base64")) throw new Error(`Checksum mismatch for ${id} ${version}`);

  const tmp = `${dir}.tmp`;
  fs.rmSync(tmp, { recursive: true, force: true });
  fs.mkdirSync(tmp, { recursive: true });
  fs.writeFileSync(path.join(tmp, "package.nupkg"), bytes);
  execFileSync(TAR, ["-xf", path.join(tmp, "package.nupkg"), "-C", tmp]);
  fs.rmSync(path.join(tmp, "package.nupkg"));
  fs.renameSync(tmp, dir);
  return dir;
}

function gacAssembly(name) {
  const base = path.join(WINDIR, "assembly", "GAC_MSIL", name);
  const versions = fs.existsSync(base) ? fs.readdirSync(base).sort().reverse() : [];
  for (const v of versions) {
    const file = path.join(base, v, `${name}.dll`);
    if (fs.existsSync(file)) return file;
  }
  throw new Error(`${name}.dll isn't in the Global Assembly Cache. Build on a PC with Microsoft Office installed.`);
}

const listFiles = (dir) =>
  fs.readdirSync(dir, { withFileTypes: true }).flatMap((e) => (e.isDirectory() ? listFiles(path.join(dir, e.name)) : [path.join(dir, e.name)]));
const sources = (...dirs) => dirs.flatMap((d) => fs.readdirSync(d).filter((f) => f.endsWith(".cs")).map((f) => path.join(d, f)));
const quote = (p) => `"${p}"`;

function csc(compiler, name, args) {
  const rsp = path.join(BUILD, `${name}.rsp`);
  const common = ["/nologo", "/nostdlib+", "/langversion:latest", "/optimize+", "/platform:anycpu", "/deterministic+",
    ...FRAMEWORK_REFS.map((r) => `/r:${quote(path.join(FRAMEWORK, `${r}.dll`))}`)];
  fs.writeFileSync(rsp, [...common, ...args].join("\n"));
  execFileSync(path.join(compiler, "tasks", "net472", "csc.exe"), ["/noconfig", `@${rsp}`], { stdio: "inherit" });
}

// An .ico made of PNG images (supported since Windows Vista).
function writeIcon(file, pngs) {
  const header = Buffer.alloc(6 + 16 * pngs.length);
  header.writeUInt16LE(1, 2);
  header.writeUInt16LE(pngs.length, 4);
  let offset = header.length;
  pngs.forEach(({ size, data }, i) => {
    const e = 6 + 16 * i;
    header.writeUInt8(size >= 256 ? 0 : size, e);
    header.writeUInt8(size >= 256 ? 0 : size, e + 1);
    header.writeUInt16LE(1, e + 4);
    header.writeUInt16LE(32, e + 6);
    header.writeUInt32LE(data.length, e + 8);
    header.writeUInt32LE(offset, e + 12);
    offset += data.length;
  });
  fs.writeFileSync(file, Buffer.concat([header, ...pngs.map((p) => p.data)]));
}

if (process.platform !== "win32") throw new Error("The Outlook add-in can only be built on Windows.");
fs.mkdirSync(CACHE, { recursive: true });
const compiler = await nugetPackage(COMPILER);
const webview2 = await nugetPackage(WEBVIEW2);
const wvLib = path.join(webview2, "lib", "net462");
const wvRefs = ["Microsoft.Web.WebView2.Core.dll", "Microsoft.Web.WebView2.WinForms.dll"].map((f) => path.join(wvLib, f));

// 1. The add-in, laid out exactly as it will be installed.
fs.rmSync(OUT, { recursive: true, force: true });
fs.mkdirSync(OUT, { recursive: true });
const dll = path.join(OUT, "ClaudeMailAssistant.dll");
console.log("Compiling ClaudeMailAssistant.dll");
csc(compiler, "addin", [
  "/target:library",
  `/out:${quote(dll)}`,
  `/link:${quote(gacAssembly("office"))}`, // embed only the Office interop types we use
  `/link:${quote(gacAssembly("Microsoft.Office.Interop.Outlook"))}`,
  ...wvRefs.map((r) => `/r:${quote(r)}`),
  `/resource:${path.join(SRC, "AddIn", "Ribbon.xml")},Ribbon.xml`,
  `/resource:${path.join(SRC, "AddIn", "office-shim.js")},office-shim.js`,
  `/resource:${path.join(ROOT, "public", "assets", "icon-80.png")},icon-80.png`,
  ...sources(SRC, path.join(SRC, "AddIn")).map(quote),
]);
for (const r of wvRefs) fs.copyFileSync(r, path.join(OUT, path.basename(r)));
for (const arch of ["x64", "x86"]) {
  fs.mkdirSync(path.join(OUT, arch));
  fs.copyFileSync(path.join(webview2, "runtimes", `win-${arch}`, "native", "WebView2Loader.dll"), path.join(OUT, arch, "WebView2Loader.dll"));
}
fs.cpSync(path.join(ROOT, "public"), path.join(OUT, "public"), { recursive: true });
const payload = listFiles(OUT);

// 2. Test harness (not shipped).
console.log("Compiling Harness.exe (test only)");
csc(compiler, "harness", ["/target:exe", `/out:${quote(path.join(OUT, "Harness.exe"))}`, `/r:${quote(dll)}`, ...wvRefs.map((r) => `/r:${quote(r)}`),
  ...sources(path.join(SRC, "Harness")).map(quote)]);

// 3. The installer, with the add-in embedded.
const icon = path.join(BUILD, "setup.ico");
writeIcon(icon, [16, 32, 64, 128].map((size) => ({ size, data: fs.readFileSync(path.join(ROOT, "public", "assets", `icon-${size}.png`)) })));
fs.mkdirSync(DIST, { recursive: true });
const setup = path.join(DIST, "ClaudeMailAssistantSetup.exe");
console.log("Compiling ClaudeMailAssistantSetup.exe");
csc(compiler, "setup", [
  "/target:winexe",
  `/out:${quote(setup)}`,
  `/win32icon:${quote(icon)}`,
  `/win32manifest:${quote(path.join(SRC, "Setup", "app.manifest"))}`,
  ...payload.map((f) => `/resource:${f},payload/${path.relative(OUT, f).split(path.sep).join("/")}`),
  quote(path.join(SRC, "Ids.cs")),
  ...sources(path.join(SRC, "Setup")).map(quote),
]);
console.log(`${path.relative(ROOT, setup)} (${(fs.statSync(setup).size / 1e6).toFixed(1)} MB, ${payload.length} files embedded)`);
