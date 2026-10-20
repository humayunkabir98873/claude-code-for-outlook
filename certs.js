// Creates and trusts the localhost HTTPS certificate Outlook needs. Uses the same folder, file names
// and CA name as office-addin-dev-certs, so certificates installed by either tool are reused, but it
// trusts the CA with the OS's own tools so it also works inside the packaged executable.
const { X509Certificate } = require("crypto");
const { execFileSync } = require("child_process");
const fs = require("fs");
const os = require("os");
const path = require("path");
const mkcert = require("mkcert");

const CA_NAME = "Developer CA for Microsoft Office Add-ins";
const DIR = process.env.CMA_CERT_DIR || path.join(os.homedir(), ".office-addin-dev-certs");
const FILES = { ca: "ca.crt", cert: "localhost.crt", key: "localhost.key" };
const VALID_DAYS = 365;
const DAY = 24 * 60 * 60 * 1000;

const filePath = (name) => path.join(DIR, FILES[name]);
const thumbprint = (cert) => cert.fingerprint.replace(/:/g, "").toUpperCase();

function readExisting() {
  try {
    const pem = Object.fromEntries(Object.keys(FILES).map((k) => [k, fs.readFileSync(filePath(k), "utf8")]));
    const ca = new X509Certificate(pem.ca);
    const cert = new X509Certificate(pem.cert);
    const fresh = (c) => new Date(c.validTo).getTime() > Date.now() + DAY;
    return fresh(ca) && fresh(cert) && cert.checkIssued(ca) ? { pem, ca } : null;
  } catch {
    return null;
  }
}

async function generate() {
  const ca = await mkcert.createCA({ organization: CA_NAME, countryCode: "US", state: "WA", locality: "Redmond", validity: VALID_DAYS });
  const leaf = await mkcert.createCert({ ca, domains: ["127.0.0.1", "localhost"], validity: VALID_DAYS });
  fs.mkdirSync(DIR, { recursive: true });
  fs.writeFileSync(filePath("ca"), ca.cert);
  fs.writeFileSync(filePath("cert"), leaf.cert);
  fs.writeFileSync(filePath("key"), leaf.key, { mode: 0o600 });
  return readExisting();
}

function isTrusted(ca) {
  try {
    if (process.platform === "win32") {
      execFileSync("certutil", ["-user", "-store", "Root", thumbprint(ca)], { stdio: "ignore" });
      return true;
    }
    if (process.platform === "darwin") {
      const out = execFileSync("security", ["find-certificate", "-a", "-Z", "-c", CA_NAME], { encoding: "utf8" });
      return out.toUpperCase().includes(thumbprint(ca));
    }
  } catch {
    /* not found */
  }
  return false;
}

function trust() {
  console.log("Adding the localhost certificate to your trusted certificates (one time).");
  if (process.platform === "win32") {
    console.log('Windows will ask you to confirm: choose "Yes".');
    execFileSync("certutil", ["-user", "-addstore", "Root", filePath("ca")], { stdio: "ignore" });
  } else if (process.platform === "darwin") {
    console.log("macOS will ask for your password.");
    const keychain = path.join(os.homedir(), "Library/Keychains/login.keychain-db");
    execFileSync("security", ["add-trusted-cert", "-d", "-r", "trustRoot", "-k", keychain, filePath("ca")], { stdio: "inherit" });
  } else {
    throw new Error(`Trust ${filePath("ca")} in your system certificate store, then start again with --skip-trust.`);
  }
}

// skipTrust: use the certificate without touching the OS trust store (for testing).
async function loadCertificates({ skipTrust = false } = {}) {
  const existing = readExisting() || (await generate());
  if (!skipTrust && !isTrusted(existing.ca)) {
    try {
      trust();
    } catch {
      throw new Error("The certificate wasn't trusted, so Outlook won't load the add-in. Start again and accept the prompt.");
    }
  }
  return { ca: existing.pem.ca, cert: existing.pem.cert, key: existing.pem.key };
}

module.exports = { loadCertificates, certificateDir: DIR };
