// Pure helpers with no Office or DOM dependencies (except analyzeLinks, which takes a parser).

export function extractJson(text) {
  const start = text.indexOf("{");
  const end = text.lastIndexOf("}");
  if (start === -1 || end <= start) throw new Error("Claude didn't return a usable answer. Try again.");
  try {
    return JSON.parse(text.slice(start, end + 1));
  } catch {
    throw new Error("Claude's answer couldn't be read. Try again.");
  }
}

export function localDate(d = new Date()) {
  const p = (n) => String(n).padStart(2, "0");
  return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}`;
}

export function parseHeaders(raw) {
  const out = {};
  const unfolded = String(raw || "").replace(/\r?\n[ \t]+/g, " ");
  for (const line of unfolded.split(/\r?\n/)) {
    const i = line.indexOf(":");
    if (i <= 0) continue;
    const name = line.slice(0, i).trim().toLowerCase();
    (out[name] ||= []).push(line.slice(i + 1).trim());
  }
  return out;
}

export function domainOf(addr) {
  const m = String(addr || "").toLowerCase().match(/@([a-z0-9.-]+\.[a-z]{2,})/);
  return m ? m[1] : "";
}

// A domain written inside text such as a display name ("PayPal paypal.com"), limited to common TLDs to avoid false hits like "john.smith".
export function domainIn(text) {
  const m = String(text || "").toLowerCase().match(/((?:[a-z0-9-]+\.)+(?:com|net|org|edu|gov|io|co|uk|de|fr|in|info|biz|ru|cn|us|ca|au|app|dev|ai|me|xyz|top|online|site|shop|store))\b/);
  return m ? m[1] : "";
}

const SECOND_LEVEL = new Set(["co", "com", "org", "net", "gov", "ac", "edu"]);
export function baseDomain(host) {
  const p = String(host || "").toLowerCase().split(".").filter(Boolean);
  if (p.length <= 2) return p.join(".");
  const [a, b] = p.slice(-2);
  if (b.length === 2 && SECOND_LEVEL.has(a)) return p.slice(-3).join(".");
  return p.slice(-2).join(".");
}

const SHORTENERS = new Set(["bit.ly", "tinyurl.com", "t.co", "goo.gl", "ow.ly", "is.gd", "buff.ly", "rebrand.ly", "cutt.ly", "shorturl.at"]);
const RISKY_EXT = new Set(["exe", "js", "jse", "vbs", "vbe", "scr", "bat", "cmd", "com", "pif", "lnk", "iso", "img", "msi", "jar", "ps1", "hta", "html", "htm", "docm", "xlsm", "pptm", "one", "svg"]);

export function analyzeAttachments(names) {
  const signals = [];
  for (const n of names || []) {
    const ext = String(n).toLowerCase().split(".").pop();
    const parts = String(n).split(".");
    if (RISKY_EXT.has(ext)) signals.push({ level: "bad", text: `Attachment "${n}" is a file type often used to spread malware` });
    else if (parts.length > 2 && RISKY_EXT.has(parts[parts.length - 1].toLowerCase()) === false && /^(pdf|docx?|xlsx?|jpg|png)$/i.test(parts[parts.length - 2]) && ext === "zip")
      signals.push({ level: "warn", text: `Attachment "${n}" has a double extension` });
  }
  return signals;
}

export function analyzeLinks(html, Parser = globalThis.DOMParser) {
  const signals = [];
  const hosts = new Set();
  const seen = new Set();
  const add = (level, text) => {
    if (!seen.has(text)) {
      seen.add(text);
      signals.push({ level, text });
    }
  };
  const doc = new Parser().parseFromString(html || "", "text/html");
  for (const a of doc.querySelectorAll("a[href]")) {
    let url;
    try {
      url = new URL(a.getAttribute("href") || "");
    } catch {
      continue;
    }
    if (!/^https?:$/.test(url.protocol)) continue;
    if (url.hostname.endsWith("safelinks.protection.outlook.com")) {
      const inner = url.searchParams.get("url");
      if (inner) {
        try {
          url = new URL(inner);
        } catch {
          /* keep the wrapper */
        }
      }
    }
    const host = url.hostname.toLowerCase();
    hosts.add(host);
    const text = (a.textContent || "").trim();
    const shown = text.match(/^(?:https?:\/\/)?((?:[a-z0-9-]+\.)+[a-z]{2,})(?:[/:?#]|$)/i);
    if (shown && baseDomain(shown[1]) !== baseDomain(host)) add("bad", `A link shows "${shown[1].toLowerCase()}" but goes to ${host}`);
    if (/^\d{1,3}(\.\d{1,3}){3}$/.test(host)) add("warn", `A link points to a raw IP address (${host})`);
    if (host.startsWith("xn--") || host.includes(".xn--")) add("warn", `A link uses an internationalized lookalike-style domain (${host})`);
    if (SHORTENERS.has(host)) add("warn", `A link uses a URL shortener (${host})`);
  }
  if (hosts.size === 0) add("ok", "No web links in the message");
  else if (!signals.some((s) => s.level !== "ok")) add("ok", "Link text matches where the links go");
  return { signals, hosts: [...hosts].slice(0, 15) };
}

export function analyzeHeaders(headers, fromName, fromAddr) {
  const signals = [];
  const fromDomain = domainOf(fromAddr);
  if (!headers) {
    signals.push({ level: "warn", text: "This version of Outlook didn't give access to message headers, so SPF, DKIM and DMARC weren't checked" });
  } else {
    const auth = (headers["authentication-results"] || []).join(" ").toLowerCase();
    for (const k of ["spf", "dkim", "dmarc"]) {
      const results = [...auth.matchAll(new RegExp(`\\b${k}=(\\w+)`, "g"))].map((m) => m[1]);
      const label = k.toUpperCase();
      if (!results.length) signals.push({ level: "warn", text: `${label} result not reported in the headers` });
      else if (results.includes("fail")) signals.push({ level: "bad", text: `${label} failed` });
      else if (results.every((r) => r === "pass")) signals.push({ level: "ok", text: `${label} passed` });
      else signals.push({ level: "warn", text: `${label} result: ${[...new Set(results)].join(", ")}` });
    }
    const replyTo = (headers["reply-to"] || [])[0];
    const rtDomain = domainOf(replyTo);
    if (rtDomain && fromDomain && baseDomain(rtDomain) !== baseDomain(fromDomain))
      signals.push({ level: "warn", text: `Replies go to ${rtDomain}, not the sender's domain (${fromDomain})` });
  }
  const nameDomain = domainIn(fromName);
  if (nameDomain && fromDomain && baseDomain(nameDomain) !== baseDomain(fromDomain))
    signals.push({ level: "bad", text: `The sender's display name mentions ${nameDomain} but the address is at ${fromDomain}` });
  return signals;
}

// ---------- drafts ----------

// Where the email being replied to or forwarded starts, in the HTML Outlook (Word), Outlook on the web and
// older Outlook versions produce. The earliest match wins.
const QUOTE_MARKERS = [
  /<div[^>]*\bid=["']?(?:divRplyFwdMsg|mail-editor-reference-message-container|appendonsend)\b/i,
  /<div[^>]*border-top:\s*solid\s+#(?:E1E1E1|B5C4DF)/i,
  /<hr\b/i,
  /-----\s*Original Message\s*-----/i,
];
// Outlook marks the start of an inserted signature with this bookmark (Outlook on the web uses a div).
const SIGNATURE_MARKERS = [/<a\s+name=["']?_MailAutoSig/i, /<div[^>]*\bid=["']?Signature\b/i];

// A marker found inside a paragraph (a bookmark, a text line) moves back to where that <p> or <div> opens.
function blockStart(html, index) {
  if (/^<(?:div|hr)\b/i.test(html.slice(index, index + 4))) return index;
  let start = index;
  for (const m of html.slice(0, index).matchAll(/<(?:p|div)\b/gi)) start = m.index;
  return start;
}

function firstMatch(html, patterns, from) {
  let best = -1;
  for (const re of patterns) {
    const i = html.slice(from).search(re);
    if (i !== -1 && (best === -1 || from + i < best)) best = from + i;
  }
  return best;
}

// Splits a draft's HTML into the part the user wrote ("mine") and everything to keep as is: the document
// head before it, then the signature and the quoted email after it ("rest"; "quote" is just the quoted email).
export function splitDraft(html) {
  html = String(html || "");
  const open = html.match(/<div\s+class=["']?WordSection1["']?[^>]*>/i) || html.match(/<body\b[^>]*>/i);
  const start = open ? open.index + open[0].length : 0;
  const close = html.search(/(?:<\/div>\s*)?<\/body>/i);
  const end = close >= start ? close : html.length;

  const quoteAt = firstMatch(html, QUOTE_MARKERS, start);
  const sigAt = firstMatch(html, SIGNATURE_MARKERS, start);
  const quote = quoteAt === -1 || quoteAt > end ? end : blockStart(html, quoteAt);
  const sig = sigAt === -1 || sigAt > quote ? quote : blockStart(html, sigAt);
  return { before: html.slice(0, start), mine: html.slice(start, sig), rest: html.slice(sig), quote: html.slice(quote, end) };
}

// Plain text as Outlook-style paragraphs, so it looks like typed text in the draft.
export function textToHtml(text) {
  const esc = (s) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  return String(text || "")
    .replace(/\r\n?/g, "\n")
    .split("\n")
    .map((line) => `<p class=MsoNormal>${line.trim() ? esc(line) : "&nbsp;"}</p>`)
    .join("\n");
}

// Visible text of an HTML fragment, one line per paragraph. HTML source line breaks (Word wraps its
// source at 80 columns) are only spaces, so collapse whitespace before turning block ends into newlines.
export function htmlToText(html, Parser = globalThis.DOMParser) {
  const marked = String(html || "")
    .replace(/<(style|script|head)\b[\s\S]*?<\/\1>/gi, "")
    .replace(/\s+/g, " ")
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<\/(?:p|div|li|h[1-6]|tr)>/gi, "$&\n");
  const text = new Parser().parseFromString(marked, "text/html").body.textContent || "";
  return text
    .replace(/ /g, " ")
    .split("\n")
    .map((l) => l.trim())
    .join("\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

const isDate = (s) => /^\d{4}-\d{2}-\d{2}$/.test(s || "");

// filters: false leaves out the attachment and date conditions (used when Graph rejects them).
export function buildKql(plan, { filters = true } = {}) {
  const clean = (s) => String(s || "").replace(/["\\()]/g, " ").trim();
  const words = (s) => clean(s).split(/\s+/).filter(Boolean);
  const parts = [];
  const text = words(plan.text).join(" ");
  if (text) parts.push(text);
  for (const w of words(plan.from)) parts.push(`from:${w}`);
  for (const w of words(plan.subject)) parts.push(`subject:${w}`);
  if (filters) {
    if (plan.hasAttachment) parts.push("hasAttachments:true");
    if (isDate(plan.after)) parts.push(`received>=${plan.after}`);
    if (isDate(plan.before)) parts.push(`received<=${plan.before}`);
  }
  return parts.join(" AND ");
}

// The plan's dates as local-time bounds: from = start of "after", to = start of the day after "before" (exclusive).
export function planDateRange(plan) {
  const day = (s, add = 0) => {
    const [y, m, d] = s.split("-").map(Number);
    return new Date(y, m - 1, d + add);
  };
  return {
    from: isDate(plan.after) ? day(plan.after) : null,
    to: isDate(plan.before) ? day(plan.before, 1) : null,
  };
}
