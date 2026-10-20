/* global Office */
import { loadSettings, saveSettings, MODELS } from "./settings.js";
import * as api from "./claude.js";
import * as mail from "./mail.js";
import * as graph from "./graph.js";
import { buildKql } from "./util.js";

const $ = (id) => document.getElementById(id);
const TABS = ["email", "search", "settings"];
const WORKING = "Working...";

Office.onReady((info) => {
  if (info.host !== Office.HostType.Outlook) return;
  init();
});

function init() {
  TABS.forEach((t) => $(`tab-${t}`).addEventListener("click", () => showTab(t)));
  $("goto-settings").addEventListener("click", (e) => (e.preventDefault(), showTab("settings")));

  // Email actions
  $("btn-summarize").onclick = () => run(() => withEmail(async (e) => showText(await api.summarize(e))));
  $("btn-draft").onclick = () => run(() => withEmail(draftReply));
  $("btn-triage").onclick = () => run(() => withEmail(async (e) => renderTriage(await api.triage(e))));
  $("btn-spam").onclick = () => run(() => withEmail(spamCheck));
  $("btn-extract").onclick = () => run(() => withEmail(async (e) => renderExtract(await api.extract(e))));
  $("btn-translate").onclick = () => run(() => withEmail(async (e) => showText(await api.translate(e, $("lang").value))));
  $("btn-ask").onclick = () =>
    run(() =>
      withEmail(async (e) => {
        const q = $("question").value.trim();
        if (!q) throw new Error("Type a question first.");
        showText(await api.ask(e, q));
      })
    );

  // Search
  $("btn-search").onclick = () => run(doSearch);
  $("btn-answer").onclick = () => run(answerFromResults);
  $("search-q").addEventListener("keydown", (e) => e.key === "Enter" && $("btn-search").click());

  // Settings
  const s = loadSettings();
  $("api-key").value = s.apiKey;
  $("user-name").value = s.name;
  $("client-id").value = s.clientId;
  $("tenant").value = s.tenant;
  // The web add-in runs on localhost:3000, the Outlook COM add-in on its own virtual host.
  $("redirect-uri").textContent = `${location.origin}/auth.html`;
  $("model").replaceChildren(...MODELS.map((m) => new Option(m.label, m.id, false, m.id === s.model)));
  $("btn-showkey").onclick = toggleKey;
  $("btn-save").onclick = save;
  $("btn-test").onclick = () => run(testKey);
  $("btn-clear").onclick = clearKey;

  Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, onItemChanged);
  onItemChanged();
  refreshSearchHint();
  if (!s.apiKey) {
    showTab("settings");
    showStatus("Welcome. Paste your Claude API key below and choose Save to get started.");
  }
}

// ---------- shell ----------
function showTab(name) {
  TABS.forEach((t) => {
    $(`tab-${t}`).setAttribute("aria-selected", String(t === name));
    $(`panel-${t}`).hidden = t !== name;
  });
  hideStatus();
}

function showStatus(msg, isError) {
  const s = $("status");
  s.textContent = msg;
  s.className = "status" + (isError ? " error" : "");
  s.hidden = false;
}
const hideStatus = () => ($("status").hidden = true);

function lock(on) {
  document.querySelectorAll("button[data-lock]").forEach((b) => {
    b.disabled = on || (b.hasAttribute("data-needs-email") && !mail.hasItem());
  });
}

async function run(fn) {
  clearEmailOutput();
  lock(true);
  showStatus(WORKING);
  try {
    await fn();
    if ($("status").textContent === WORKING) hideStatus();
  } catch (err) {
    showStatus(err.message || String(err), true);
  } finally {
    lock(false);
  }
}

function requireKey() {
  if (!loadSettings().apiKey) {
    showTab("settings");
    throw new Error("Add your Claude API key in Settings first, then try again.");
  }
}

async function withEmail(fn) {
  requireKey();
  return fn(await mail.readEmail());
}

function onItemChanged() {
  clearEmailOutput();
  const it = Office.context.mailbox.item;
  $("subject").textContent = it && it.subject ? it.subject : "No email selected";
  lock(false);
}

// ---------- output helpers ----------
function clearEmailOutput() {
  $("result").hidden = true;
  $("result").replaceChildren();
  $("followup").hidden = true;
  $("followup").replaceChildren();
}

function showResult(node) {
  const r = $("result");
  r.replaceChildren(node);
  r.hidden = false;
  r.scrollIntoView({ block: "nearest" });
}

function showText(text) {
  const r = $("result");
  r.textContent = text;
  r.hidden = false;
  r.scrollIntoView({ block: "nearest" });
}

function el(tag, text, cls) {
  const n = document.createElement(tag);
  if (text != null) n.textContent = text;
  if (cls) n.className = cls;
  return n;
}

function addFollowup(label, onClick, secondary) {
  const b = el("button", label, secondary ? "secondary" : "");
  b.type = "button";
  b.addEventListener("click", async () => {
    try {
      await onClick();
    } catch (err) {
      showStatus(err.message || String(err), true);
    }
  });
  $("followup").append(b);
  $("followup").hidden = false;
}

// Outlook's WebView2 pane can deny navigator.clipboard, so fall back to a hidden textarea.
function legacyCopy(text) {
  const ta = el("textarea");
  ta.value = text;
  ta.style.cssText = "position:fixed;opacity:0";
  document.body.append(ta);
  ta.select();
  try {
    return document.execCommand("copy");
  } finally {
    ta.remove();
  }
}

async function copy(text) {
  try {
    await navigator.clipboard.writeText(text);
    showStatus("Copied.");
  } catch {
    if (legacyCopy(text)) showStatus("Copied.");
    else showStatus("Couldn't copy. Select the text and copy it manually.", true);
  }
}

// ---------- email features ----------
async function draftReply(e) {
  const text = await api.draft(e, { tone: $("tone").value, notes: $("notes").value.trim(), name: loadSettings().name });
  showText(text);
  addFollowup("Open as reply", () => mail.openReply(text));
  addFollowup("Copy", () => copy(text), true);
}

function renderTriage(t) {
  const dl = el("dl");
  [
    ["Priority", t.priority],
    ["Category", t.category],
    ["Needs a reply", t.needs_reply ? "Yes" : "No"],
    ["Deadline", t.deadline || "None found"],
    ["Why", t.reason],
  ].forEach(([k, v]) => dl.append(el("dt", k), el("dd", String(v))));
  showResult(dl);
  const name = `Claude: ${t.category}`;
  addFollowup(`Apply category "${name}"`, async () => {
    await mail.applyCategory(name);
    showStatus(`Applied "${name}".`);
  });
}

const VERDICTS = { safe: "Looks safe", suspicious: "Suspicious", spam: "Spam", phishing: "Likely phishing" };
const LEVELS = { ok: "Pass", warn: "Warning", bad: "Red flag" };

async function spamCheck(e) {
  const signals = await mail.collectSignals();
  const v = await api.spamCheck(e, signals);
  const box = el("div");
  const verdict = VERDICTS[v.verdict] ? v.verdict : "suspicious";
  box.append(el("div", `${VERDICTS[verdict]} (${v.confidence || "unknown"} confidence)`, `verdict ${verdict}`));
  if (v.advice) box.append(el("p", v.advice));
  if (Array.isArray(v.reasons) && v.reasons.length) {
    box.append(el("h3", "Why"));
    const ul = el("ul");
    v.reasons.forEach((r) => ul.append(el("li", String(r))));
    box.append(ul);
  }
  box.append(el("h3", "Technical checks"));
  const checks = el("ul");
  signals.forEach((s) => {
    const li = el("li");
    li.append(el("span", `${LEVELS[s.level]}: `, `lvl ${s.level}`), document.createTextNode(s.text));
    checks.append(li);
  });
  box.append(checks);
  showResult(box);
}

function renderExtract(r) {
  const box = el("div");
  const tasks = r.tasks || [];
  const events = r.events || [];
  box.append(el("h3", "To-dos"));
  if (!tasks.length) box.append(el("p", "None found."));
  else {
    const ul = el("ul");
    tasks.forEach((t) => ul.append(el("li", `${t.task}${t.owner ? ` (${t.owner})` : ""}${t.due ? `, due ${t.due}` : ""}`)));
    box.append(ul);
  }
  box.append(el("h3", "Meetings and dates"));
  if (!events.length) box.append(el("p", "None found."));
  else {
    const ul = el("ul");
    events.forEach((ev) => {
      const li = el("li", `${ev.title}: ${ev.start ? new Date(ev.start).toLocaleString() : "no time"}${ev.location ? `, ${ev.location}` : ""} `);
      const b = el("button", "Add to calendar", "secondary");
      b.type = "button";
      b.addEventListener("click", () => {
        try {
          mail.createEvent(ev);
        } catch (err) {
          showStatus(err.message, true);
        }
      });
      li.append(b);
      ul.append(li);
    });
    box.append(ul);
  }
  showResult(box);
}

// ---------- search ----------
let lastHits = [];
let lastQuery = "";

function refreshSearchHint() {
  $("search-setup").hidden = !!loadSettings().clientId;
}

async function doSearch() {
  const s = loadSettings();
  if (!s.clientId) {
    $("search-setup").hidden = false;
    throw new Error("Set up mail search in Settings first.");
  }
  const q = $("search-q").value.trim();
  if (!q) throw new Error("Type what you're looking for.");
  lastQuery = q;
  $("search-results").replaceChildren();
  $("search-followup").hidden = true;
  $("search-answer").hidden = true;

  // Claude turns the request into search fields; without a key the text is searched as typed.
  const plan = s.apiKey ? await api.searchPlan(q) : { text: q };
  const kql = buildKql(plan);
  if (!kql) throw new Error("Couldn't turn that into a search. Try adding a name or keyword.");
  $("search-plan").textContent = `Searching for: ${kql}`;
  $("search-plan").hidden = false;

  showStatus("Signing in to Microsoft...");
  await graph.getToken(s.clientId, s.tenant);
  showStatus("Searching...");
  const { hits, query } = await graph.searchMessages(plan);
  lastHits = hits;
  $("search-plan").textContent = `Searching for: ${query}`;
  renderHits(lastHits);
  showStatus(lastHits.length ? `${lastHits.length} result${lastHits.length === 1 ? "" : "s"}.` : "No emails matched. Try fewer or broader words.");
}

function renderHits(hits) {
  const ul = $("search-results");
  ul.replaceChildren();
  hits.forEach((h) => {
    const li = el("li");
    li.append(el("div", h.subject, "t"), el("div", `${h.from} - ${h.date}${h.hasAttachments ? " - has attachment" : ""}`, "m"), el("div", h.preview, "p"));
    const b = el("button", "Open", "secondary");
    b.type = "button";
    b.addEventListener("click", () => {
      try {
        mail.openMessage(h.id);
      } catch (err) {
        showStatus(err.message, true);
      }
    });
    li.append(b);
    ul.append(li);
  });
  $("search-followup").hidden = !hits.length || !loadSettings().apiKey;
}

async function answerFromResults() {
  requireKey();
  await graph.getToken(loadSettings().clientId, loadSettings().tenant);
  const top = lastHits.slice(0, 5);
  const emails = await Promise.all(top.map((h) => graph.getMessageText(h.id)));
  const answer = await api.answerFromEmails(lastQuery, emails);
  const box = $("search-answer");
  box.textContent = answer;
  box.hidden = false;
  box.scrollIntoView({ block: "nearest" });
}

// ---------- settings ----------
function toggleKey() {
  const f = $("api-key");
  const show = f.type === "password";
  f.type = show ? "text" : "password";
  $("btn-showkey").textContent = show ? "Hide" : "Show";
}

function save() {
  const ok = saveSettings({
    apiKey: $("api-key").value.trim(),
    model: $("model").value,
    name: $("user-name").value.trim(),
    clientId: $("client-id").value.trim(),
    tenant: $("tenant").value.trim() || "common",
  });
  refreshSearchHint();
  showStatus(ok ? "Saved." : "Couldn't save to disk, so these settings will last only until you close Outlook.", !ok);
}

async function testKey() {
  const key = $("api-key").value.trim();
  if (!key) throw new Error("Paste your API key first.");
  await api.testKey(key, $("model").value);
  showStatus("The key works. Choose Save to keep it.");
}

function clearKey() {
  $("api-key").value = "";
  saveSettings({ apiKey: "" });
  showStatus("API key removed.");
}
