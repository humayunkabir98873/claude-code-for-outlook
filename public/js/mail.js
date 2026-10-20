/* global Office */
import { parseHeaders, analyzeHeaders, analyzeLinks, analyzeAttachments } from "./util.js";

const item = () => Office.context.mailbox.item;

function getBody(type) {
  return new Promise((resolve, reject) => {
    item().body.getAsync(type, (r) => {
      if (r.status === Office.AsyncResultStatus.Succeeded) resolve(r.value);
      else reject(new Error(r.error.message));
    });
  });
}

function getRawHeaders() {
  return new Promise((resolve) => {
    if (!Office.context.requirements.isSetSupported("Mailbox", "1.8") || !item().getAllInternetHeadersAsync) return resolve(null);
    item().getAllInternetHeadersAsync((r) => resolve(r.status === Office.AsyncResultStatus.Succeeded ? r.value : null));
  });
}

export const hasItem = () => !!(Office.context.mailbox && item());

export async function readEmail() {
  const it = item();
  if (!it) throw new Error("Select an email first.");
  const from = it.from ? `${it.from.displayName} <${it.from.emailAddress}>` : "";
  return {
    subject: it.subject,
    from,
    date: it.dateTimeCreated ? it.dateTimeCreated.toString() : "",
    body: await getBody(Office.CoercionType.Text),
  };
}

export async function collectSignals() {
  const it = item();
  const raw = await getRawHeaders();
  const headers = raw ? parseHeaders(raw) : null;
  const signals = analyzeHeaders(headers, it.from?.displayName || "", it.from?.emailAddress || "");

  let links = { signals: [], hosts: [] };
  try {
    links = analyzeLinks(await getBody(Office.CoercionType.Html));
  } catch {
    links.signals.push({ level: "warn", text: "Couldn't read the message's links" });
  }
  signals.push(...links.signals);
  signals.push(...analyzeAttachments((it.attachments || []).map((a) => a.name)));
  return signals;
}

export function openReply(text) {
  const html = text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/\r?\n/g, "<br>");
  item().displayReplyForm(html);
}

export function openMessage(restId) {
  const ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
  Office.context.mailbox.displayMessageForm(ewsId);
}

export function createEvent(ev) {
  const start = new Date(ev.start);
  if (isNaN(start)) throw new Error("Claude couldn't work out a valid start time for this event.");
  let end = ev.end ? new Date(ev.end) : null;
  if (!end || isNaN(end) || end <= start) end = new Date(start.getTime() + 60 * 60 * 1000);
  Office.context.mailbox.displayNewAppointmentForm({
    subject: ev.title || "Meeting",
    start,
    end,
    location: ev.location || "",
    body: `Created from: ${item().subject || ""}`,
  });
}

export function applyCategory(name) {
  return new Promise((resolve, reject) => {
    if (!Office.context.requirements.isSetSupported("Mailbox", "1.8")) {
      return reject(new Error("This version of Outlook can't apply categories from add-ins."));
    }
    item().categories.addAsync([name], (r) => {
      if (r.status === Office.AsyncResultStatus.Succeeded) resolve();
      else reject(new Error(`Couldn't apply the category. Create "${name}" in Outlook first, then try again.`));
    });
  });
}
