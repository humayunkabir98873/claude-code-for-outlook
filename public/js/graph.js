/* global Office */
import { buildKql, planDateRange } from "./util.js";

let token = null;
let expiresAt = 0;

const DIALOG_ERRORS = {
  12002: "The sign-in page couldn't be loaded. Make sure the add-in server (start.bat) is running.",
  12003: "The sign-in window was sent to an insecure (http) page.",
  12004: "Outlook blocked the sign-in window because its address isn't trusted by the add-in.",
  12007: "A sign-in window is already open. Finish or close it first.",
  12009: "The sign-in window was blocked. Allow pop-ups for Outlook and try again.",
  12011: "Your browser settings blocked the sign-in window. Allow pop-ups for Outlook and try again.",
};
const dialogError = (code, fallback) => new Error(DIALOG_ERRORS[code] || fallback);

// Opens auth.html in an Office dialog, which signs in with MSAL and sends the token back.
export function getToken(clientId, tenant) {
  if (token && Date.now() < expiresAt - 60_000) return Promise.resolve(token);
  return new Promise((resolve, reject) => {
    const url = `${location.origin}/auth.html?clientId=${encodeURIComponent(clientId)}&tenant=${encodeURIComponent(tenant || "common")}`;
    Office.context.ui.displayDialogAsync(url, { height: 60, width: 30 }, (r) => {
      if (r.status !== Office.AsyncResultStatus.Succeeded) return reject(dialogError(r.error.code, r.error.message));
      const dialog = r.value;
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
        dialog.close();
        let msg;
        try {
          msg = JSON.parse(arg.message);
        } catch {
          return reject(new Error("Sign-in returned an unreadable response."));
        }
        if (msg.error) return reject(new Error(msg.error));
        token = msg.token;
        expiresAt = msg.expiresOn ? new Date(msg.expiresOn).getTime() : Date.now() + 50 * 60_000;
        resolve(token);
      });
      dialog.addEventHandler(Office.EventType.DialogEventReceived, (arg) =>
        reject(dialogError(arg.error, "The sign-in window was closed before finishing."))
      );
    });
  });
}

async function graphFetch(path, params, extraHeaders = {}) {
  const url = new URL(`https://graph.microsoft.com/v1.0${path}`);
  for (const [k, v] of Object.entries(params || {})) url.searchParams.set(k, v);
  let res;
  try {
    res = await fetch(url, { headers: { Authorization: `Bearer ${token}`, ...extraHeaders } });
  } catch {
    throw new Error("Couldn't reach Microsoft Graph. Check your connection.");
  }
  if (res.status === 401) {
    token = null;
    throw new Error("Your Microsoft sign-in expired. Run the search again to sign in.");
  }
  const data = await res.json().catch(() => ({}));
  if (!res.ok) throw Object.assign(new Error(data?.error?.message || `Mail search failed (${res.status})`), { status: res.status });
  return data;
}

const SELECT = "id,subject,from,receivedDateTime,bodyPreview,hasAttachments";
const MAX_HITS = 15;

const toHit = (m) => ({
  id: m.id,
  subject: m.subject || "(no subject)",
  from: m.from?.emailAddress ? `${m.from.emailAddress.name || ""} <${m.from.emailAddress.address}>`.trim() : "",
  date: m.receivedDateTime ? new Date(m.receivedDateTime).toLocaleString() : "",
  preview: m.bodyPreview || "",
  hasAttachments: !!m.hasAttachments,
});

// Returns { hits, query } where query is what was actually sent, for display.
export async function searchMessages(plan) {
  const kql = buildKql(plan);
  const { from, to } = planDateRange(plan);
  try {
    const data = await graphFetch("/me/messages", { $search: `"${kql}"`, $top: String(MAX_HITS), $select: SELECT });
    return { hits: (data.value || []).map(toHit), query: kql };
  } catch (err) {
    if (err.status !== 400 || (!from && !to && !plan.hasAttachment)) throw err;
  }

  // Graph rejected the received>= / hasAttachments KQL. Search the words only and apply the
  // date and attachment conditions here, or use $filter when there are no words at all.
  const keep = (m) =>
    (!from || new Date(m.receivedDateTime) >= from) && (!to || new Date(m.receivedDateTime) < to) && (!plan.hasAttachment || m.hasAttachments);
  const words = buildKql(plan, { filters: false });
  if (words) {
    const data = await graphFetch("/me/messages", { $search: `"${words}"`, $top: "100", $select: SELECT });
    return { hits: (data.value || []).filter(keep).slice(0, MAX_HITS).map(toHit), query: `${words} (dates and attachments checked locally)` };
  }
  // $orderby properties must also appear first in $filter, so receivedDateTime always leads.
  const filter = [`receivedDateTime ge ${(from || new Date(0)).toISOString()}`];
  if (to) filter.push(`receivedDateTime lt ${to.toISOString()}`);
  if (plan.hasAttachment) filter.push("hasAttachments eq true");
  const data = await graphFetch("/me/messages", {
    $filter: filter.join(" and "),
    $orderby: "receivedDateTime desc",
    $top: String(MAX_HITS),
    $select: SELECT,
  });
  return { hits: (data.value || []).map(toHit), query: filter.join(" and ") };
}

export async function getMessageText(id) {
  const m = await graphFetch(`/me/messages/${encodeURIComponent(id)}`, { $select: "subject,from,receivedDateTime,body" }, { Prefer: 'outlook.body-content-type="text"' });
  return {
    subject: m.subject || "",
    from: m.from?.emailAddress ? `${m.from.emailAddress.name || ""} <${m.from.emailAddress.address}>` : "",
    date: m.receivedDateTime || "",
    body: m.body?.content || "",
  };
}
