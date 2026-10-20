// Stand-in for office.js inside the Outlook COM add-in. The C# host serves it whenever a page asks for
// appsforoffice.microsoft.com/.../office.js, so public/ runs unchanged. It implements only the parts of the
// Office API this add-in uses, backed by the host through window.chrome.webview messages (see Bridge.cs).
(() => {
  "use strict";
  const host = window.chrome && window.chrome.webview;
  const Succeeded = "succeeded";
  const Failed = "failed";
  const ItemChanged = "olkItemSelectedChanged";
  const DialogMessageReceived = "dialogMessageReceived";
  const DialogEventReceived = "dialogEventReceived";

  const pending = new Map();
  let nextId = 1;
  const itemHandlers = [];
  let dialogHandlers = {};
  let item = null;

  const call = (op, args = {}) =>
    new Promise((resolve, reject) => {
      if (!host) return reject(new Error("This page must run inside the Claude Mail Assistant add-in."));
      const id = nextId++;
      pending.set(id, { resolve, reject });
      host.postMessage({ id, op, args });
    });

  // Office callbacks receive an AsyncResult; the optional "options" argument may sit before the callback.
  const callbackOf = (a, b) => (typeof a === "function" ? a : b);
  const settle = (promise, cb) =>
    promise.then(
      (value) => cb && cb({ status: Succeeded, value }),
      (err) => cb && cb({ status: Failed, error: { code: err.code || 0, name: "Error", message: err.message } })
    );
  const report = (err) => call("showError", { message: err.message || String(err) }).catch(() => {});

  if (host) {
    host.addEventListener("message", (e) => {
      const m = e.data || {};
      if (m.id) {
        const p = pending.get(m.id);
        if (!p) return;
        pending.delete(m.id);
        if (m.error) p.reject(Object.assign(new Error(m.error.message), { code: m.error.code }));
        else p.resolve(m.result);
      } else if (m.event === "itemChanged") {
        item = makeItem(m.item);
        itemHandlers.forEach((h) => h({ type: ItemChanged }));
      } else if (m.event === "dialogMessage" && dialogHandlers[DialogMessageReceived]) {
        dialogHandlers[DialogMessageReceived]({ type: DialogMessageReceived, message: m.message, origin: location.origin });
      } else if (m.event === "dialogEvent" && dialogHandlers[DialogEventReceived]) {
        dialogHandlers[DialogEventReceived]({ type: DialogEventReceived, error: m.error });
      }
    });
  }

  function makeItem(s) {
    if (!s) return null;
    const entryId = s.entryId;
    return {
      itemType: "message",
      subject: s.subject,
      from: { displayName: s.fromName, emailAddress: s.fromAddress },
      dateTimeCreated: s.created ? new Date(s.created) : null,
      attachments: (s.attachments || []).map((name) => ({ name })),
      body: {
        getAsync: (type, a, b) => settle(call("getBody", { entryId, html: type === "html" }), callbackOf(a, b)),
      },
      getAllInternetHeadersAsync: (a, b) => settle(call("getHeaders", { entryId }), callbackOf(a, b)),
      displayReplyForm: (form) => {
        const html = typeof form === "string" ? form : (form && form.htmlBody) || "";
        call("reply", { entryId, html }).catch(report);
      },
      categories: {
        addAsync: (names, a, b) => settle(call("addCategory", { entryId, names }), callbackOf(a, b)),
      },
    };
  }

  // Search results carry Graph (REST) ids, but Outlook opens items by MAPI EntryID. graph.js's token is
  // private to it, so remember the last one it sent and use it to look up the EntryID.
  let graphAuth = null;
  const realFetch = window.fetch.bind(window);
  window.fetch = (input, init) => {
    try {
      const url = String(input && input.url ? input.url : input);
      const headers = (init && init.headers) || {};
      const auth = headers.Authorization || headers.authorization;
      if (url.startsWith("https://graph.microsoft.com/") && auth) graphAuth = auth;
    } catch {
      /* never break the real request */
    }
    return realFetch(input, init);
  };

  async function openRestId(restId) {
    if (!graphAuth) throw new Error("Run the search again, then choose Open.");
    const url = new URL(`https://graph.microsoft.com/v1.0/me/messages/${encodeURIComponent(restId)}`);
    url.searchParams.set("$select", "id");
    url.searchParams.set("$expand", "singleValueExtendedProperties($filter=id eq 'Binary 0x0FFF')"); // PR_ENTRYID
    const res = await realFetch(url, { headers: { Authorization: graphAuth } });
    const data = await res.json().catch(() => ({}));
    const prop = (data.singleValueExtendedProperties || [])[0];
    if (!res.ok || !prop) throw new Error("Couldn't open that email from here. Find it in Outlook instead.");
    const bytes = atob(prop.value.replace(/-/g, "+").replace(/_/g, "/"));
    const hex = Array.from(bytes, (c) => c.charCodeAt(0).toString(16).padStart(2, "0")).join("").toUpperCase();
    await call("openItem", { entryId: hex });
  }

  const iso = (d) => (d instanceof Date ? d.toISOString() : d ? new Date(d).toISOString() : "");

  const ready = host
    ? call("hello").then((info) => {
        if (info && info.role === "pane") item = makeItem(info.item);
        return { host: "Outlook", platform: "PC" };
      })
    : Promise.resolve({ host: null, platform: null });

  window.Office = {
    HostType: { Outlook: "Outlook" },
    PlatformType: { PC: "PC" },
    AsyncResultStatus: { Succeeded, Failed },
    CoercionType: { Text: "text", Html: "html" },
    EventType: { ItemChanged, DialogMessageReceived, DialogEventReceived },
    MailboxEnums: { RestVersion: { v1_0: "v1.0", v2_0: "v2.0", Beta: "beta" } },

    onReady: (cb) => (cb ? ready.then(cb) : ready),

    context: {
      requirements: { isSetSupported: () => true },
      mailbox: {
        get item() {
          return item;
        },
        addHandlerAsync(type, handler, a, b) {
          if (type === ItemChanged) itemHandlers.push(handler);
          const cb = callbackOf(a, b);
          if (cb) cb({ status: Succeeded });
        },
        convertToEwsId: (id) => id, // the host resolves Graph ids itself, see openRestId
        displayMessageForm: (id) => {
          openRestId(id).catch(report);
        },
        displayNewAppointmentForm: (p) => {
          const start = iso(p.start);
          call("newAppointment", { subject: p.subject || "", start, end: iso(p.end) || start, location: p.location || "", body: p.body || "" }).catch(report);
        },
      },
      ui: {
        displayDialogAsync(url, options, callback) {
          if (typeof options === "function") [callback, options] = [options, {}];
          options = options || {};
          const dialog = {
            addEventHandler(type, handler) {
              dialogHandlers[type] = handler;
            },
            close() {
              dialogHandlers = {};
              call("closeDialog").catch(() => {});
            },
          };
          const args = { url: new URL(url, location.href).href, width: options.width || 30, height: options.height || 60 };
          settle(call("openDialog", args).then(() => dialog), callback);
        },
        messageParent(message) {
          call("messageParent", { message: String(message) }).catch(() => {});
        },
      },
    },
  };
})();
