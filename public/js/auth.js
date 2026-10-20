/* global Office, msal */
const send = (obj) => Office.context.ui.messageParent(JSON.stringify(obj));

Office.onReady(async () => {
  try {
    // Microsoft redirects back to plain /auth.html (no query), so keep the config in sessionStorage.
    // Fresh query parameters always win over a config left behind by an abandoned sign-in.
    const p = new URLSearchParams(location.search);
    const cfg = p.get("clientId")
      ? { clientId: p.get("clientId"), tenant: p.get("tenant") || "common" }
      : JSON.parse(sessionStorage.getItem("authCfg") || "null");
    if (!cfg?.clientId) throw new Error("No Azure client ID is set. Add it in Settings.");
    sessionStorage.setItem("authCfg", JSON.stringify(cfg));

    const app = new msal.PublicClientApplication({
      auth: {
        clientId: cfg.clientId,
        authority: `https://login.microsoftonline.com/${cfg.tenant}`,
        redirectUri: `${location.origin}/auth.html`,
      },
      cache: { cacheLocation: "sessionStorage" },
    });
    await app.initialize();

    // Handle the response right here instead of letting MSAL bounce back to the original ?clientId= URL.
    const result = await app.handleRedirectPromise({ navigateToLoginRequestUrl: false });
    if (result) {
      sessionStorage.removeItem("authCfg");
      return send({ token: result.accessToken, expiresOn: result.expiresOn });
    }
    await app.loginRedirect({ scopes: ["Mail.Read"], prompt: "select_account" });
  } catch (err) {
    sessionStorage.removeItem("authCfg");
    const code = err.errorCode ? ` (${err.errorCode})` : "";
    send({ error: (err.errorMessage || err.message || String(err)) + code });
  }
});
