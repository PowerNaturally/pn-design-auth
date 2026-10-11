import { PublicClientApplication, InteractionRequiredAuthError } from "@azure/msal-browser";

const CONFIG = {
  tenantId: "88df17e8-4f1c-45a7-8ed9-6571a1161e90",
  clientId: "30e376f3-2fda-4d28-ad73-bf386c51ad02",
  scope: "api://d7901a65-c502-441c-b38e-3dd03d3f6283/Calculations.Access",
  redirectUri: "https://orange-mushroom-08238f410.5.azurestaticapps.net/broker.html",
  simproOrigin: "https://powernaturally.simprosuite.com"
};

const status = document.getElementById("status");
const statusOnly = new URLSearchParams(window.location.search).get("mode") === "status";
const pca = new PublicClientApplication({
  auth: {
    clientId: CONFIG.clientId,
    authority: `https://login.microsoftonline.com/${CONFIG.tenantId}`,
    redirectUri: CONFIG.redirectUri,
    navigateToLoginRequestUrl: false
  },
  cache: { cacheLocation: "localStorage" }
});

function targetWindow() {
  if (statusOnly && window.parent && window.parent !== window) return window.parent;
  if (window.opener && !window.opener.closed) return window.opener;
  throw new Error("The SimPro window is no longer available.");
}
function send(type, payload = {}) {
  targetWindow().postMessage({ source: "pn-design-auth", type, ...payload }, CONFIG.simproOrigin);
}
function returnResult(result, account, type) {
  send(type, {
    authenticated: true,
    accessToken: result.accessToken,
    expiresOn: result.expiresOn ? result.expiresOn.toISOString() : null,
    username: account?.username || ""
  });
  window.setTimeout(() => {
    if (statusOnly) return;
    window.close();
  }, 100);
}
async function run() {
  await pca.initialize();
  const redirectResult = await pca.handleRedirectPromise();
  if (redirectResult?.account) pca.setActiveAccount(redirectResult.account);
  const account = pca.getActiveAccount() || pca.getAllAccounts()[0] || null;

  if (statusOnly) {
    if (!account) {
      send("PN_AUTH_STATUS", { authenticated: false });
      return;
    }
    try {
      const result = await pca.acquireTokenSilent({ account, scopes: [CONFIG.scope] });
      returnResult(result, account, "PN_AUTH_STATUS");
    } catch {
      send("PN_AUTH_STATUS", { authenticated: false, username: account.username || "" });
    }
    return;
  }

  if (!account) {
    status.textContent = "Redirecting to Microsoft 365 sign-in...";
    await pca.loginRedirect({ scopes: [CONFIG.scope], prompt: "select_account" });
    return;
  }
  try {
    const result = redirectResult?.accessToken
      ? redirectResult
      : await pca.acquireTokenSilent({ account, scopes: [CONFIG.scope] });
    returnResult(result, account, "PN_AUTH_SUCCESS");
  } catch (error) {
    if (error instanceof InteractionRequiredAuthError) {
      status.textContent = "Additional Microsoft 365 authorisation is required...";
      await pca.acquireTokenRedirect({ account, scopes: [CONFIG.scope] });
      return;
    }
    throw error;
  }
}
run().catch(error => {
  console.error("PowerNaturally authentication broker failed", error);
  if (status) status.textContent = `Authentication failed: ${error.message}`;
  try { send("PN_AUTH_ERROR", { message: error.message }); } catch {}
});
