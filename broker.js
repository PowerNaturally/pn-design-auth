import { PublicClientApplication, InteractionRequiredAuthError } from "@azure/msal-browser";

const CONFIG = {
  tenantId: "88df17e8-4f1c-45a7-8ed9-6571a1161e90",
  clientId: "30e376f3-2fda-4d28-ad73-bf386c51ad02",
  scope: "api://d7901a65-c502-441c-b38e-3dd03d3f6283/Calculations.Access",
  redirectUri: "https://orange-mushroom-08238f410.5.azurestaticapps.net/broker.html",
  simproOrigin: "https://powernaturally.simprosuite.com"
};

const status = document.getElementById("status");
const parameters = new URLSearchParams(window.location.search);
const statusOnly = parameters.get("mode") === "status";
const loginHint = (parameters.get("login_hint") || "").trim();

const pca = new PublicClientApplication({
  auth: {
    clientId: CONFIG.clientId,
    authority: `https://login.microsoftonline.com/${CONFIG.tenantId}`,
    redirectUri: CONFIG.redirectUri,
    navigateToLoginRequestUrl: false
  },
  cache: {
    cacheLocation: "localStorage"
  },
  system: {
    iframeHashTimeout: 15000
  }
});

function targetWindow() {
  if (statusOnly && window.parent && window.parent !== window) return window.parent;
  if (window.opener && !window.opener.closed) return window.opener;
  throw new Error("The SimPro window is no longer available.");
}

function send(type, payload = {}) {
  targetWindow().postMessage(
    { source: "pn-design-auth", type, ...payload },
    CONFIG.simproOrigin
  );
}

function returnResult(result, account, type) {
  send(type, {
    authenticated: true,
    accessToken: result.accessToken,
    expiresOn: result.expiresOn ? result.expiresOn.toISOString() : null,
    username: account?.username || result.account?.username || ""
  });

  window.setTimeout(() => {
    if (!statusOnly) window.close();
  }, 100);
}

function unauthenticated(reason, account = null, error = null) {
  send("PN_AUTH_STATUS", {
    authenticated: false,
    username: account?.username || loginHint || "",
    reason,
    errorCode: error?.errorCode || error?.code || ""
  });
}

async function silentFromExistingEntraSession() {
  const request = {
    scopes: [CONFIG.scope],
    prompt: "none"
  };

  if (loginHint) request.loginHint = loginHint;

  const result = await pca.ssoSilent(request);
  if (result?.account) pca.setActiveAccount(result.account);
  return result;
}

async function tokenFromCache(account) {
  return pca.acquireTokenSilent({
    account,
    scopes: [CONFIG.scope]
  });
}

async function runStatusMode(cachedAccount) {
  if (cachedAccount) {
    try {
      const result = await tokenFromCache(cachedAccount);
      returnResult(result, cachedAccount, "PN_AUTH_STATUS");
      return;
    } catch (error) {
      console.info("Cached broker account could not provide a token; trying Entra SSO.", error);
    }
  }

  try {
    const result = await silentFromExistingEntraSession();
    returnResult(result, result.account, "PN_AUTH_STATUS");
  } catch (error) {
    console.info("Silent Microsoft Entra SSO was unavailable.", error);
    unauthenticated("silent_sso_unavailable", cachedAccount, error);
  }
}

async function run() {
  await pca.initialize();

  const redirectResult = await pca.handleRedirectPromise();
  if (redirectResult?.account) pca.setActiveAccount(redirectResult.account);

  const account = pca.getActiveAccount() || pca.getAllAccounts()[0] || null;

  if (statusOnly) {
    await runStatusMode(account);
    return;
  }

  if (!account) {
    status.textContent = "Redirecting to Microsoft 365 sign-in...";
    await pca.loginRedirect({
      scopes: [CONFIG.scope],
      prompt: "select_account"
    });
    return;
  }

  try {
    const result = redirectResult?.accessToken
      ? redirectResult
      : await tokenFromCache(account);

    returnResult(result, account, "PN_AUTH_SUCCESS");
  } catch (error) {
    if (error instanceof InteractionRequiredAuthError) {
      status.textContent = "Additional Microsoft 365 authorisation is required...";
      await pca.acquireTokenRedirect({
        account,
        scopes: [CONFIG.scope]
      });
      return;
    }
    throw error;
  }
}

run().catch(error => {
  console.error("PowerNaturally authentication broker failed", error);
  if (status) status.textContent = `Authentication failed: ${error.message}`;
  try {
    send("PN_AUTH_ERROR", {
      message: error.message,
      errorCode: error.errorCode || error.code || ""
    });
  } catch {}
});