import {
  InteractionRequiredAuthError,
  type AccountInfo,
  PublicClientApplication
} from "@azure/msal-browser";

const authority =
  import.meta.env.VITE_MICROSOFT_AUTHORITY ??
  "https://login.microsoftonline.com/common";
const clientId = import.meta.env.VITE_MICROSOFT_CLIENT_ID ?? "";
const apiScope = import.meta.env.VITE_MICROSOFT_API_SCOPE ?? "";
const redirectUri = import.meta.env.VITE_MICROSOFT_REDIRECT_URI ?? `${window.location.origin}/`;

export const msalInstance = new PublicClientApplication({
  auth: {
    clientId,
    authority,
    redirectUri,
    postLogoutRedirectUri: redirectUri
  },
  cache: {
    cacheLocation: "localStorage"
  }
});

const loginScopes = ["openid", "profile", "offline_access", apiScope].filter(Boolean);

export async function initializeMicrosoftAuth() {
  await msalInstance.initialize();
  const redirectResult = await msalInstance.handleRedirectPromise();
  const activeAccount =
    redirectResult?.account ?? msalInstance.getActiveAccount() ?? msalInstance.getAllAccounts()[0] ?? null;
  if (activeAccount) {
    msalInstance.setActiveAccount(activeAccount);
  }
  return activeAccount;
}

export function getMicrosoftAccount() {
  return msalInstance.getActiveAccount() ?? msalInstance.getAllAccounts()[0] ?? null;
}

export async function loginWithMicrosoft() {
  await msalInstance.loginRedirect({
    scopes: loginScopes
  });
}

export async function logoutFromMicrosoft() {
  const account = getMicrosoftAccount();
  await msalInstance.logoutRedirect({
    account: account ?? undefined
  });
}

export async function acquireMicrosoftApiToken() {
  const account = getMicrosoftAccount();
  if (!account) {
    throw new Error("Microsoft is not connected for this browser session.");
  }

  try {
    const result = await msalInstance.acquireTokenSilent({
      account,
      scopes: [apiScope]
    });
    return result.accessToken;
  } catch (error) {
    if (error instanceof InteractionRequiredAuthError) {
      await msalInstance.acquireTokenRedirect({
        account,
        scopes: [apiScope]
      });
    }
    throw error;
  }
}

export type MicrosoftAccount = AccountInfo;
