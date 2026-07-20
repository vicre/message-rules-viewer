import {
  InteractionRequiredAuthError,
  PublicClientApplication,
} from "@azure/msal-browser";

const TENANT_ID = "f251f123-c9ce-448e-9277-34bb285911d9";

const FRONTEND_CLIENT_ID =
  "07021449-4d20-419d-9b3a-83838f937de8";

const BACKEND_CLIENT_ID =
  "ea452940-e61c-40a7-ac8d-285861986423";

export const API_SCOPE =
  `api://${BACKEND_CLIENT_ID}/access_as_user`;

export const msalInstance = new PublicClientApplication({
  auth: {
    clientId: FRONTEND_CLIENT_ID,
    authority: `https://login.microsoftonline.com/${TENANT_ID}`,
    redirectUri: window.location.origin,
    postLogoutRedirectUri: window.location.origin,
    navigateToLoginRequestUrl: true,
  },

  cache: {
    cacheLocation: "sessionStorage",
    storeAuthStateInCookie: false,
  },
});

let initializationPromise = null;

export async function initializeAuthentication() {
  if (!initializationPromise) {
    initializationPromise = initializeMsal();
  }

  return initializationPromise;
}

async function initializeMsal() {
  await msalInstance.initialize();

  /*
   * This must run before loginRedirect(), acquireTokenRedirect(),
   * or any other interactive authentication call.
   */
  const redirectResult =
    await msalInstance.handleRedirectPromise();

  if (redirectResult?.account) {
    msalInstance.setActiveAccount(
      redirectResult.account,
    );

    return redirectResult.account;
  }

  const activeAccount =
    msalInstance.getActiveAccount();

  if (activeAccount) {
    return activeAccount;
  }

  const accounts =
    msalInstance.getAllAccounts();

  if (accounts.length > 0) {
    msalInstance.setActiveAccount(accounts[0]);
    return accounts[0];
  }

  return null;
}

export async function signIn() {
  await initializeAuthentication();

  /*
   * This navigates away from the page.
   * It normally does not return because Entra redirects the browser.
   */
  await msalInstance.loginRedirect({
    scopes: [API_SCOPE],
    prompt: "select_account",
  });
}

export async function getBackendAccessToken() {
  await initializeAuthentication();

  const account =
    msalInstance.getActiveAccount();

  if (!account) {
    throw new Error(
      "You must sign in before querying mailbox rules.",
    );
  }

  const tokenRequest = {
    account,
    scopes: [API_SCOPE],
  };

  try {
    const tokenResult =
      await msalInstance.acquireTokenSilent(
        tokenRequest,
      );

    return tokenResult.accessToken;
  } catch (error) {
    if (
      error instanceof InteractionRequiredAuthError ||
      error?.errorCode === "interaction_required" ||
      error?.errorCode === "consent_required" ||
      error?.errorCode === "login_required"
    ) {
      /*
       * Store the current page so the application can continue
       * after the token redirect.
       */
      sessionStorage.setItem(
        "messageRules.resumeQuery",
        "true",
      );

      await msalInstance.acquireTokenRedirect(
        tokenRequest,
      );

      /*
       * acquireTokenRedirect navigates away, so execution
       * should not normally reach this point.
       */
      throw new Error(
        "Authentication redirect started.",
      );
    }

    throw error;
  }
}

export async function signOut() {
  await initializeAuthentication();

  const account =
    msalInstance.getActiveAccount();

  await msalInstance.logoutRedirect({
    account: account ?? undefined,
    postLogoutRedirectUri:
      window.location.origin,
  });
}