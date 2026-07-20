import "./style.css";

import {
  getBackendAccessToken,
  initializeAuthentication,
  signIn,
  signOut,
} from "./auth.js";

const STORAGE_KEY =
  "messageRules.lastUserPrincipalName";

const API_BASE_URL =
  import.meta.env.VITE_API_BASE_URL ??
  "http://localhost:3001";

document.querySelector("#app").innerHTML = `
  <main class="container">
    <header class="page-header">
      <div>
        <h1>Mailbox message rules</h1>
        <p class="subtitle">
          Query Inbox rules for a DTU mailbox.
        </p>
      </div>

      <div class="auth-panel">
        <div id="signed-in-user" class="signed-in-user" hidden>
          <span class="user-label">Signed in as</span>
          <strong id="user-name"></strong>
        </div>

        <button
          id="login-button"
          class="primary-button"
          type="button"
        >
          Sign in with DTU
        </button>

        <button
          id="logout-button"
          class="secondary-button"
          type="button"
          hidden
        >
          Sign out
        </button>
      </div>
    </header>

    <section
      id="query-section"
      class="query-section"
      hidden
    >
      <form id="search-form">
        <label for="upn">
          User principal name
        </label>

        <div class="search-row">
          <input
            id="upn"
            name="upn"
            type="email"
            placeholder="kijens@dtu.dk"
            autocomplete="off"
            spellcheck="false"
            required
          />

          <button
            id="query-button"
            class="primary-button"
            type="submit"
          >
            Query rules
          </button>
        </div>
      </form>
    </section>

    <p
      id="status"
      class="status"
      role="status"
      aria-live="polite"
    ></p>

    <section
      id="results"
      class="results"
      aria-live="polite"
    ></section>
  </main>
`;

const loginButton =
  document.querySelector("#login-button");

const logoutButton =
  document.querySelector("#logout-button");

const queryButton =
  document.querySelector("#query-button");

const querySection =
  document.querySelector("#query-section");

const signedInUser =
  document.querySelector("#signed-in-user");

const userNameElement =
  document.querySelector("#user-name");

const form =
  document.querySelector("#search-form");

const upnInput =
  document.querySelector("#upn");

const statusElement =
  document.querySelector("#status");

const resultsElement =
  document.querySelector("#results");

upnInput.value =
  localStorage.getItem(STORAGE_KEY) ?? "";

loginButton.addEventListener(
  "click",
  handleLogin,
);

logoutButton.addEventListener(
  "click",
  handleLogout,
);

form.addEventListener(
  "submit",
  handleQuery,
);

await startApplication();

/**
 * Checks whether MSAL already has an authenticated account.
 */
async function startApplication() {
  setStatus("Checking sign-in status…");

  try {
    const account =
      await initializeAuthentication();

    if (account) {
      showAuthenticatedState(account);
      setStatus("");
      return;
    }

    showUnauthenticatedState();
    setStatus(
      "Sign in with your DTU account to continue.",
    );
  } catch (error) {
    console.error(
      "Authentication initialization failed:",
      error,
    );

    showUnauthenticatedState();
    setStatus(getErrorMessage(error), true);
  }
}

/**
 * Handles the Sign in with DTU button.
 */
async function handleLogin() {
  setLoginLoading(true);
  setStatus("Opening Microsoft sign-in…");

  try {
    const account = await signIn();

    showAuthenticatedState(account);
    clearResults();
    setStatus("Signed in successfully.");
  } catch (error) {
    console.error("Sign-in failed:", error);

    if (
      error?.errorCode === "user_cancelled" ||
      error?.errorCode === "popup_window_error"
    ) {
      setStatus("Sign-in was cancelled.");
      return;
    }

    setStatus(getErrorMessage(error), true);
  } finally {
    setLoginLoading(false);
  }
}

/**
 * Handles the Sign out button.
 */
async function handleLogout() {
  logoutButton.disabled = true;
  setStatus("Signing out…");

  try {
    await signOut();

    showUnauthenticatedState();
    clearResults();
    setStatus("Signed out.");
  } catch (error) {
    console.error("Sign-out failed:", error);
    setStatus(getErrorMessage(error), true);
  } finally {
    logoutButton.disabled = false;
  }
}

/**
 * Gets a backend access token and queries the Express API.
 */
async function handleQuery(event) {
  event.preventDefault();

  const upn =
    upnInput.value.trim().toLowerCase();

  if (!isValidDtuUpn(upn)) {
    setStatus(
      "Enter a valid address ending in @dtu.dk.",
      true,
    );

    upnInput.focus();
    return;
  }

  localStorage.setItem(STORAGE_KEY, upn);

  setQueryLoading(true);
  clearResults();

  setStatus(
    `Querying message rules for ${upn}…`,
  );

  try {
    const accessToken =
      await getBackendAccessToken();

    const endpoint =
      `${API_BASE_URL}/message-rules` +
      `?upn=${encodeURIComponent(upn)}`;

    const response = await fetch(endpoint, {
      method: "GET",

      headers: {
        Authorization: `Bearer ${accessToken}`,
        Accept: "application/json",
      },
    });

    const result =
      await readResponseBody(response);

    if (!response.ok) {
      throw new Error(
        result?.error ??
          result?.message ??
          `The server returned HTTP ${response.status}.`,
      );
    }

    const rules =
      Array.isArray(result?.rules)
        ? result.rules
        : [];

    renderRules(rules);

    setStatus(
      `${rules.length} rule(s) found for ` +
        `${result?.userPrincipalName ?? upn}.`,
    );
  } catch (error) {
    console.error(
      "Mailbox rule query failed:",
      error,
    );

    setStatus(getErrorMessage(error), true);
  } finally {
    setQueryLoading(false);
  }
}

function showAuthenticatedState(account) {
  const accountName =
    account?.username ??
    account?.name ??
    "Authenticated user";

  userNameElement.textContent = accountName;

  signedInUser.hidden = false;
  loginButton.hidden = true;
  logoutButton.hidden = false;
  querySection.hidden = false;

  upnInput.focus();
}

function showUnauthenticatedState() {
  userNameElement.textContent = "";

  signedInUser.hidden = true;
  loginButton.hidden = false;
  logoutButton.hidden = true;
  querySection.hidden = true;
}

function setLoginLoading(isLoading) {
  loginButton.disabled = isLoading;

  loginButton.textContent = isLoading
    ? "Signing in…"
    : "Sign in with DTU";
}

function setQueryLoading(isLoading) {
  queryButton.disabled = isLoading;
  upnInput.disabled = isLoading;

  queryButton.textContent = isLoading
    ? "Loading…"
    : "Query rules";
}

function setStatus(message, isError = false) {
  statusElement.textContent = message;
  statusElement.classList.toggle(
    "error",
    isError,
  );
}

function clearResults() {
  resultsElement.replaceChildren();
}

function isValidDtuUpn(value) {
  return /^[a-z0-9._'-]+@dtu\.dk$/i.test(value);
}

async function readResponseBody(response) {
  const contentType =
    response.headers.get("content-type") ?? "";

  if (
    contentType.includes("application/json")
  ) {
    return response.json();
  }

  const text = await response.text();

  return text
    ? { error: text }
    : {};
}

function getErrorMessage(error) {
  if (
    error instanceof TypeError &&
    error.message.toLowerCase().includes("fetch")
  ) {
    return (
      "The backend could not be reached. " +
      "Confirm that it is running on " +
      `${API_BASE_URL}.`
    );
  }

  if (
    error instanceof Error &&
    error.message
  ) {
    return error.message;
  }

  return "An unexpected error occurred.";
}

function renderRules(rules) {
  clearResults();

  if (rules.length === 0) {
    const emptyMessage =
      document.createElement("p");

    emptyMessage.className = "empty-result";
    emptyMessage.textContent =
      "No Inbox message rules were found.";

    resultsElement.append(emptyMessage);
    return;
  }

  for (const rule of rules) {
    resultsElement.append(
      createRuleCard(rule),
    );
  }
}

function createRuleCard(rule) {
  const article =
    document.createElement("article");

  article.className = "rule";

  const heading =
    document.createElement("h2");

  heading.textContent =
    rule.displayName || "Unnamed rule";

  const metadata =
    document.createElement("dl");

  metadata.className = "rule-metadata";

  appendMetadata(
    metadata,
    "Sequence",
    rule.sequence ?? "Unknown",
  );

  appendMetadata(
    metadata,
    "Enabled",
    rule.isEnabled ? "Yes" : "No",
  );

  appendMetadata(
    metadata,
    "Has error",
    rule.hasError ? "Yes" : "No",
  );

  appendMetadata(
    metadata,
    "Read only",
    rule.isReadOnly ? "Yes" : "No",
  );

  article.append(
    heading,
    metadata,
    createJsonSection(
      "Conditions",
      rule.conditions ?? {},
    ),
    createJsonSection(
      "Actions",
      rule.actions ?? {},
    ),
  );

  return article;
}

function appendMetadata(
  container,
  label,
  value,
) {
  const wrapper =
    document.createElement("div");

  const term =
    document.createElement("dt");

  const description =
    document.createElement("dd");

  term.textContent = label;
  description.textContent = String(value);

  wrapper.append(term, description);
  container.append(wrapper);
}

function createJsonSection(title, value) {
  const section =
    document.createElement("section");

  const heading =
    document.createElement("h3");

  const pre =
    document.createElement("pre");

  heading.textContent = title;
  pre.textContent =
    JSON.stringify(value, null, 2);

  section.append(heading, pre);

  return section;
}