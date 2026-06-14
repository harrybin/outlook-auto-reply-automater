/**
 * authService.ts
 *
 * MSAL-based authentication for Microsoft Graph API access.
 * Uses the MSAL Browser library for SPA / add-in flows.
 */

import { PublicClientApplication, type AccountInfo, type AuthenticationResult } from "@azure/msal-browser";
import { Client } from "@microsoft/microsoft-graph-client";

const SCOPES = [
  "User.Read",
  "MailboxSettings.ReadWrite",
  "Calendars.Read",
  "Presence.ReadWrite",
];

// Replace with your Azure AD app registration client ID
const CLIENT_ID = import.meta.env.VITE_AAD_CLIENT_ID as string ?? "";
const TENANT_ID = import.meta.env.VITE_AAD_TENANT_ID as string ?? "common";

let msalInstance: PublicClientApplication | null = null;

export function getMsalInstance(): PublicClientApplication {
  if (!msalInstance) {
    msalInstance = new PublicClientApplication({
      auth: {
        clientId: CLIENT_ID,
        authority: `https://login.microsoftonline.com/${TENANT_ID}`,
        redirectUri: window.location.origin,
      },
      cache: {
        cacheLocation: "sessionStorage",
        storeAuthStateInCookie: false,
      },
    });
  }
  return msalInstance;
}

export async function signIn(): Promise<AccountInfo> {
  const msal = getMsalInstance();
  await msal.initialize();
  const result = await msal.loginPopup({ scopes: SCOPES });
  return result.account!;
}

export async function getAccount(): Promise<AccountInfo | null> {
  const msal = getMsalInstance();
  await msal.initialize();
  const accounts = msal.getAllAccounts();
  return accounts.length > 0 ? accounts[0] : null;
}

export async function signOut(): Promise<void> {
  const msal = getMsalInstance();
  await msal.initialize();
  const account = await getAccount();
  if (account) {
    await msal.logoutPopup({ account });
  }
}

async function getAccessToken(): Promise<string> {
  const msal = getMsalInstance();
  await msal.initialize();
  const account = await getAccount();
  if (!account) throw new Error("No signed-in account");

  let result: AuthenticationResult;
  try {
    result = await msal.acquireTokenSilent({ scopes: SCOPES, account });
  } catch {
    result = await msal.acquireTokenPopup({ scopes: SCOPES, account });
  }
  return result.accessToken;
}

/** Returns an authenticated Microsoft Graph client */
export function getGraphClient(): Client {
  return Client.init({
    authProvider: async (done) => {
      try {
        const token = await getAccessToken();
        done(null, token);
      } catch (err) {
        done(err as Error, null);
      }
    },
  });
}
