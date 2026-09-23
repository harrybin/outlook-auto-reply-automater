import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import process from "node:process";
import { PublicClientApplication } from "@azure/msal-node";

const GRAPH_ROOT = "https://graph.microsoft.com/v1.0";
const DEFAULT_INTERVAL_SECONDS = 60;
const DEFAULT_SETTINGS_PATH = path.join(
  process.env.APPDATA ?? os.homedir(),
  "outlook-auto-reply-automater",
  "settings.json",
);
const TOKEN_CACHE_PATH = path.join(
  process.env.LOCALAPPDATA ?? os.homedir(),
  "outlook-auto-reply-automater",
  "token-cache.json",
);

function argument(name, fallback) {
  const index = process.argv.indexOf(name);
  return index >= 0 ? process.argv[index + 1] ?? fallback : fallback;
}

const settingsPath = path.resolve(argument("--settings", DEFAULT_SETTINGS_PATH));
const intervalSeconds = Math.max(
  Number(argument("--interval", DEFAULT_INTERVAL_SECONDS)) || DEFAULT_INTERVAL_SECONDS,
  30,
);
const clientId = process.env.VITE_AAD_CLIENT_ID ?? process.env.AAD_CLIENT_ID;
const tenantId = process.env.VITE_AAD_TENANT_ID ?? process.env.AAD_TENANT_ID ?? "common";

if (!clientId) {
  throw new Error("Set VITE_AAD_CLIENT_ID or AAD_CLIENT_ID before starting the runner.");
}

await fs.mkdir(path.dirname(TOKEN_CACHE_PATH), { recursive: true });

const cachePlugin = {
  async beforeCacheAccess(context) {
    try {
      context.tokenCache.deserialize(await fs.readFile(TOKEN_CACHE_PATH, "utf8"));
    } catch (error) {
      if (error.code !== "ENOENT") throw error;
    }
  },
  async afterCacheAccess(context) {
    if (context.cacheHasChanged) {
      await fs.writeFile(TOKEN_CACHE_PATH, context.tokenCache.serialize(), "utf8");
    }
  },
};

const msal = new PublicClientApplication({
  auth: {
    clientId,
    authority: `https://login.microsoftonline.com/${tenantId}`,
  },
  cache: { cachePlugin },
});

async function accessToken() {
  const accounts = await msal.getTokenCache().getAllAccounts();
  if (accounts[0]) {
    try {
      const result = await msal.acquireTokenSilent({
        account: accounts[0],
        scopes: ["User.Read", "Calendars.Read", "MailboxSettings.ReadWrite", "Presence.ReadWrite"],
      });
      return result.accessToken;
    } catch {
      // Device code below obtains a fresh token when silent acquisition fails.
    }
  }

  const result = await msal.acquireTokenByDeviceCode({
    scopes: ["User.Read", "Calendars.Read", "MailboxSettings.ReadWrite", "Presence.ReadWrite"],
    deviceCodeCallback: (response) => console.log(response.message),
  });
  return result.accessToken;
}

async function graph(pathname, token, options = {}) {
  const response = await fetch(`${GRAPH_ROOT}${pathname}`, {
    ...options,
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
      ...(options.headers ?? {}),
    },
  });
  if (!response.ok) {
    throw new Error(`Microsoft Graph ${response.status}: ${await response.text()}`);
  }
  return response.status === 204 ? null : response.json();
}

async function loadSettings() {
  return JSON.parse(await fs.readFile(settingsPath, "utf8"));
}

function fieldValue(event, field) {
  if (field === "title") return event.subject ?? "";
  if (field === "location") return event.location?.displayName ?? "";
  if (field === "category") return (event.categories ?? []).join(" ");
  if (field === "body") return event.bodyPreview ?? "";
  if (field === "organizer") return event.organizer?.emailAddress?.name ?? "";
  return "";
}

function keywordMatches(event, rule) {
  const actual = fieldValue(event, rule.field);
  const expected = rule.caseSensitive ? rule.value : rule.value.toLowerCase();
  const value = rule.caseSensitive ? actual : actual.toLowerCase();
  if (rule.operator === "contains") return value.includes(expected);
  if (rule.operator === "startsWith") return value.startsWith(expected);
  if (rule.operator === "endsWith") return value.endsWith(expected);
  if (rule.operator === "equals") return value === expected;
  try {
    return new RegExp(rule.value, rule.caseSensitive ? "" : "i").test(actual);
  } catch {
    return false;
  }
}

function profileMatches(event, profile) {
  const rules = profile.matchRules ?? {};
  const results = [];
  if ((rules.keywordRules ?? []).length) {
    const matches = rules.keywordRules.map((rule) => keywordMatches(event, rule));
    results.push(rules.combinator === "OR" ? matches.some(Boolean) : matches.every(Boolean));
  }
  const duration = (new Date(event.end.dateTime) - new Date(event.start.dateTime)) / 3_600_000;
  if (rules.durationRule?.enabled) {
    results.push(
      (rules.durationRule.minHours === undefined || duration >= rules.durationRule.minHours) &&
      (rules.durationRule.maxHours === undefined || duration <= rules.durationRule.maxHours),
    );
  }
  if (rules.busyStatusRule?.enabled && rules.busyStatusRule.statuses?.length) {
    const busyStatus = {
      free: "free",
      tentative: "tentative",
      busy: "busy",
      oof: "outOfOffice",
      workingElsewhere: "workingElsewhere",
    }[event.showAs] ?? "busy";
    results.push(rules.busyStatusRule.statuses.includes(busyStatus));
  }
  return results.length === 0 || (rules.combinator === "OR" ? results.some(Boolean) : results.every(Boolean));
}

function inTimingWindow(event, profile, now) {
  const timing = profile.timingSettings ?? {};
  const before = timing.enableBefore ? (timing.hoursBeforeAppointment ?? 0) * 3_600_000 : 0;
  const after = timing.enableAfter ? (timing.hoursAfterAppointment ?? 0) * 3_600_000 : 0;
  return now >= new Date(event.start.dateTime).getTime() - before &&
    now <= new Date(event.end.dateTime).getTime() + after;
}

function render(template, event) {
  if (!event) return template;
  return template
    .replaceAll("{{appointment.title}}", event.subject ?? "")
    .replaceAll("{{appointment.start}}", new Date(event.start.dateTime).toLocaleDateString())
    .replaceAll("{{appointment.end}}", new Date(event.end.dateTime).toLocaleDateString())
    .replaceAll("{{appointment.location}}", event.location?.displayName ?? "");
}

async function desiredState(settings, token) {
  const now = new Date();
  const lower = new Date(now.getTime() - 86_400_000).toISOString();
  const upper = new Date(now.getTime() + 86_400_000).toISOString();
  const events = (await graph(
    `/me/calendarView?startDateTime=${encodeURIComponent(lower)}&endDateTime=${encodeURIComponent(upper)}&$top=50`,
    token,
  )).value ?? [];
  const profiles = settings.automationProfiles ?? [];
  const messages = settings.autoReplyMessages ?? [];
  for (const event of events.sort((a, b) => new Date(a.start.dateTime) - new Date(b.start.dateTime))) {
    for (const profile of profiles) {
      if (!profile.enabled || !profileMatches(event, profile) || !inTimingWindow(event, profile, now.getTime())) continue;
      const message = messages.find((item) => item.id === profile.autoReplyMessageId);
      if (message) return { id: `appointment:${event.id}:${profile.id}:${message.updatedAt}`, event, profile, message };
    }
  }
  return null;
}

async function apply(state, token) {
  const audience = state.profile.autoReplyAudience ?? "both";
  if (state.profile.enableAutoReply !== false) {
    const contentType = state.message.isHtml ? "html" : "text";
    const content = render(state.message.body, state.event);
    await graph("/me/mailboxSettings/automaticRepliesSetting", token, {
      method: "PATCH",
      body: JSON.stringify({
        status: "alwaysEnabled",
        externalAudience: audience === "internal" ? "none" : "all",
        ...(audience !== "external" ? { internalReplyMessage: { contentType, content } } : {}),
        ...(audience !== "internal" ? { externalReplyMessage: { contentType, content } } : {}),
      }),
    });
  }
  const teams = state.profile.teamsStatusSettings;
  if (teams?.enabled) {
    await graph("/me/presence/setPresence", token, {
      method: "POST",
      body: JSON.stringify({
        sessionId: "outlook-auto-reply-automater-local-runner",
        availability: teams.statusWhenActive,
        activity: teams.statusWhenActive,
        expirationDuration: "PT24H",
      }),
    });
  }
}

let appliedId = null;
let appliedState = null;
async function evaluate() {
  const settings = await loadSettings();
  const token = await accessToken();
  const desired = await desiredState(settings, token);
  if (desired?.id === appliedId) return;
  if (appliedState?.profile.enableAutoReply !== false) {
    await graph("/me/mailboxSettings/automaticRepliesSetting", token, {
      method: "PATCH",
      body: JSON.stringify({ status: "disabled" }),
    });
  }
  if (appliedState?.profile.teamsStatusSettings?.enabled &&
      appliedState.profile.teamsStatusSettings.restoreOnEnd) {
    await graph("/me/presence/clearPresence", token, {
      method: "POST",
      body: JSON.stringify({ sessionId: "outlook-auto-reply-automater-local-runner" }),
    });
  }
  if (desired) await apply(desired, token);
  appliedId = desired?.id ?? null;
  appliedState = desired;
  console.log(`[${new Date().toISOString()}] ${desired ? `Applied ${desired.id}` : "Cleared managed state"}`);
}

console.log(`Reading settings from ${settingsPath}`);
await evaluate();
setInterval(() => evaluate().catch((error) => console.error(`[${new Date().toISOString()}] ${error.message}`)), intervalSeconds * 1000);
