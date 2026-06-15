/**
 * teamsService.ts
 *
 * Sets and clears the Microsoft Teams user presence (status) via the
 * Microsoft Graph API.
 *
 * Required scope: Presence.ReadWrite
 *
 * Docs: https://learn.microsoft.com/graph/api/presence-setpresence
 */

import type { TeamsPresenceStatus } from "../types";

type GraphClient = {
  api: (path: string) => {
    post: (body: unknown) => Promise<unknown>;
    get: () => Promise<Record<string, unknown>>;
    delete: () => Promise<void>;
  };
};

const APP_ID = "a1b2c3d4-e5f6-7890-abcd-ef1234567890"; // must match manifest id

/**
 * Sets the signed-in user's Teams presence.
 * @param graphClient – authenticated Microsoft Graph client
 * @param status       – the desired Teams presence status
 * @param message      – optional status message (up to 280 chars)
 * @param expirationDurationMinutes – optional expiry (max 1440 min / 24 h)
 */
export async function setTeamsPresence(
  graphClient: GraphClient,
  status: TeamsPresenceStatus,
  message = "",
  expirationDurationMinutes?: number
): Promise<void> {
  const body: Record<string, unknown> = {
    sessionId: APP_ID,
    availability: status,
    activity: mapAvailabilityToActivity(status),
  };

  if (expirationDurationMinutes !== undefined) {
    body.expirationDuration = `PT${expirationDurationMinutes}M`;
  }

  await graphClient.api("/me/presence/setPresence").post(body);

  if (message) {
    await setTeamsStatusMessage(graphClient, message, expirationDurationMinutes);
  }
}

/**
 * Clears the Teams presence override set by this add-in.
 */
export async function clearTeamsPresence(graphClient: GraphClient): Promise<void> {
  await graphClient.api("/me/presence/clearPresence").post({ sessionId: APP_ID });
}

/**
 * Sets a custom Teams status message.
 * @param message – display text (up to 280 chars)
 * @param expirationDurationMinutes – optional expiry in minutes
 */
export async function setTeamsStatusMessage(
  graphClient: GraphClient,
  message: string,
  expirationDurationMinutes?: number
): Promise<void> {
  const body: Record<string, unknown> = {
    statusMessage: {
      message: {
        content: message.slice(0, 280),
        contentType: "text",
      },
    },
  };

  if (expirationDurationMinutes !== undefined) {
    const expiry = new Date(Date.now() + expirationDurationMinutes * 60_000);
    body["statusMessage"] = {
      ...(body["statusMessage"] as object),
      expiryDateTime: { dateTime: expiry.toISOString(), timeZone: "UTC" },
    };
  }

  await graphClient.api("/me/presence/setStatusMessage").post(body);
}

/**
 * Reads the current Teams presence.
 */
export async function getTeamsPresence(
  graphClient: GraphClient
): Promise<Record<string, unknown>> {
  return graphClient.api("/me/presence").get();
}

function mapAvailabilityToActivity(status: TeamsPresenceStatus): string {
  switch (status) {
    case "Available":
      return "Available";
    case "Busy":
      return "InACall";
    case "DoNotDisturb":
      return "DoNotDisturb";
    case "BeRightBack":
      return "BeRightBack";
    case "Away":
      return "Away";
    case "Offline":
      return "OffWork";
  }
}
