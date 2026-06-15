/**
 * autoReplyService.ts
 *
 * Sets and clears Outlook auto-reply (out-of-office) messages via
 * the Microsoft Graph API.  The add-in must be consented to the
 * MailboxSettings.ReadWrite delegated scope.
 */

import type { AutoReplyMessage } from "../types";

export interface AutoReplyPayload {
  status: "alwaysEnabled" | "disabled";
  externalAudience?: "none" | "contactsOnly" | "all";
  internalReplyMessage?: { contentType: "html" | "text"; content: string };
  externalReplyMessage?: { contentType: "html" | "text"; content: string };
}

/**
 * Enables the Outlook auto-reply using the Graph API.
 * @param graphClient – an authenticated Microsoft Graph client
 * @param message – the auto-reply message to activate
 */
export async function enableAutoReply(
  graphClient: { api: (path: string) => { patch: (body: unknown) => Promise<unknown> } },
  message: AutoReplyMessage
): Promise<void> {
  const contentType = message.isHtml ? "html" : "text";
  const payload: AutoReplyPayload = {
    status: "alwaysEnabled",
    externalAudience: "all",
    internalReplyMessage: { contentType, content: message.body },
    externalReplyMessage: { contentType, content: message.body },
  };

  await graphClient.api("/me/mailboxSettings/automaticRepliesSetting").patch(payload);
}

/**
 * Disables the Outlook auto-reply using the Graph API.
 */
export async function disableAutoReply(
  graphClient: { api: (path: string) => { patch: (body: unknown) => Promise<unknown> } }
): Promise<void> {
  const payload: AutoReplyPayload = { status: "disabled" };
  await graphClient.api("/me/mailboxSettings/automaticRepliesSetting").patch(payload);
}

/**
 * Retrieves the current auto-reply setting from the Graph API.
 */
export async function getAutoReplySetting(
  graphClient: { api: (path: string) => { get: () => Promise<Record<string, unknown>> } }
): Promise<Record<string, unknown>> {
  return graphClient.api("/me/mailboxSettings/automaticRepliesSetting").get();
}
