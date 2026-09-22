/**
 * autoReplyService.ts
 *
 * Sets and clears Outlook auto-reply (out-of-office) messages via
 * the Microsoft Graph API.  The add-in must be consented to the
 * MailboxSettings.ReadWrite delegated scope.
 */

import type {
  AppointmentInfo,
  AutoReplyMessage,
  AutomationProfile,
} from "../types";
import { renderMessageTemplate } from "./messageTemplateService";

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
 * @param appointment – optional calendar appointment used to resolve placeholders
 * @param profile – optional matching profile used to resolve rule placeholders
 */
export async function enableAutoReply(
  graphClient: {
    api: (path: string) => { patch: (body: unknown) => Promise<unknown> };
  },
  message: AutoReplyMessage,
  appointment?: AppointmentInfo,
  profile?: AutomationProfile,
): Promise<void> {
  const contentType = message.isHtml ? "html" : "text";
  const content = renderMessageTemplate(message.body, appointment, profile);
  const payload: AutoReplyPayload = {
    status: "alwaysEnabled",
    externalAudience: "all",
    internalReplyMessage: { contentType, content },
    externalReplyMessage: { contentType, content },
  };

  await graphClient
    .api("/me/mailboxSettings/automaticRepliesSetting")
    .patch(payload);
}

/**
 * Disables the Outlook auto-reply using the Graph API.
 */
export async function disableAutoReply(graphClient: {
  api: (path: string) => { patch: (body: unknown) => Promise<unknown> };
}): Promise<void> {
  const payload: AutoReplyPayload = { status: "disabled" };
  await graphClient
    .api("/me/mailboxSettings/automaticRepliesSetting")
    .patch(payload);
}

/**
 * Retrieves the current auto-reply setting from the Graph API.
 */
export async function getAutoReplySetting(graphClient: {
  api: (path: string) => { get: () => Promise<Record<string, unknown>> };
}): Promise<Record<string, unknown>> {
  return graphClient.api("/me/mailboxSettings/automaticRepliesSetting").get();
}
