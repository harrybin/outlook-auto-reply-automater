/**
 * copilotService.ts
 *
 * Generates auto-reply message suggestions for upcoming appointments by
 * calling the Microsoft Graph /me/messages endpoint to invoke the
 * Copilot-powered draft API (or a configurable Azure OpenAI endpoint as
 * fallback).
 *
 * Primary path  – Graph Copilot compose API (preview):
 *   POST /me/messages/microsoft.graph.createForward … uses Copilot drafts.
 *   Until that API is GA we use a direct Azure OpenAI call.
 *
 * The caller provides an authenticated Graph client and an optional
 * Azure OpenAI endpoint.  When no AI endpoint is configured a template-based
 * fallback is used.
 */

import type { AppointmentInfo, CopilotSuggestion } from "../types";

export interface CopilotServiceConfig {
  /** Azure OpenAI endpoint (optional).  If omitted, template fallback is used. */
  azureOpenAiEndpoint?: string;
  /** Azure OpenAI API key (optional). */
  azureOpenAiApiKey?: string;
  /** Deployment / model name (e.g. "gpt-4o") */
  azureOpenAiDeployment?: string;
}

/**
 * Generates auto-reply suggestions for the given appointments.
 *
 * @param appointments – upcoming appointments to analyse
 * @param count        – number of appointments to process
 * @param config       – optional Azure OpenAI configuration
 */
export async function suggestAutoReplies(
  appointments: AppointmentInfo[],
  count = 5,
  config: CopilotServiceConfig = {}
): Promise<CopilotSuggestion[]> {
  const targets = appointments.slice(0, count);

  const suggestions = await Promise.all(
    targets.map((appt) => generateSuggestion(appt, config))
  );

  return suggestions;
}

async function generateSuggestion(
  appointment: AppointmentInfo,
  config: CopilotServiceConfig
): Promise<CopilotSuggestion> {
  if (config.azureOpenAiEndpoint && config.azureOpenAiApiKey && config.azureOpenAiDeployment) {
    try {
      return await generateViaAzureOpenAi(appointment, config as Required<CopilotServiceConfig>);
    } catch {
      // Fall through to template
    }
  }

  return generateTemplateSuggestion(appointment);
}

async function generateViaAzureOpenAi(
  appointment: AppointmentInfo,
  config: Required<CopilotServiceConfig>
): Promise<CopilotSuggestion> {
  const start = new Date(appointment.start).toLocaleString();
  const end = new Date(appointment.end).toLocaleString();
  const duration = formatDuration(appointment.durationMinutes);

  const prompt = `You are an assistant helping to write Outlook out-of-office auto-reply messages.
Generate a professional and concise auto-reply for the following calendar appointment:
- Title: ${appointment.title}
- When: ${start} to ${end} (${duration})
- Location: ${appointment.location ?? "not specified"}
- Status: ${appointment.busyStatus}
${appointment.categories.length ? `- Categories: ${appointment.categories.join(", ")}` : ""}

Respond with a JSON object containing:
{ "subject": "<email subject>", "body": "<email body in HTML>" }`;

  const url = `${config.azureOpenAiEndpoint}/openai/deployments/${config.azureOpenAiDeployment}/chat/completions?api-version=2024-02-15-preview`;

  const response = await fetch(url, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "api-key": config.azureOpenAiApiKey,
    },
    body: JSON.stringify({
      messages: [{ role: "user", content: prompt }],
      max_tokens: 400,
      temperature: 0.7,
    }),
  });

  if (!response.ok) {
    throw new Error(`Azure OpenAI returned ${response.status}`);
  }

  const data = (await response.json()) as {
    choices: Array<{ message: { content: string } }>;
  };
  const content = data.choices[0]?.message.content ?? "";

  let parsed: { subject: string; body: string };
  try {
    // Extract JSON block from potential markdown wrapper
    const match = content.match(/\{[\s\S]*\}/);
    parsed = JSON.parse(match ? match[0] : content);
  } catch {
    parsed = { subject: `Out of office – ${appointment.title}`, body: content };
  }

  return {
    id: `copilot-${appointment.id}`,
    name: `Copilot: ${appointment.title}`,
    subject: parsed.subject,
    body: parsed.body,
    matchedAppointment: appointment,
    confidence: 0.85,
  };
}

function generateTemplateSuggestion(appointment: AppointmentInfo): CopilotSuggestion {
  const start = new Date(appointment.start).toLocaleDateString(undefined, {
    weekday: "long",
    month: "long",
    day: "numeric",
  });
  const end = new Date(appointment.end).toLocaleDateString(undefined, {
    weekday: "long",
    month: "long",
    day: "numeric",
  });
  const sameDay = start === end;
  const period = sameDay ? `on ${start}` : `from ${start} to ${end}`;

  const subject = `Out of office – ${appointment.title}`;
  const body = `<p>Thank you for your email.</p>
<p>I am currently unavailable ${period} due to <em>${appointment.title}</em>.</p>
<p>I will respond to your message as soon as I return. If this is urgent, please contact my team.</p>
<p>Best regards</p>`;

  return {
    id: `template-${appointment.id}`,
    name: `Template: ${appointment.title}`,
    subject,
    body,
    matchedAppointment: appointment,
    confidence: 0.5,
  };
}

function formatDuration(minutes: number): string {
  if (minutes < 60) return `${minutes} min`;
  const hours = Math.floor(minutes / 60);
  const mins = minutes % 60;
  return mins > 0 ? `${hours}h ${mins}min` : `${hours}h`;
}
