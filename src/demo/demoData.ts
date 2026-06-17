/**
 * demoData.ts – Pre-seeded sample data for the demo / mock environment.
 *
 * Loaded into localStorage on first visit so the add-in UI is immediately
 * populated with realistic content without needing real Microsoft 365 credentials.
 */

import type { AppSettings } from "../taskpane/types";

const now = new Date().toISOString();

export const DEMO_STORAGE_KEY = "outlookAutoReplyAutomater_settings";

export const DEMO_DATA: AppSettings = {
  activeAutoReplyId: null,
  autoReplyMessages: [
    {
      id: "demo-msg-1",
      name: "Vacation – Out of Office",
      subject: "Out of Office – Vacation",
      body: "<p>Thank you for your email. I am currently on vacation and will return on <strong>Monday, 30 June</strong>.</p><p>For urgent matters, please contact my colleague at <a href='mailto:colleague@example.com'>colleague@example.com</a>.</p><p>Best regards</p>",
      isHtml: true,
      createdAt: now,
      updatedAt: now,
    },
    {
      id: "demo-msg-2",
      name: "In a Meeting",
      subject: "Currently in a Meeting",
      body: "I am currently in a meeting and will respond to your message as soon as possible. If urgent, please call me directly.",
      isHtml: false,
      createdAt: now,
      updatedAt: now,
    },
    {
      id: "demo-msg-3",
      name: "Public Holiday",
      subject: "Public Holiday – Office Closed",
      body: "<p>Our office is closed today due to a public holiday. We will be back in the office tomorrow.</p><p>Thank you for your understanding.</p>",
      isHtml: true,
      createdAt: now,
      updatedAt: now,
    },
  ],
  automationProfiles: [
    {
      id: "demo-profile-1",
      name: "Vacation Auto-Reply",
      enabled: true,
      autoReplyMessageId: "demo-msg-1",
      priority: 10,
      matchRules: {
        combinator: "OR",
        keywordRules: [
          {
            id: "demo-kw-1",
            field: "title",
            operator: "contains",
            value: "Vacation",
            caseSensitive: false,
          },
          {
            id: "demo-kw-2",
            field: "title",
            operator: "contains",
            value: "Urlaub",
            caseSensitive: false,
          },
        ],
        durationRule: { enabled: true, minMinutes: 480 },
        busyStatusRule: { enabled: true, statuses: ["outOfOffice"] },
      },
      timingSettings: {
        enableBefore: true,
        hoursBeforeAppointment: 1,
        enableAfter: true,
        hoursAfterAppointment: 0,
      },
      teamsStatusSettings: {
        enabled: true,
        statusWhenActive: "Away",
        statusMessageWhenActive: "On vacation – back Monday",
        restoreOnEnd: true,
      },
      createdAt: now,
      updatedAt: now,
    },
    {
      id: "demo-profile-2",
      name: "Meeting Reply",
      enabled: true,
      autoReplyMessageId: "demo-msg-2",
      priority: 50,
      matchRules: {
        combinator: "AND",
        keywordRules: [
          {
            id: "demo-kw-3",
            field: "title",
            operator: "contains",
            value: "Meeting",
            caseSensitive: false,
          },
        ],
        durationRule: { enabled: false },
        busyStatusRule: { enabled: true, statuses: ["busy"] },
      },
      timingSettings: {
        enableBefore: false,
        hoursBeforeAppointment: 0,
        enableAfter: false,
        hoursAfterAppointment: 0,
      },
      teamsStatusSettings: {
        enabled: true,
        statusWhenActive: "DoNotDisturb",
        statusMessageWhenActive: "In a meeting",
        restoreOnEnd: true,
      },
      createdAt: now,
      updatedAt: now,
    },
    {
      id: "demo-profile-3",
      name: "Public Holiday",
      enabled: false,
      autoReplyMessageId: "demo-msg-3",
      priority: 20,
      matchRules: {
        combinator: "OR",
        keywordRules: [
          {
            id: "demo-kw-4",
            field: "title",
            operator: "contains",
            value: "Holiday",
            caseSensitive: false,
          },
          {
            id: "demo-kw-5",
            field: "title",
            operator: "contains",
            value: "Feiertag",
            caseSensitive: false,
          },
        ],
        durationRule: { enabled: true, minMinutes: 480 },
        busyStatusRule: { enabled: false, statuses: [] },
      },
      timingSettings: {
        enableBefore: false,
        hoursBeforeAppointment: 0,
        enableAfter: false,
        hoursAfterAppointment: 0,
      },
      teamsStatusSettings: {
        enabled: false,
        statusWhenActive: "Away",
        statusMessageWhenActive: "",
        restoreOnEnd: false,
      },
      createdAt: now,
      updatedAt: now,
    },
  ],
  locationSettings: {
    enabled: false,
    pollIntervalSeconds: 60,
    rules: [],
  },
};

/** Seeds localStorage with demo data if it hasn't been seeded yet in this browser. */
export function seedDemoDataIfNeeded(): void {
  try {
    const existing = localStorage.getItem(DEMO_STORAGE_KEY);
    if (!existing) {
      localStorage.setItem(DEMO_STORAGE_KEY, JSON.stringify(DEMO_DATA));
    }
  } catch {
    // localStorage not available – silently skip
  }
}

/** Resets localStorage to the original demo data (used by the Reset button). */
export function resetDemoData(): void {
  try {
    localStorage.setItem(DEMO_STORAGE_KEY, JSON.stringify(DEMO_DATA));
  } catch {
    // silently skip
  }
}
