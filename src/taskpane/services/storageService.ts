import type {
  AppSettings,
  AutoReplyMessage,
  AutomationProfile,
  LocationSettings,
} from "../types";

const STORAGE_KEY = "outlookAutoReplyAutomater_settings";

const PRESET_CREATED_AT = "2026-09-22T00:00:00.000Z";

const DEFAULT_MESSAGES: AutoReplyMessage[] = [
  {
    id: "preset-message-vacation",
    name: "Vacation – Out of Office 🌴",
    subject: "Out of Office – Vacation 🌴",
    body: "<p>Thank you for your email. I am currently unavailable due to <strong>{{appointment.title}}</strong> and will return on <strong>{{appointment.nextWorkingDayAfterEnd}}</strong>.</p><p>For urgent matters, please contact my team.</p><p>Best regards</p>",
    isHtml: true,
    createdAt: PRESET_CREATED_AT,
    updatedAt: PRESET_CREATED_AT,
  },
  {
    id: "preset-message-meeting",
    name: "In a Meeting 🗓️",
    subject: "Currently in a Meeting 🗓️",
    body: "I am currently in {{appointment.title}} until {{appointment.end}} and will respond to your message as soon as possible. If urgent, please call me directly.",
    isHtml: false,
    createdAt: PRESET_CREATED_AT,
    updatedAt: PRESET_CREATED_AT,
  },
  {
    id: "preset-message-public-holiday",
    name: "Public Holiday 🎉",
    subject: "Public Holiday – Office Closed 🎉",
    body: "<p>Our office is closed due to {{appointment.title}}. We will be back in the office on {{appointment.nextWorkingDayAfterEnd}}.</p><p>Thank you for your understanding.</p>",
    isHtml: true,
    createdAt: PRESET_CREATED_AT,
    updatedAt: PRESET_CREATED_AT,
  },
  {
    id: "preset-message-training",
    name: "Training – Teams only 🎓",
    subject: "In Training 🎓",
    body: "I am in {{rule.match}} until {{appointment.end}}. I will respond as soon as possible.",
    isHtml: false,
    createdAt: PRESET_CREATED_AT,
    updatedAt: PRESET_CREATED_AT,
  },
  {
    id: "preset-message-traveling",
    name: "Traveling ✈️",
    subject: "Currently Traveling ✈️",
    body: "<p>Thank you for your email. ✈️</p><p>I am currently traveling for {{appointment.title}} and have limited availability. I will reply from {{appointment.nextWorkingDayAfterEnd}}.</p>",
    isHtml: true,
    createdAt: PRESET_CREATED_AT,
    updatedAt: PRESET_CREATED_AT,
  },
];

const DEFAULT_PROFILES: AutomationProfile[] = [
  {
    id: "preset-profile-vacation",
    name: "Vacation Auto-Reply 🌴",
    enabled: false,
    autoReplyAudience: "both",
    autoReplyMessageId: "preset-message-vacation",
    matchRules: {
      keywordRules: [
        {
          id: "preset-rule-vacation",
          field: "title",
          operator: "contains",
          value: "Vacation",
          caseSensitive: false,
        },
      ],
      durationRule: { enabled: true, minHours: 8 },
      busyStatusRule: { enabled: true, statuses: ["outOfOffice"] },
      combinator: "AND",
    },
    timingSettings: {
      enableBefore: false,
      hoursBeforeAppointment: 0,
      enableAfter: false,
      hoursAfterAppointment: 0,
    },
    teamsStatusSettings: {
      enabled: true,
      statusWhenActive: "Away",
      statusMessageWhenActive: "On vacation 🌴",
      restoreOnEnd: true,
    },
    priority: 10,
    createdAt: PRESET_CREATED_AT,
    updatedAt: PRESET_CREATED_AT,
  },
  {
    id: "preset-profile-meeting",
    name: "Meeting Reply 🗓️",
    enabled: false,
    autoReplyAudience: "internal",
    autoReplyMessageId: "preset-message-meeting",
    matchRules: {
      keywordRules: [
        {
          id: "preset-rule-meeting",
          field: "title",
          operator: "contains",
          value: "Meeting",
          caseSensitive: false,
        },
      ],
      durationRule: { enabled: true, maxHours: 479 / 60 },
      busyStatusRule: { enabled: true, statuses: ["busy"] },
      combinator: "AND",
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
      statusMessageWhenActive: "In a meeting 🗓️",
      restoreOnEnd: true,
    },
    priority: 10,
    createdAt: PRESET_CREATED_AT,
    updatedAt: PRESET_CREATED_AT,
  },
  {
    id: "preset-profile-public-holiday",
    name: "Public Holiday 🎉",
    enabled: false,
    autoReplyAudience: "both",
    autoReplyMessageId: "preset-message-public-holiday",
    matchRules: {
      keywordRules: [
        {
          id: "preset-rule-public-holiday",
          field: "title",
          operator: "contains",
          value: "Holiday",
          caseSensitive: false,
        },
      ],
      durationRule: { enabled: true, minHours: 8 },
      busyStatusRule: { enabled: false, statuses: [] },
      combinator: "AND",
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
      statusMessageWhenActive: "Public Holiday 🎉",
      restoreOnEnd: true,
    },
    priority: 10,
    createdAt: PRESET_CREATED_AT,
    updatedAt: PRESET_CREATED_AT,
  },
  {
    id: "preset-profile-training",
    name: "Training – Teams only 🎓",
    enabled: false,
    enableAutoReply: false,
    autoReplyAudience: "internal",
    autoReplyMessageId: "preset-message-training",
    matchRules: {
      keywordRules: [
        {
          id: "preset-rule-training",
          field: "title",
          operator: "regex",
          value: "/\\b(Training)\\b/i",
          caseSensitive: false,
        },
      ],
      durationRule: { enabled: false },
      busyStatusRule: { enabled: true, statuses: ["busy"] },
      combinator: "AND",
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
      statusMessageWhenActive: "In Training 🎓",
      restoreOnEnd: true,
    },
    priority: 40,
    createdAt: PRESET_CREATED_AT,
    updatedAt: PRESET_CREATED_AT,
  },
  {
    id: "preset-profile-traveling",
    name: "Traveling – journey, drive or flight ✈️",
    enabled: false,
    autoReplyAudience: "internal",
    autoReplyMessageId: "preset-message-traveling",
    matchRules: {
      keywordRules: [
        {
          id: "preset-rule-traveling",
          field: "title",
          operator: "regex",
          value:
            "/(?:\\btravel(?:ing)?\\b|\\bdrive\\b|\\bdriving\\b|\\bcar\\b|\\btrain\\b|\\bflight\\b|\\bflying\\b|\\bjourney\\b|\\bunterwegs\\b|\\bauto\\b|\\bzug\\b|\\bflugzeug\\b|\\bfahrt\\b|\\bflieg(?:e|en)\\b|\\bnach\\b|\\bzurück\\b|\\brückfahrt\\b)/i",
          caseSensitive: false,
        },
      ],
      durationRule: { enabled: true, maxHours: 479 / 60 },
      busyStatusRule: { enabled: true, statuses: ["busy"] },
      combinator: "AND",
    },
    timingSettings: {
      enableBefore: false,
      hoursBeforeAppointment: 0,
      enableAfter: false,
      hoursAfterAppointment: 0,
    },
    teamsStatusSettings: {
      enabled: true,
      statusWhenActive: "Away",
      statusMessageWhenActive: "Traveling ✈️",
      restoreOnEnd: true,
    },
    priority: 30,
    createdAt: PRESET_CREATED_AT,
    updatedAt: PRESET_CREATED_AT,
  },
];

const DEFAULT_SETTINGS: AppSettings = {
  autoReplyMessages: [],
  automationProfiles: [],
  locationSettings: {
    enabled: false,
    pollIntervalSeconds: 60,
    rules: [],
  },
  activeAutoReplyId: null,
  hasHandledDefaultRulesPrompt: false,
};

export function getDefaultRulesAndMessages(): Pick<
  AppSettings,
  "autoReplyMessages" | "automationProfiles"
> {
  return {
    autoReplyMessages: DEFAULT_MESSAGES.map((message) => ({ ...message })),
    automationProfiles: DEFAULT_PROFILES.map((profile) => ({
      ...profile,
      matchRules: {
        ...profile.matchRules,
        keywordRules: profile.matchRules.keywordRules.map((rule) => ({
          ...rule,
        })),
        durationRule: { ...profile.matchRules.durationRule },
        busyStatusRule: {
          ...profile.matchRules.busyStatusRule,
          statuses: [...profile.matchRules.busyStatusRule.statuses],
        },
      },
      timingSettings: { ...profile.timingSettings },
      teamsStatusSettings: { ...profile.teamsStatusSettings },
    })),
  };
}

export function loadSettings(): AppSettings {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return { ...DEFAULT_SETTINGS };
    const stored = JSON.parse(raw) as Partial<AppSettings>;
    return {
      ...DEFAULT_SETTINGS,
      ...stored,
      // Existing installations predate the prompt and must not see it later
      // merely because all profiles were removed.
      hasHandledDefaultRulesPrompt:
        typeof stored.hasHandledDefaultRulesPrompt === "boolean"
          ? stored.hasHandledDefaultRulesPrompt
          : true,
    };
  } catch {
    return { ...DEFAULT_SETTINGS };
  }
}

export function saveSettings(settings: AppSettings): void {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(settings));
}

export function saveMessages(messages: AutoReplyMessage[]): void {
  const settings = loadSettings();
  saveSettings({ ...settings, autoReplyMessages: messages });
}

export function saveProfiles(profiles: AutomationProfile[]): void {
  const settings = loadSettings();
  saveSettings({ ...settings, automationProfiles: profiles });
}

export function saveLocationSettings(locationSettings: LocationSettings): void {
  const settings = loadSettings();
  saveSettings({ ...settings, locationSettings });
}

export function clearSettings(): void {
  localStorage.removeItem(STORAGE_KEY);
}
