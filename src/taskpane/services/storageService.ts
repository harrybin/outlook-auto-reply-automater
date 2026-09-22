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
    id: "preset-message-travel",
    name: "Reise",
    subject: "Auf Reisen",
    body: "<p>Vielen Dank für Ihre Nachricht. ✈️</p><p>Ich bin derzeit auf Reisen und melde mich ab {{appointment.nextWorkingDayAfterEnd}} bei Ihnen zurück.</p>",
    isHtml: true,
    createdAt: PRESET_CREATED_AT,
    updatedAt: PRESET_CREATED_AT,
  },
  {
    id: "preset-message-training",
    name: "Training",
    subject: "Im Training",
    body: "<p>Vielen Dank für Ihre Nachricht. 📚</p><p>Ich bin derzeit bis {{appointment.end}} in einem Training und nur eingeschränkt erreichbar. Ich melde mich danach bei Ihnen zurück.</p>",
    isHtml: true,
    createdAt: PRESET_CREATED_AT,
    updatedAt: PRESET_CREATED_AT,
  },
  {
    id: "preset-message-customer-visit",
    name: "Außer Haus beim Kunden",
    subject: "Beim Kunden vor Ort",
    body: "<p>Vielen Dank für Ihre Nachricht. 🤝</p><p>Ich bin derzeit außer Haus bei einem Kunden in {{appointment.location}} und melde mich ab {{appointment.nextWorkingDayAfterEnd}} bei Ihnen zurück.</p>",
    isHtml: true,
    createdAt: PRESET_CREATED_AT,
    updatedAt: PRESET_CREATED_AT,
  },
];

const DEFAULT_PROFILES: AutomationProfile[] = [
  {
    id: "preset-profile-travel",
    name: "Reise",
    enabled: true,
    autoReplyMessageId: "preset-message-travel",
    matchRules: {
      keywordRules: [
        {
          id: "preset-rule-travel",
          field: "title",
          operator: "contains",
          value: "Reise",
          caseSensitive: false,
        },
      ],
      durationRule: { enabled: false },
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
      statusMessageWhenActive: "",
      restoreOnEnd: true,
    },
    priority: 10,
    createdAt: PRESET_CREATED_AT,
    updatedAt: PRESET_CREATED_AT,
  },
  {
    id: "preset-profile-training",
    name: "Training ab 4 Stunden",
    enabled: true,
    autoReplyMessageId: "preset-message-training",
    matchRules: {
      keywordRules: [
        {
          id: "preset-rule-training",
          field: "title",
          operator: "contains",
          value: "Training",
          caseSensitive: false,
        },
      ],
      durationRule: { enabled: true, minMinutes: 240 },
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
      statusMessageWhenActive: "",
      restoreOnEnd: true,
    },
    priority: 10,
    createdAt: PRESET_CREATED_AT,
    updatedAt: PRESET_CREATED_AT,
  },
  {
    id: "preset-profile-customer-visit",
    name: "Außer Haus beim Kunden",
    enabled: true,
    autoReplyMessageId: "preset-message-customer-visit",
    matchRules: {
      keywordRules: [
        {
          id: "preset-rule-customer-visit",
          field: "title",
          operator: "contains",
          value: "außer Haus beim Kunden",
          caseSensitive: false,
        },
      ],
      durationRule: { enabled: false },
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
      statusMessageWhenActive: "",
      restoreOnEnd: true,
    },
    priority: 10,
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
