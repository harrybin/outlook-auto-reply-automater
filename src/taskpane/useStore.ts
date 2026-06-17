/**
 * useStore.ts – Zustand global store for the add-in.
 */

import { create } from "zustand";
import { nanoid } from "./utils/nanoid";
import type {
  AppSettings,
  AppointmentBusyStatus,
  AppointmentMatchField,
  AppointmentMatchOperator,
  AutoReplyMessage,
  AutomationProfile,
  LocationRule,
  LocationTrigger,
  LocationSettings,
  TeamsPresenceStatus,
} from "./types";
import { loadSettings, saveSettings } from "./services/storageService";

interface AppStore extends AppSettings {
  // ── Message CRUD ───────────────────────────────────────────────────────────
  addMessage: (
    msg: Omit<AutoReplyMessage, "id" | "createdAt" | "updatedAt">,
  ) => void;
  updateMessage: (
    id: string,
    changes: Partial<Omit<AutoReplyMessage, "id">>,
  ) => void;
  deleteMessage: (id: string) => void;

  // ── Profile CRUD ──────────────────────────────────────────────────────────
  addProfile: (
    profile: Omit<AutomationProfile, "id" | "createdAt" | "updatedAt">,
  ) => void;
  updateProfile: (
    id: string,
    changes: Partial<Omit<AutomationProfile, "id">>,
  ) => void;
  deleteProfile: (id: string) => void;

  // ── Location settings ─────────────────────────────────────────────────────
  updateLocationSettings: (settings: Partial<LocationSettings>) => void;

  // ── Active auto-reply ─────────────────────────────────────────────────────
  setActiveAutoReplyId: (id: string | null) => void;

  // ── Persistence ───────────────────────────────────────────────────────────
  loadFromStorage: () => void;
  exportSettings: () => string;
  importSettings: (
    rawSettings: string,
  ) => { success: true } | { success: false; error: string };
}

const BUSY_STATUSES: AppointmentBusyStatus[] = [
  "free",
  "tentative",
  "busy",
  "outOfOffice",
  "workingElsewhere",
];

const MATCH_FIELDS: AppointmentMatchField[] = [
  "title",
  "body",
  "location",
  "organizer",
  "category",
];

const MATCH_OPERATORS: AppointmentMatchOperator[] = [
  "contains",
  "startsWith",
  "endsWith",
  "equals",
  "regex",
];

const TEAMS_STATUSES: TeamsPresenceStatus[] = [
  "Available",
  "Busy",
  "DoNotDisturb",
  "BeRightBack",
  "Away",
  "Offline",
];

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

function isString(value: unknown): value is string {
  return typeof value === "string";
}

function isBoolean(value: unknown): value is boolean {
  return typeof value === "boolean";
}

function asNonNegativeNumber(value: unknown, fallback: number): number {
  return typeof value === "number" && Number.isFinite(value) && value >= 0
    ? value
    : fallback;
}

function asString(value: unknown, fallback = ""): string {
  return isString(value) ? value : fallback;
}

function asBoolean(value: unknown, fallback: boolean): boolean {
  return isBoolean(value) ? value : fallback;
}

function asOneOf<T extends string>(
  value: unknown,
  allowed: readonly T[],
  fallback: T,
): T {
  return isString(value) && allowed.includes(value as T)
    ? (value as T)
    : fallback;
}

function asOptionalNumber(value: unknown): number | undefined {
  return typeof value === "number" && Number.isFinite(value)
    ? value
    : undefined;
}

function normalizeMessage(value: unknown): AutoReplyMessage | null {
  if (!isRecord(value)) return null;
  if (
    !isString(value.id) ||
    !isString(value.name) ||
    !isString(value.subject) ||
    !isString(value.body)
  ) {
    return null;
  }

  return {
    id: value.id,
    name: value.name,
    subject: value.subject,
    body: value.body,
    isHtml: asBoolean(value.isHtml, true),
    createdAt: asString(value.createdAt, now()),
    updatedAt: asString(value.updatedAt, now()),
  };
}

function normalizeLocationTrigger(value: unknown): LocationTrigger | null {
  if (!isRecord(value) || !isString(value.type)) return null;

  if (value.type === "geofence") {
    if (
      !isString(value.label) ||
      typeof value.latitude !== "number" ||
      typeof value.longitude !== "number" ||
      typeof value.radiusMeters !== "number"
    ) {
      return null;
    }

    return {
      type: "geofence",
      label: value.label,
      latitude: value.latitude,
      longitude: value.longitude,
      radiusMeters: value.radiusMeters,
    };
  }

  if (value.type === "wifi") {
    if (!isString(value.ssid)) return null;
    return { type: "wifi", ssid: value.ssid };
  }

  if (value.type === "noInternet") {
    return { type: "noInternet" };
  }

  return null;
}

function normalizeLocationRule(value: unknown): LocationRule | null {
  if (!isRecord(value)) return null;
  const trigger = normalizeLocationTrigger(value.trigger);
  if (!trigger || !isString(value.id) || !isString(value.autoReplyMessageId)) {
    return null;
  }

  const teamsStatus = isString(value.teamsStatus)
    ? asOneOf(value.teamsStatus, TEAMS_STATUSES, "Away")
    : undefined;

  return {
    id: value.id,
    enabled: asBoolean(value.enabled, true),
    trigger,
    autoReplyMessageId: value.autoReplyMessageId,
    teamsStatus,
    teamsStatusMessage: isString(value.teamsStatusMessage)
      ? value.teamsStatusMessage
      : undefined,
  };
}

function normalizeLocationSettings(value: unknown): LocationSettings {
  if (!isRecord(value)) {
    return {
      enabled: false,
      pollIntervalSeconds: 60,
      rules: [],
    };
  }

  const rulesSource = Array.isArray(value.rules) ? value.rules : [];
  return {
    enabled: asBoolean(value.enabled, false),
    pollIntervalSeconds: asNonNegativeNumber(value.pollIntervalSeconds, 60),
    rules: rulesSource
      .map((rule) => normalizeLocationRule(rule))
      .filter((rule): rule is LocationRule => rule !== null),
  };
}

function normalizeProfile(value: unknown): AutomationProfile | null {
  if (
    !isRecord(value) ||
    !isString(value.id) ||
    !isString(value.name) ||
    !isString(value.autoReplyMessageId)
  ) {
    return null;
  }

  const keywordRulesSource =
    isRecord(value.matchRules) && Array.isArray(value.matchRules.keywordRules)
      ? value.matchRules.keywordRules
      : [];

  const keywordRules = keywordRulesSource
    .map((rule) => {
      if (!isRecord(rule) || !isString(rule.id) || !isString(rule.value))
        return null;
      return {
        id: rule.id,
        field: asOneOf(rule.field, MATCH_FIELDS, "title"),
        operator: asOneOf(rule.operator, MATCH_OPERATORS, "contains"),
        value: rule.value,
        caseSensitive: asBoolean(rule.caseSensitive, false),
      };
    })
    .filter(
      (rule): rule is AutomationProfile["matchRules"]["keywordRules"][number] =>
        rule !== null,
    );

  const durationRule =
    isRecord(value.matchRules) && isRecord(value.matchRules.durationRule)
      ? value.matchRules.durationRule
      : {};
  const busyStatusRule =
    isRecord(value.matchRules) && isRecord(value.matchRules.busyStatusRule)
      ? value.matchRules.busyStatusRule
      : {};

  const statusesSource = Array.isArray(busyStatusRule.statuses)
    ? busyStatusRule.statuses
    : [];
  const statuses = statusesSource
    .map((status) =>
      isString(status) ? asOneOf(status, BUSY_STATUSES, "busy") : null,
    )
    .filter((status): status is AppointmentBusyStatus => status !== null);

  const timingSettings = isRecord(value.timingSettings)
    ? value.timingSettings
    : {};
  const teamsStatusSettings = isRecord(value.teamsStatusSettings)
    ? value.teamsStatusSettings
    : {};

  return {
    id: value.id,
    name: value.name,
    enabled: asBoolean(value.enabled, true),
    autoReplyMessageId: value.autoReplyMessageId,
    priority: asNonNegativeNumber(value.priority, 50),
    createdAt: asString(value.createdAt, now()),
    updatedAt: asString(value.updatedAt, now()),
    matchRules: {
      keywordRules,
      durationRule: {
        enabled: asBoolean(durationRule.enabled, false),
        minMinutes: asOptionalNumber(durationRule.minMinutes),
        maxMinutes: asOptionalNumber(durationRule.maxMinutes),
      },
      busyStatusRule: {
        enabled: asBoolean(busyStatusRule.enabled, false),
        statuses,
      },
      combinator: asOneOf(
        value.matchRules && isRecord(value.matchRules)
          ? value.matchRules.combinator
          : undefined,
        ["AND", "OR"],
        "AND",
      ),
    },
    timingSettings: {
      enableBefore: asBoolean(timingSettings.enableBefore, false),
      hoursBeforeAppointment: asNonNegativeNumber(
        timingSettings.hoursBeforeAppointment,
        0,
      ),
      enableAfter: asBoolean(timingSettings.enableAfter, false),
      hoursAfterAppointment: asNonNegativeNumber(
        timingSettings.hoursAfterAppointment,
        0,
      ),
    },
    teamsStatusSettings: {
      enabled: asBoolean(teamsStatusSettings.enabled, false),
      statusWhenActive: asOneOf(
        teamsStatusSettings.statusWhenActive,
        TEAMS_STATUSES,
        "Away",
      ),
      statusMessageWhenActive: asString(
        teamsStatusSettings.statusMessageWhenActive,
      ),
      restoreOnEnd: asBoolean(teamsStatusSettings.restoreOnEnd, true),
    },
  };
}

function normalizeImportedSettings(value: unknown): AppSettings | null {
  if (!isRecord(value)) return null;

  const messagesSource = Array.isArray(value.autoReplyMessages)
    ? value.autoReplyMessages
    : [];
  const profilesSource = Array.isArray(value.automationProfiles)
    ? value.automationProfiles
    : [];

  const autoReplyMessages = messagesSource
    .map((message) => normalizeMessage(message))
    .filter((message): message is AutoReplyMessage => message !== null);

  const automationProfiles = profilesSource
    .map((profile) => normalizeProfile(profile))
    .filter((profile): profile is AutomationProfile => profile !== null);

  const activeAutoReplyId =
    isString(value.activeAutoReplyId) || value.activeAutoReplyId === null
      ? value.activeAutoReplyId
      : null;

  return {
    autoReplyMessages,
    automationProfiles,
    locationSettings: normalizeLocationSettings(value.locationSettings),
    activeAutoReplyId,
  };
}

function now(): string {
  return new Date().toISOString();
}

export const useStore = create<AppStore>((set, get) => ({
  ...loadSettings(),

  // ── Messages ──────────────────────────────────────────────────────────────
  addMessage(msg) {
    const message: AutoReplyMessage = {
      ...msg,
      id: nanoid(),
      createdAt: now(),
      updatedAt: now(),
    };
    set((s) => {
      const updated = {
        ...s,
        autoReplyMessages: [...s.autoReplyMessages, message],
      };
      saveSettings(updated);
      return updated;
    });
  },

  updateMessage(id, changes) {
    set((s) => {
      const msgs = s.autoReplyMessages.map((m) =>
        m.id === id ? { ...m, ...changes, updatedAt: now() } : m,
      );
      const updated = { ...s, autoReplyMessages: msgs };
      saveSettings(updated);
      return updated;
    });
  },

  deleteMessage(id) {
    set((s) => {
      const updated = {
        ...s,
        autoReplyMessages: s.autoReplyMessages.filter((m) => m.id !== id),
      };
      saveSettings(updated);
      return updated;
    });
  },

  // ── Profiles ──────────────────────────────────────────────────────────────
  addProfile(profile) {
    const newProfile: AutomationProfile = {
      ...profile,
      id: nanoid(),
      createdAt: now(),
      updatedAt: now(),
    };
    set((s) => {
      const updated = {
        ...s,
        automationProfiles: [...s.automationProfiles, newProfile],
      };
      saveSettings(updated);
      return updated;
    });
  },

  updateProfile(id, changes) {
    set((s) => {
      const profiles = s.automationProfiles.map((p) =>
        p.id === id ? { ...p, ...changes, updatedAt: now() } : p,
      );
      const updated = { ...s, automationProfiles: profiles };
      saveSettings(updated);
      return updated;
    });
  },

  deleteProfile(id) {
    set((s) => {
      const updated = {
        ...s,
        automationProfiles: s.automationProfiles.filter((p) => p.id !== id),
      };
      saveSettings(updated);
      return updated;
    });
  },

  // ── Location ──────────────────────────────────────────────────────────────
  updateLocationSettings(settings) {
    set((s) => {
      const updated = {
        ...s,
        locationSettings: { ...s.locationSettings, ...settings },
      };
      saveSettings(updated);
      return updated;
    });
  },

  // ── Active auto-reply ─────────────────────────────────────────────────────
  setActiveAutoReplyId(id) {
    set((s) => {
      const updated = { ...s, activeAutoReplyId: id };
      saveSettings(updated);
      return updated;
    });
  },

  // ── Storage ───────────────────────────────────────────────────────────────
  loadFromStorage() {
    const stored = loadSettings();
    set((s) => ({ ...s, ...stored }));
  },

  exportSettings() {
    const current = get();
    const settings: AppSettings = {
      autoReplyMessages: current.autoReplyMessages,
      automationProfiles: current.automationProfiles,
      locationSettings: current.locationSettings,
      activeAutoReplyId: current.activeAutoReplyId,
    };
    return JSON.stringify(settings, null, 2);
  },

  importSettings(rawSettings) {
    let parsed: unknown;

    try {
      parsed = JSON.parse(rawSettings);
    } catch {
      return { success: false, error: "Invalid JSON file." };
    }

    const normalized = normalizeImportedSettings(parsed);
    if (!normalized) {
      return { success: false, error: "Invalid settings format." };
    }

    set((s) => {
      const updated = { ...s, ...normalized };
      saveSettings(updated);
      return updated;
    });

    return { success: true };
  },

  // Expose a getter to satisfy TypeScript (used internally)
  get() {
    return get();
  },
}));
