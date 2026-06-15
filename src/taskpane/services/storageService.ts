import type { AppSettings, AutoReplyMessage, AutomationProfile, LocationSettings } from "../types";

const STORAGE_KEY = "outlookAutoReplyAutomater_settings";

const DEFAULT_SETTINGS: AppSettings = {
  autoReplyMessages: [],
  automationProfiles: [],
  locationSettings: {
    enabled: false,
    pollIntervalSeconds: 60,
    rules: [],
  },
  activeAutoReplyId: null,
};

export function loadSettings(): AppSettings {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return { ...DEFAULT_SETTINGS };
    return { ...DEFAULT_SETTINGS, ...JSON.parse(raw) };
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
