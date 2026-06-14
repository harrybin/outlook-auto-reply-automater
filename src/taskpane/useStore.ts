/**
 * useStore.ts – Zustand global store for the add-in.
 */

import { create } from "zustand";
import { nanoid } from "./utils/nanoid";
import type {
  AppSettings,
  AutoReplyMessage,
  AutomationProfile,
  LocationSettings,
} from "./types";
import { loadSettings, saveSettings } from "./services/storageService";

interface AppStore extends AppSettings {
  // ── Message CRUD ───────────────────────────────────────────────────────────
  addMessage: (msg: Omit<AutoReplyMessage, "id" | "createdAt" | "updatedAt">) => void;
  updateMessage: (id: string, changes: Partial<Omit<AutoReplyMessage, "id">>) => void;
  deleteMessage: (id: string) => void;

  // ── Profile CRUD ──────────────────────────────────────────────────────────
  addProfile: (profile: Omit<AutomationProfile, "id" | "createdAt" | "updatedAt">) => void;
  updateProfile: (id: string, changes: Partial<Omit<AutomationProfile, "id">>) => void;
  deleteProfile: (id: string) => void;

  // ── Location settings ─────────────────────────────────────────────────────
  updateLocationSettings: (settings: Partial<LocationSettings>) => void;

  // ── Active auto-reply ─────────────────────────────────────────────────────
  setActiveAutoReplyId: (id: string | null) => void;

  // ── Persistence ───────────────────────────────────────────────────────────
  loadFromStorage: () => void;
}

function now(): string {
  return new Date().toISOString();
}

export const useStore = create<AppStore>((set, get) => ({
  ...loadSettings(),

  // ── Messages ──────────────────────────────────────────────────────────────
  addMessage(msg) {
    const message: AutoReplyMessage = { ...msg, id: nanoid(), createdAt: now(), updatedAt: now() };
    set((s) => {
      const updated = { ...s, autoReplyMessages: [...s.autoReplyMessages, message] };
      saveSettings(updated);
      return updated;
    });
  },

  updateMessage(id, changes) {
    set((s) => {
      const msgs = s.autoReplyMessages.map((m) =>
        m.id === id ? { ...m, ...changes, updatedAt: now() } : m
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
      const updated = { ...s, automationProfiles: [...s.automationProfiles, newProfile] };
      saveSettings(updated);
      return updated;
    });
  },

  updateProfile(id, changes) {
    set((s) => {
      const profiles = s.automationProfiles.map((p) =>
        p.id === id ? { ...p, ...changes, updatedAt: now() } : p
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

  // Expose a getter to satisfy TypeScript (used internally)
  get() {
    return get();
  },
}));
