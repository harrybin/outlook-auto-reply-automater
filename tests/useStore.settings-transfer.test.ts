import { beforeEach, describe, expect, it } from "vitest";
import type { AppSettings } from "@/taskpane/types";
import { useStore } from "@/taskpane/useStore";

const BASE_SETTINGS: AppSettings = {
  autoReplyMessages: [
    {
      id: "msg-1",
      name: "Out of office",
      subject: "Away from desk",
      body: "I will reply when I am back.",
      isHtml: true,
      createdAt: "2026-01-01T00:00:00.000Z",
      updatedAt: "2026-01-01T00:00:00.000Z",
    },
  ],
  automationProfiles: [
    {
      id: "profile-1",
      name: "Default profile",
      enabled: true,
      autoReplyMessageId: "msg-1",
      priority: 10,
      createdAt: "2026-01-01T00:00:00.000Z",
      updatedAt: "2026-01-01T00:00:00.000Z",
      matchRules: {
        keywordRules: [],
        durationRule: { enabled: false },
        busyStatusRule: { enabled: true, statuses: ["busy"] },
        combinator: "AND",
      },
      timingSettings: {
        enableBefore: false,
        hoursBeforeAppointment: 0,
        enableAfter: true,
        hoursAfterAppointment: 1,
      },
      teamsStatusSettings: {
        enabled: true,
        statusWhenActive: "Away",
        statusMessageWhenActive: "In meetings",
        restoreOnEnd: true,
      },
    },
  ],
  locationSettings: {
    enabled: true,
    pollIntervalSeconds: 45,
    rules: [
      {
        id: "location-rule-1",
        enabled: true,
        autoReplyMessageId: "msg-1",
        trigger: {
          type: "noInternet",
        },
      },
    ],
  },
  activeAutoReplyId: "msg-1",
};

const EMPTY_SETTINGS: AppSettings = {
  autoReplyMessages: [],
  automationProfiles: [],
  locationSettings: {
    enabled: false,
    pollIntervalSeconds: 60,
    rules: [],
  },
  activeAutoReplyId: null,
};

describe("useStore settings import/export", () => {
  beforeEach(() => {
    localStorage.clear();
    useStore.setState((state) => ({ ...state, ...EMPTY_SETTINGS }));
  });

  it("exports current settings as JSON", () => {
    useStore.setState((state) => ({ ...state, ...BASE_SETTINGS }));

    const output = useStore.getState().exportSettings();
    const parsed = JSON.parse(output) as AppSettings;

    expect(parsed.autoReplyMessages).toHaveLength(1);
    expect(parsed.automationProfiles).toHaveLength(1);
    expect(parsed.locationSettings.pollIntervalSeconds).toBe(45);
    expect(parsed.activeAutoReplyId).toBe("msg-1");
  });

  it("imports valid settings JSON", () => {
    const json = JSON.stringify(BASE_SETTINGS);

    const result = useStore.getState().importSettings(json);

    expect(result.success).toBe(true);
    const state = useStore.getState();
    expect(state.autoReplyMessages).toHaveLength(1);
    expect(state.automationProfiles).toHaveLength(1);
    expect(state.locationSettings.rules).toHaveLength(1);
    expect(state.activeAutoReplyId).toBe("msg-1");
  });

  it("rejects invalid JSON", () => {
    const result = useStore.getState().importSettings("{not-valid-json}");

    expect(result.success).toBe(false);
    if (!result.success) {
      expect(result.error).toBe("Invalid JSON file.");
    }
  });
});
