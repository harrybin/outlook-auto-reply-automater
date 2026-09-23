import { describe, expect, it } from "vitest";
import type {
  AppointmentInfo,
  AutomationProfile,
  KeywordRule,
} from "@/taskpane/types";
import {
  appointmentMatchesProfile,
  findMatchingAppointments,
} from "@/taskpane/services/calendarService";

function createProfile(rule: KeywordRule): AutomationProfile {
  return {
    id: "profile-1",
    name: "Regex profile",
    enabled: true,
    autoReplyMessageId: "message-1",
    priority: 1,
    createdAt: "2026-01-01T00:00:00.000Z",
    updatedAt: "2026-01-01T00:00:00.000Z",
    matchRules: {
      keywordRules: [rule],
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
      restoreOnEnd: false,
    },
  };
}

const APPOINTMENT: AppointmentInfo = {
  id: "appt-1",
  title: "Vacation 2026",
  start: "2026-01-01T00:00:00.000Z",
  end: "2026-01-01T02:00:00.000Z",
  location: "Berlin",
  busyStatus: "busy",
  isAllDay: false,
  durationMinutes: 120,
  categories: [],
};

describe("appointmentMatchesProfile regex rules", () => {
  it("matches bare regex patterns", () => {
    const profile = createProfile({
      id: "rule-1",
      field: "title",
      operator: "regex",
      value: "^vacation\\s+\\d+$",
      caseSensitive: false,
    });

    expect(appointmentMatchesProfile(APPOINTMENT, profile)).toBe(true);
  });

  it("matches regex literal syntax with flags", () => {
    const profile = createProfile({
      id: "rule-2",
      field: "title",
      operator: "regex",
      value: "/^vacation\\s+\\d+$/i",
      caseSensitive: false,
    });

    expect(appointmentMatchesProfile(APPOINTMENT, profile)).toBe(true);
  });

  it("does not throw on invalid regex patterns", () => {
    const profile = createProfile({
      id: "rule-3",
      field: "title",
      operator: "regex",
      value: "(unclosed",
      caseSensitive: false,
    });

    expect(appointmentMatchesProfile(APPOINTMENT, profile)).toBe(false);
  });
});

describe("appointmentMatchesProfile training preset", () => {
  it("matches Training appointments of four hours or longer, including multi-day appointments", () => {
    const profile = createProfile({
      id: "rule-training",
      field: "title",
      operator: "contains",
      value: "Workshop",
      caseSensitive: false,
    });
    profile.matchRules = {
      ...profile.matchRules,
      keywordRules: [
        {
          id: "training-keyword",
          field: "title",
          operator: "contains",
          value: "Training",
          caseSensitive: false,
        },
      ],
      durationRule: { enabled: true, minHours: 4 },
    };

    expect(
      appointmentMatchesProfile(
        { ...APPOINTMENT, title: "Training", durationMinutes: 240 },
        profile,
      ),
    ).toBe(true);
    expect(
      appointmentMatchesProfile(
        { ...APPOINTMENT, title: "Training", durationMinutes: 2_880 },
        profile,
      ),
    ).toBe(true);
    expect(
      appointmentMatchesProfile(
        { ...APPOINTMENT, title: "Training", durationMinutes: 239 },
        profile,
      ),
    ).toBe(false);
  });

  it("applies minimum and maximum hour limits together", () => {
    const profile = createProfile({
      id: "rule-duration-range",
      field: "title",
      operator: "contains",
      value: "Vacation",
      caseSensitive: false,
    });
    profile.matchRules = {
      ...profile.matchRules,
      durationRule: {
        enabled: true,
        minHours: 2,
        maxHours: 4,
      },
    };

    expect(
      appointmentMatchesProfile(
        { ...APPOINTMENT, durationMinutes: 120 },
        profile,
      ),
    ).toBe(true);
    expect(
      appointmentMatchesProfile(
        { ...APPOINTMENT, durationMinutes: 241 },
        profile,
      ),
    ).toBe(false);
  });

  it("matches a Travel appointment by its title", () => {
    const profile = createProfile({
      id: "rule-travel",
      field: "title",
      operator: "contains",
      value: "unused",
      caseSensitive: false,
    });
    profile.matchRules = {
      ...profile.matchRules,
      keywordRules: [
        {
          id: "travel-keyword",
          field: "title",
          operator: "contains",
          value: "Reise",
          caseSensitive: false,
        },
      ],
    };

    expect(
      appointmentMatchesProfile(
        { ...APPOINTMENT, title: "Reise nach Berlin" },
        profile,
      ),
    ).toBe(true);
  });
});

describe("findMatchingAppointments list order precedence", () => {
  it("uses the order of the profile list instead of numeric priority", () => {
    const firstProfile = createProfile({
      id: "rule-first",
      field: "title",
      operator: "contains",
      value: "Vacation",
      caseSensitive: false,
    });
    firstProfile.priority = 50;

    const secondProfile = createProfile({
      id: "rule-second",
      field: "title",
      operator: "contains",
      value: "Vacation",
      caseSensitive: false,
    });
    secondProfile.id = "profile-second";
    secondProfile.priority = 10;

    const matches = findMatchingAppointments(
      [APPOINTMENT],
      [firstProfile, secondProfile],
    );

    expect(matches).toHaveLength(1);
    expect(matches[0]?.profile.id).toBe(firstProfile.id);
  });
});
