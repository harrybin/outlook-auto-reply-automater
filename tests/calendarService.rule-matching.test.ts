import { describe, expect, it } from "vitest";
import type { AppointmentInfo, AutomationProfile, KeywordRule } from "@/taskpane/types";
import { appointmentMatchesProfile } from "@/taskpane/services/calendarService";

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
      caseSensitive: true,
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
