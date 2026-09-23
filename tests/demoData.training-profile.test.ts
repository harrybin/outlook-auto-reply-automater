import { describe, expect, it } from "vitest";
import { DEMO_DATA } from "@/demo/demoData";
import { appointmentMatchesProfile } from "@/taskpane/services/calendarService";
import type { AppointmentInfo } from "@/taskpane/types";

describe("demo training profile", () => {
  it("sets the Teams status without enabling an Outlook auto-reply", () => {
    const profile = DEMO_DATA.automationProfiles.find(
      (candidate) => candidate.id === "demo-profile-4",
    );

    expect(profile).toMatchObject({
      enabled: true,
      enableAutoReply: false,
      teamsStatusSettings: {
        enabled: true,
        statusWhenActive: "DoNotDisturb",
      },
    });
  });

  it("matches travel appointments by the configured travel terms", () => {
    const profile = DEMO_DATA.automationProfiles.find(
      (candidate) => candidate.id === "demo-profile-5",
    );
    const appointment: AppointmentInfo = {
      id: "appointment-1",
      title: "Zugfahrt nach Berlin",
      start: "2026-09-22T09:00:00.000Z",
      end: "2026-09-22T11:00:00.000Z",
      busyStatus: "busy",
      isAllDay: false,
      durationMinutes: 120,
      categories: [],
    };

    expect(profile).toBeDefined();
    expect(appointmentMatchesProfile(appointment, profile!)).toBe(true);
  });

  it("uses recipient and duration defaults that match the recommendations", () => {
    const vacation = DEMO_DATA.automationProfiles.find(
      (candidate) => candidate.id === "demo-profile-1",
    );
    const meeting = DEMO_DATA.automationProfiles.find(
      (candidate) => candidate.id === "demo-profile-2",
    );
    const traveling = DEMO_DATA.automationProfiles.find(
      (candidate) => candidate.id === "demo-profile-5",
    );

    expect(vacation).toMatchObject({
      autoReplyAudience: "both",
      matchRules: {
        durationRule: { enabled: true, minHours: 8 },
      },
    });
    expect(meeting).toMatchObject({
      autoReplyAudience: "internal",
      matchRules: {
        durationRule: { enabled: true, maxHours: 479 / 60 },
      },
    });
    expect(traveling).toMatchObject({
      autoReplyAudience: "internal",
      matchRules: {
        durationRule: { enabled: true, maxHours: 479 / 60 },
      },
    });
  });
});
