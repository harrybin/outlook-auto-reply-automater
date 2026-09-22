import { describe, expect, it } from "vitest";
import type { AppointmentInfo, AutomationProfile } from "@/taskpane/types";
import { renderMessageTemplate } from "@/taskpane/services/messageTemplateService";

const APPOINTMENT: AppointmentInfo = {
  id: "appointment-1",
  title: "Urlaub",
  start: "2026-09-21T09:00:00.000Z",
  end: "2026-09-25T17:00:00.000Z",
  location: "Berlin",
  busyStatus: "outOfOffice",
  isAllDay: false,
  durationMinutes: 7_680,
  categories: [],
};

const TRAINING_PROFILE: AutomationProfile = {
  id: "profile-1",
  name: "Training",
  enabled: true,
  autoReplyMessageId: "message-1",
  priority: 1,
  createdAt: "2026-09-22T00:00:00.000Z",
  updatedAt: "2026-09-22T00:00:00.000Z",
  matchRules: {
    keywordRules: [
      {
        id: "rule-1",
        field: "title",
        operator: "regex",
        value: "/\\b(Training)\\b/i",
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
    enabled: true,
    statusWhenActive: "DoNotDisturb",
    statusMessageWhenActive: "In Training",
    restoreOnEnd: true,
  },
};

describe("renderMessageTemplate", () => {
  it("fills appointment values and calculates the next working day after a Friday", () => {
    const result = renderMessageTemplate(
      "{{appointment.title}} ends {{appointment.end}}. Back {{appointment.nextWorkingDayAfterEnd}} in {{appointment.location}}.",
      APPOINTMENT,
    );

    expect(result).toBe("Urlaub ends 25.09.2026. Back 28.09.2026 in Berlin.");
  });

  it("keeps a template unchanged when no appointment is available", () => {
    const template = "Back {{appointment.nextWorkingDayAfterEnd}}.";

    expect(renderMessageTemplate(template)).toBe(template);
  });

  it("fills a rule placeholder from the first regex capture group", () => {
    const appointment = { ...APPOINTMENT, title: "Training in Berlin" };

    expect(
      renderMessageTemplate(
        "Currently in {{rule.match}}.",
        appointment,
        TRAINING_PROFILE,
      ),
    ).toBe("Currently in Training.");
  });
});
