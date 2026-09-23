import { render, screen, waitFor } from "@testing-library/react";
import { createElement } from "react";
import { beforeEach, describe, expect, it, vi } from "vitest";
import {
  buildCopilotDraftForRule,
  canCreateOutlookMessageForRule,
  getProfileRecommendation,
  ProfileList,
} from "@/taskpane/components/ProfileList";
import type { AutoReplyMessage, AutomationProfile } from "@/taskpane/types";
import { useStore } from "@/taskpane/useStore";

const PROFILE: AutomationProfile = {
  id: "profile-1",
  name: "OOO <rule>",
  enabled: true,
  autoReplyMessageId: "msg-1",
  priority: 5,
  createdAt: "2026-01-01T00:00:00.000Z",
  updatedAt: "2026-01-01T00:00:00.000Z",
  matchRules: {
    combinator: "AND",
    keywordRules: [
      {
        id: "kw-1",
        field: "title",
        operator: "contains",
        value: "Vacation",
        caseSensitive: false,
      },
    ],
    durationRule: { enabled: true, minHours: 0.5, maxHours: 2 },
    busyStatusRule: { enabled: true, statuses: ["busy", "outOfOffice"] },
  },
  timingSettings: {
    enableBefore: true,
    hoursBeforeAppointment: 2,
    enableAfter: true,
    hoursAfterAppointment: 3,
  },
  teamsStatusSettings: {
    enabled: true,
    statusWhenActive: "Away",
    statusMessageWhenActive: "In a meeting",
    restoreOnEnd: true,
  },
};

describe("buildCopilotDraftForRule", () => {
  it("builds an Outlook draft with Copilot context and rule details", () => {
    const message: AutoReplyMessage = {
      id: "msg-1",
      name: "Vacation response",
      subject: "Out of office",
      body: "<p>Thanks for your message.</p>",
      isHtml: true,
      createdAt: "2026-01-01T00:00:00.000Z",
      updatedAt: "2026-01-01T00:00:00.000Z",
    };

    const draft = buildCopilotDraftForRule(PROFILE, message);

    expect(draft.subject).toBe("Out of office");
    expect(draft.htmlBody).toContain("Copilot context for Outlook");
    expect(draft.htmlBody).toContain("OOO &lt;rule&gt;");
    expect(draft.htmlBody).toContain('title contains "Vacation"');
    expect(draft.htmlBody).toContain("Duration: 0.5-2 hours");
    expect(draft.htmlBody).toContain("Busy status: busy, outOfOffice");
    expect(draft.htmlBody).toContain("<p>Thanks for your message.</p>");
  });

  describe("getProfileRecommendation", () => {
    it("suggests Teams-only behavior when no duration filter is set", () => {
      expect(
        getProfileRecommendation({
          ...PROFILE,
          matchRules: {
            ...PROFILE.matchRules,
            durationRule: { enabled: false },
          },
        }),
      ).toBe("teamsOnly");
    });

    it("suggests internal replies for appointments shorter than eight hours", () => {
      expect(
        getProfileRecommendation({
          ...PROFILE,
          autoReplyAudience: "both",
          matchRules: {
            ...PROFILE.matchRules,
            durationRule: { enabled: true, maxHours: 479 / 60 },
          },
        }),
      ).toBe("internalOnly");
    });

    it("suggests an eight-hour minimum for external replies", () => {
      expect(
        getProfileRecommendation({
          ...PROFILE,
          autoReplyAudience: "external",
          matchRules: {
            ...PROFILE.matchRules,
            durationRule: { enabled: true, minHours: 1 },
          },
        }),
      ).toBe("minimumEightHours");
    });

    it("does not suggest changes for long internal-and-external replies", () => {
      expect(
        getProfileRecommendation({
          ...PROFILE,
          autoReplyAudience: "both",
          matchRules: {
            ...PROFILE.matchRules,
            durationRule: { enabled: true, minHours: 8 },
          },
        }),
      ).toBeNull();
    });
  });

  it("escapes and preserves line breaks for plain text message bodies", () => {
    const plainTextMessage: AutoReplyMessage = {
      id: "msg-2",
      name: "Plain text",
      subject: "Auto reply",
      body: "Line <1>\nLine 2",
      isHtml: false,
      createdAt: "2026-01-01T00:00:00.000Z",
      updatedAt: "2026-01-01T00:00:00.000Z",
    };

    const draft = buildCopilotDraftForRule(PROFILE, plainTextMessage);

    expect(draft.htmlBody).toContain("Line &lt;1&gt;<br/>Line 2");
  });
});

describe("canCreateOutlookMessageForRule", () => {
  beforeEach(() => {
    useStore.setState((state) => ({
      ...state,
      autoReplyMessages: [],
      automationProfiles: [],
    }));
  });

  it("returns true only when Outlook compose API is available", () => {
    const globalWithOffice = globalThis as unknown as {
      Office?: { context?: { mailbox?: { displayNewMessageForm?: () => void } } };
    };
    const office = globalWithOffice.Office;
    globalWithOffice.Office = {
      context: {
        mailbox: {
          displayNewMessageForm: () => undefined,
        },
      },
    };

    expect(canCreateOutlookMessageForRule()).toBe(true);

    globalWithOffice.Office = {
      context: {
        mailbox: {},
      },
    };

    expect(canCreateOutlookMessageForRule()).toBe(false);

    globalWithOffice.Office = office;
  });

  it("re-enables Create Message after Office.onReady populates mailbox", async () => {
    let resolveOfficeReady: ((value: unknown) => void) | undefined;
    const officeReady = new Promise((resolve) => {
      resolveOfficeReady = resolve;
    });
    const displayNewMessageForm = vi.fn();
    const globalWithOffice = globalThis as unknown as {
      Office?: {
        onReady?: () => Promise<unknown>;
        context?: { mailbox?: { displayNewMessageForm?: typeof displayNewMessageForm } };
      };
    };
    const office = globalWithOffice.Office;

    globalWithOffice.Office = {
      onReady: vi.fn(() => officeReady),
      context: {
        mailbox: {},
      },
    };

    useStore.setState((state) => ({
      ...state,
      autoReplyMessages: [
        {
          id: "msg-1",
          name: "Vacation response",
          subject: "Out of office",
          body: "Thanks for your message.",
          isHtml: false,
          createdAt: "2026-01-01T00:00:00.000Z",
          updatedAt: "2026-01-01T00:00:00.000Z",
        },
      ],
      automationProfiles: [PROFILE],
    }));

    render(createElement(ProfileList));

    const createMessageButton = screen.getByRole("button", {
      name: "Preview",
    });
    expect(createMessageButton).toBeDisabled();

    globalWithOffice.Office.context = {
      mailbox: {
        displayNewMessageForm,
      },
    };
    resolveOfficeReady?.({ host: "Outlook", platform: "Web" });

    await waitFor(() => {
      expect(createMessageButton).toBeEnabled();
    });

    globalWithOffice.Office = office;
  });
});
