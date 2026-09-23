import { describe, expect, it, vi } from "vitest";
import { enableAutoReply } from "@/taskpane/services/autoReplyService";
import type {
  AutoReplyAudience,
  AutoReplyMessage,
  AutomationProfile,
} from "@/taskpane/types";

const MESSAGE: AutoReplyMessage = {
  id: "message-1",
  name: "Away",
  subject: "Away",
  body: "I am away.",
  isHtml: false,
  createdAt: "2026-01-01T00:00:00.000Z",
  updatedAt: "2026-01-01T00:00:00.000Z",
};

function createProfile(
  autoReplyAudience: AutoReplyAudience,
): AutomationProfile {
  return {
    id: "profile-1",
    name: "Away",
    enabled: true,
    enableAutoReply: true,
    autoReplyAudience,
    autoReplyMessageId: MESSAGE.id,
    matchRules: {
      keywordRules: [],
      durationRule: { enabled: true, minHours: 8 },
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
    priority: 1,
    createdAt: "2026-01-01T00:00:00.000Z",
    updatedAt: "2026-01-01T00:00:00.000Z",
  };
}

function createGraphClient() {
  const patch = vi.fn().mockResolvedValue(undefined);
  return {
    patch,
    client: {
      api: vi.fn(() => ({ patch })),
    },
  };
}

describe("enableAutoReply recipient audience", () => {
  it.each([
    [
      "internal",
      {
        status: "alwaysEnabled",
        externalAudience: "none",
        internalReplyMessage: { contentType: "text", content: "I am away." },
      },
    ],
    [
      "external",
      {
        status: "alwaysEnabled",
        externalAudience: "all",
        externalReplyMessage: { contentType: "text", content: "I am away." },
      },
    ],
    [
      "both",
      {
        status: "alwaysEnabled",
        externalAudience: "all",
        internalReplyMessage: { contentType: "text", content: "I am away." },
        externalReplyMessage: { contentType: "text", content: "I am away." },
      },
    ],
  ] as const)("builds the %s recipient payload", async (audience, expected) => {
    const graph = createGraphClient();

    await enableAutoReply(
      graph.client,
      MESSAGE,
      undefined,
      createProfile(audience),
    );

    expect(graph.patch).toHaveBeenCalledWith(expected);
  });
});
