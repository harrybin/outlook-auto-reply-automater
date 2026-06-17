import { Page } from "@playwright/test";

/**
 * Test utilities for mocking Office context, Graph API, and common test patterns.
 */

/**
 * Mock the Office.onReady callback and Office context.
 * Returns an object to track calls and simulate responses.
 */
export function mockOfficeContext(page: Page) {
  const context = {
    readyCallbacks: [] as any[],
    callCount: 0,
  };

  page.addInitScript(() => {
    // Setup global Office mock
    (window as any).Office = {
      onReady: (callback: (info: any) => void) => {
        callback({
          host: "Outlook",
          platform: "Web",
        });
      },
      context: {
        mailbox: {
          item: {},
          userProfile: {
            displayName: "Test User",
            emailAddress: "test@example.com",
          },
          diagnostics: {
            hostVersion: "16.0.0",
          },
        },
      },
    };
  });

  return context;
}

/**
 * Mock the Microsoft Graph client and expose it globally for service calls.
 * Returns an object to track API calls and configure responses.
 */
export function mockGraphClient(page: Page) {
  const responses: Record<string, any> = {
    "/me": {
      id: "user-id-123",
      displayName: "Test User",
      mail: "test@example.com",
    },
    "/me/mailboxSettings/automaticRepliesSetting": {
      isScheduled: false,
      scheduledStartDateTime: null,
      scheduledEndDateTime: null,
      externalAudience: "all",
      internalReplyMessage: "",
      externalReplyMessage: "",
    },
    "/me/calendarview": {
      value: [
        {
          id: "event-1",
          subject: "Team Meeting",
          start: new Date(),
          end: new Date(Date.now() + 3600000),
          isReminderOn: true,
          isAllDay: false,
          showAs: "busy",
          categories: ["work"],
        },
      ],
    },
    "/me/presence": {
      id: "user-id-123",
      availability: "Available",
      activity: "Available",
    },
  };

  const callLog: any[] = [];

  page.addInitScript(() => {
    // Mock the Graph client globally
    (window as any).__MOCK_GRAPH_CLIENT__ = {
      api: (path: string) => ({
        get: async () => {
          (window as any).__GRAPH_CALL_LOG__ ||= [];
          (window as any).__GRAPH_CALL_LOG__.push({ method: "GET", path });
          return (window as any).__MOCK_GRAPH_RESPONSES__?.[path] || {};
        },
        post: async (body: any) => {
          (window as any).__GRAPH_CALL_LOG__ ||= [];
          (window as any).__GRAPH_CALL_LOG__.push({
            method: "POST",
            path,
            body,
          });
          return { status: 204 };
        },
        patch: async (body: any) => {
          (window as any).__GRAPH_CALL_LOG__ ||= [];
          (window as any).__GRAPH_CALL_LOG__.push({
            method: "PATCH",
            path,
            body,
          });
          return { status: 200 };
        },
        put: async (body: any) => {
          (window as any).__GRAPH_CALL_LOG__ ||= [];
          (window as any).__GRAPH_CALL_LOG__.push({
            method: "PUT",
            path,
            body,
          });
          return { status: 200 };
        },
        delete: async () => {
          (window as any).__GRAPH_CALL_LOG__ ||= [];
          (window as any).__GRAPH_CALL_LOG__.push({ method: "DELETE", path });
          return { status: 204 };
        },
      }),
    };

    // Set up mock responses globally
    (window as any).__MOCK_GRAPH_RESPONSES__ = {};
  });

  // Configure responses from within Playwright context
  page.evaluateHandle((resp: any) => {
    (window as any).__MOCK_GRAPH_RESPONSES__ = resp;
  }, responses);

  return {
    responses,
    callLog,

    /**
     * Get all API calls made by the Graph client.
     */
    async getCalls() {
      return page.evaluate(() => (window as any).__GRAPH_CALL_LOG__ || []);
    },

    /**
     * Get calls filtered by method and path.
     */
    async getCallsFor(method: string, pathPattern: string | RegExp) {
      const calls = await this.getCalls();
      const regex =
        typeof pathPattern === "string" ? new RegExp(pathPattern) : pathPattern;
      return calls.filter(
        (c: any) => c.method === method && regex.test(c.path),
      );
    },

    /**
     * Configure a mock response for a given endpoint.
     */
    async setResponse(path: string, response: any) {
      await page.evaluate(
        ({ path, resp }) => {
          (window as any).__MOCK_GRAPH_RESPONSES__[path] = resp;
        },
        { path, resp: response },
      );
    },

    /**
     * Clear all call logs.
     */
    async clearCallLog() {
      await page.evaluate(() => {
        (window as any).__GRAPH_CALL_LOG__ = [];
      });
    },
  };
}

/**
 * Wait for a condition to be true with polling.
 */
export async function waitFor(
  predicate: () => boolean | Promise<boolean>,
  timeoutMs = 3000,
  pollIntervalMs = 100,
): Promise<void> {
  const startTime = Date.now();
  while (Date.now() - startTime < timeoutMs) {
    if (await predicate()) {
      return;
    }
    await new Promise((resolve) => setTimeout(resolve, pollIntervalMs));
  }
  throw new Error(`Timeout waiting for condition after ${timeoutMs}ms`);
}

/**
 * Create a test message object.
 */
export function createTestMessage(overrides?: Partial<any>) {
  return {
    name: "Test Reply",
    subject: "Re: Meeting",
    body: "Thanks for the message, I will respond when available.",
    isHtml: false,
    ...overrides,
  };
}

/**
 * Create a test profile object with default rules.
 */
export function createTestProfile(messageId: string, overrides?: Partial<any>) {
  return {
    name: "Test Profile",
    enabled: true,
    priority: 1,
    messageId,
    matchRules: {
      keywordRules: [
        {
          id: "rule-1",
          field: "subject",
          operator: "contains",
          value: "meeting",
          caseSensitive: false,
        },
      ],
      combinator: "any",
      durationRule: null,
      busyStatusRule: null,
    },
    timingSettings: {
      enableBefore: false,
      hoursBeforeAppointment: 0,
      enableAfter: true,
      hoursAfterAppointment: 2,
    },
    teamsSettings: {
      updateTeamsStatus: false,
      teamsStatus: "away",
      teamsMessage: "In a meeting",
      restoreTeamsStatusOnEnd: false,
    },
    ...overrides,
  };
}

/**
 * Simulate a calendar appointment for testing.
 */
export function createTestAppointment(overrides?: Partial<any>) {
  const now = new Date();
  const endTime = new Date(now.getTime() + 3600000); // 1 hour from now

  return {
    id: "apt-" + Math.random().toString(36).slice(2, 9),
    title: "Team Sync",
    subject: "Team Sync",
    start: now,
    end: endTime,
    isAllDay: false,
    durationMinutes: 60,
    busyStatus: "busy",
    location: undefined,
    organizer: undefined,
    categories: [],
    ...overrides,
  };
}

/**
 * Take a screenshot for visual regression testing.
 */
export async function takeScreenshot(page: Page, name: string) {
  const snapshotPath = `tests/e2e/__screenshots__/${name}.png`;
  await page.screenshot({ path: snapshotPath, fullPage: true });
  return snapshotPath;
}
