import { Page } from "@playwright/test";

/**
 * Custom Playwright test helpers for testing the Outlook Auto-Reply Automater.
 * Provides AppPage helper class for interacting with the app.
 */

/**
 * AppPage wraps Playwright Page with custom helpers for interacting with the app.
 */
export class AppPage {
  constructor(readonly page: Page) {}

  /**
   * Navigate to the app and wait for initialization.
   */
  async goto() {
    await this.setupMocks();
    await this.page.goto("/");
    await this.waitForAppReady();
  }

  /**
   * Setup global mocks for Office and Graph APIs.
   */
  private async setupMocks() {
    await this.page.addInitScript(() => {
      // Mock Office context
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
          },
        },
      };

      // Mock Graph client
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
        }),
      };

      (window as any).__MOCK_GRAPH_RESPONSES__ = {};
    });
  }

  /**
   * Wait for the app to be ready (FluentUI provider, store hydrated).
   */
  async waitForAppReady() {
    try {
      await this.page.waitForTimeout(500); // Brief pause for app initialization
      await this.page.waitForLoadState("domcontentloaded");
    } catch {
      // App may be ready even if waits timeout
    }
  }

  /**
   * Get the Zustand store state by reading from a data attribute or window object.
   */
  async getStoreState() {
    const stateJson = await this.page.evaluate(() => {
      return (window as any).__STORE_STATE__ || null;
    });
    return stateJson ? JSON.parse(stateJson) : null;
  }

  /**
   * Wait for the store to be updated (used after CRUD operations).
   */
  async waitForStoreUpdate(
    predicate: (state: any) => boolean,
    timeoutMs = 3000,
  ) {
    const startTime = Date.now();
    while (Date.now() - startTime < timeoutMs) {
      const state = await this.getStoreState();
      if (state && predicate(state)) {
        return state;
      }
      await this.page.waitForTimeout(100);
    }
    throw new Error(
      `Store update timeout: predicate not satisfied after ${timeoutMs}ms`,
    );
  }

  /**
   * Fill a form with values using accessibility-friendly selectors.
   */
  async fillForm(fields: Record<string, string>) {
    for (const [label, value] of Object.entries(fields)) {
      const field = this.page
        .locator(`label:has-text("${label}"), [aria-label="${label}"]`)
        .first();
      const input = field.locator("..").locator("input, textarea").first();
      await input.fill(value);
    }
  }

  /**
   * Click a button by accessible name.
   */
  async clickButton(name: string | RegExp) {
    const button = this.page
      .locator(
        `button:has-text("${name}"), [role="button"]:has-text("${name}")`,
      )
      .first();
    await button.click();
  }

  /**
   * Check if an element with text is visible.
   */
  async isTextVisible(text: string | RegExp): Promise<boolean> {
    try {
      const selector =
        typeof text === "string" ? `text=${text}` : `text=/${text.source}/`;
      await this.page
        .locator(selector)
        .first()
        .waitFor({ state: "visible", timeout: 1000 });
      return true;
    } catch {
      return false;
    }
  }

  /**
   * Get all rows from a list (e.g., AutoReplyList, ProfileList).
   */
  async getListItems(listSelector: string): Promise<(string | null)[]> {
    const items = await this.page.locator(listSelector).locator("> *").all();
    return Promise.all(items.map((item) => item.textContent()));
  }

  /**
   * Get all API calls made by the Graph client.
   */
  async getGraphCalls() {
    return this.page.evaluate(() => (window as any).__GRAPH_CALL_LOG__ || []);
  }
}

/**
 * Create an AppPage instance for testing.
 */
export async function createAppPage(page: Page): Promise<AppPage> {
  const appPage = new AppPage(page);
  await appPage.goto();
  return appPage;
}

// AppPage class is exported above; test and expect should be imported directly from @playwright/test
