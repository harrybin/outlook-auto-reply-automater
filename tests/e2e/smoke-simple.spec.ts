import { test, expect, Page } from "@playwright/test";

/**
 * Minimal E2E test to verify test infrastructure is working.
 * All setup inlined to avoid import issues.
 */

async function setupAppPage(page: Page) {
  await page.addInitScript(() => {
    // Mock Office context
    (window as any).Office = {
      onReady: (callback: (info: any) => void) => {
        callback({ host: "Outlook", platform: "Web" });
      },
      context: {
        mailbox: { item: {}, userProfile: { displayName: "Test" } },
      },
    };
  });

  await page.goto("/src/taskpane/index.html");
  await page.waitForLoadState("domcontentloaded");
}

test("App loads successfully", async ({ page }) => {
  await setupAppPage(page);
  const title = await page.title();
  expect(title).toBeTruthy();
});

test("Office context is mocked", async ({ page }) => {
  await setupAppPage(page);
  const hasOffice = await page.evaluate(
    () => typeof window.Office !== "undefined",
  );
  expect(hasOffice).toBe(true);
});

test("localStorage is available", async ({ page }) => {
  await setupAppPage(page);
  const hasStorage = await page.evaluate(
    () => typeof localStorage !== "undefined",
  );
  expect(hasStorage).toBe(true);
});
