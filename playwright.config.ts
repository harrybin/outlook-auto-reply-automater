import { defineConfig, devices } from "@playwright/test";

/**
 * Playwright configuration for E2E testing the Outlook Auto-Reply Automater add-in.
 * Tests run against the React app with mocked Office context and Graph APIs.
 */
export default defineConfig({
  testDir: "./tests/e2e",
  fullyParallel: true,
  forbidOnly: !!process.env.CI,
  retries: process.env.CI ? 2 : 0,
  workers: process.env.CI ? 1 : undefined,
  timeout: 60_000,
  reporter: [
    ["html"],
    ["list"],
    ["json", { outputFile: "test-results/e2e-results.json" }],
  ],
  use: {
    baseURL: "http://localhost:3000",
    trace: "on-first-retry",
    screenshot: "only-on-failure",
  },

  webServer: {
    command: "npx vite --strictPort --port 3000",
    url: "http://localhost:3000/src/taskpane/index.html",
    reuseExistingServer: !process.env.CI,
    timeout: 120_000,
  },

  projects: [
    {
      name: "edge",
      use: { ...devices["Desktop Chrome"], channel: "msedge" },
    },
  ],
});
