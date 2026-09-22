import { render, waitFor } from "@testing-library/react";
import { createElement } from "react";
import { afterEach, describe, expect, it, vi } from "vitest";
import { App } from "@/taskpane/App";

vi.mock("@/taskpane/components/AutoReplyList", () => ({
  AutoReplyList: () => createElement("div", { "data-testid": "auto-reply-list" }),
}));

vi.mock("@/taskpane/components/ProfileList", () => ({
  ProfileList: () => createElement("div", { "data-testid": "profile-list" }),
}));

describe("App Office theme handlers", () => {
  const originalOffice = globalThis.Office;

  afterEach(() => {
    globalThis.Office = originalOffice;
  });

  it("registers and removes the Office theme handler when supported", async () => {
    const addHandlerAsync = vi.fn();
    const removeHandlerAsync = vi.fn();

    globalThis.Office = {
      EventType: {
        OfficeThemeChanged: "officeThemeChanged",
      },
      onReady: vi.fn(() => Promise.resolve({ host: "Outlook", platform: "Web" })),
      context: {
        officeTheme: {
          bodyBackgroundColor: "#ffffff",
        },
        requirements: {
          isSetSupported: vi.fn(() => true),
        },
        mailbox: {
          addHandlerAsync,
          removeHandlerAsync,
        },
      },
    } as unknown as typeof Office;

    const { unmount } = render(createElement(App));

    await waitFor(() => {
      expect(addHandlerAsync).toHaveBeenCalledWith(
        globalThis.Office.EventType.OfficeThemeChanged,
        expect.any(Function),
      );
    });

    unmount();

    expect(removeHandlerAsync).toHaveBeenCalledWith(
      globalThis.Office.EventType.OfficeThemeChanged,
    );
  });
});
