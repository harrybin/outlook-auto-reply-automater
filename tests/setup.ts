import { afterEach, vi } from "vitest";
import "@testing-library/jest-dom";

/**
 * Vitest setup file: Global configuration for unit and component tests.
 * Mocks Office context, localStorage, and browser APIs.
 */

// Mock localStorage
const localStorageMock: Storage = (() => {
  let store: Record<string, string> = {};
  return {
    getItem: (key: string) => store[key] || null,
    setItem: (key: string, value: string) => {
      store[key] = value.toString();
    },
    removeItem: (key: string) => {
      delete store[key];
    },
    clear: () => {
      store = {};
    },
    get length() {
      return Object.keys(store).length;
    },
    key: (index: number) => Object.keys(store)[index] || null,
  };
})();

Object.defineProperty(window, "localStorage", {
  value: localStorageMock,
});

// Mock Office.onReady (Office Add-in initialization)
window.Office = {
  onReady: vi.fn((callback) => {
    // Call immediately with default context
    callback({ host: "Outlook", platform: "Web" } as any);
  }),
  context: {
    mailbox: {
      item: {} as any,
      userProfile: {} as any,
      diagnostics: {} as any,
    } as any,
  } as any,
} as any;

// Mock navigator.geolocation
Object.defineProperty(navigator, "geolocation", {
  writable: true,
  value: {
    getCurrentPosition: vi.fn((success) => {
      // Mock successful geolocation with default coords
      success({
        coords: {
          latitude: 37.7749,
          longitude: -122.4194,
          accuracy: 10,
          altitude: null,
          altitudeAccuracy: null,
          heading: null,
          speed: null,
        },
        timestamp: Date.now(),
      } as GeolocationPosition);
    }),
    watchPosition: vi.fn(() => 1),
    clearWatch: vi.fn(),
  } as any,
});

// Clean up localStorage between tests
afterEach(() => {
  localStorageMock.clear();
});
