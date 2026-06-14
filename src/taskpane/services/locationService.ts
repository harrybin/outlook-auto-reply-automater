/**
 * locationService.ts
 *
 * Detects the device's geographic location and network connectivity to
 * trigger location-based auto-reply rules.
 *
 * Network/WiFi detection uses the experimental Network Information API where
 * available and falls back to navigator.onLine.
 */

import type { GeofenceTrigger, LocationRule, LocationTrigger, WifiTrigger } from "../types";

// ─── Geolocation ─────────────────────────────────────────────────────────────

export interface Coordinates {
  latitude: number;
  longitude: number;
  accuracy: number;
}

export function getCurrentPosition(): Promise<Coordinates> {
  return new Promise((resolve, reject) => {
    if (!navigator.geolocation) {
      reject(new Error("Geolocation is not supported by this browser."));
      return;
    }
    navigator.geolocation.getCurrentPosition(
      (pos) =>
        resolve({
          latitude: pos.coords.latitude,
          longitude: pos.coords.longitude,
          accuracy: pos.coords.accuracy,
        }),
      (err) => reject(err),
      { enableHighAccuracy: false, timeout: 10_000, maximumAge: 30_000 }
    );
  });
}

/** Haversine distance in metres between two coordinates */
export function distanceMeters(
  lat1: number,
  lon1: number,
  lat2: number,
  lon2: number
): number {
  const R = 6_371_000;
  const φ1 = (lat1 * Math.PI) / 180;
  const φ2 = (lat2 * Math.PI) / 180;
  const Δφ = ((lat2 - lat1) * Math.PI) / 180;
  const Δλ = ((lon2 - lon1) * Math.PI) / 180;
  const a =
    Math.sin(Δφ / 2) ** 2 +
    Math.cos(φ1) * Math.cos(φ2) * Math.sin(Δλ / 2) ** 2;
  return R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
}

export async function isInsideGeofence(trigger: GeofenceTrigger): Promise<boolean> {
  try {
    const pos = await getCurrentPosition();
    const dist = distanceMeters(
      pos.latitude,
      pos.longitude,
      trigger.latitude,
      trigger.longitude
    );
    return dist <= trigger.radiusMeters;
  } catch {
    return false;
  }
}

// ─── Network / WiFi ───────────────────────────────────────────────────────────

/** Returns the current WiFi SSID using the Network Information API (experimental). */
export function getCurrentSsid(): string | null {
  // The Network Information API does not expose SSID in browsers for privacy.
  // In Electron / Office desktop environments the host app may expose it via
  // a custom API.  We return null when unavailable.
  return null;
}

export function isOnline(): boolean {
  return navigator.onLine;
}

// ─── Trigger evaluation ───────────────────────────────────────────────────────

export async function evaluateTrigger(trigger: LocationTrigger): Promise<boolean> {
  switch (trigger.type) {
    case "noInternet":
      return !isOnline();

    case "wifi": {
      const wifiTrigger = trigger as WifiTrigger;
      const ssid = getCurrentSsid();
      if (ssid === null) return false; // SSID unavailable
      return ssid.toLowerCase().includes(wifiTrigger.ssid.toLowerCase());
    }

    case "geofence":
      return isInsideGeofence(trigger as GeofenceTrigger);
  }
}

/**
 * Evaluates all enabled location rules and returns those whose trigger
 * condition is currently satisfied.
 */
export async function getMatchingLocationRules(
  rules: LocationRule[]
): Promise<LocationRule[]> {
  const enabled = rules.filter((r) => r.enabled);
  const results = await Promise.all(
    enabled.map(async (rule) => ({
      rule,
      matches: await evaluateTrigger(rule.trigger),
    }))
  );
  return results.filter((r) => r.matches).map((r) => r.rule);
}
