/**
 * calendarService.ts
 *
 * Reads upcoming calendar appointments via the Microsoft Graph API and
 * evaluates them against automation profiles to find matches.
 */

import type {
  AppointmentInfo,
  AutomationProfile,
  AppointmentBusyStatus,
  KeywordRule,
  AppointmentMatchField,
} from "../types";

interface GraphEvent {
  id: string;
  subject: string;
  bodyPreview: string;
  location?: { displayName?: string };
  organizer?: { emailAddress?: { name?: string } };
  categories: string[];
  start: { dateTime: string; timeZone: string };
  end: { dateTime: string; timeZone: string };
  isAllDay: boolean;
  showAs: string;
}

type GraphClient = {
  api: (path: string) => {
    select: (fields: string) => {
      filter: (expr: string) => {
        top: (n: number) => { get: () => Promise<{ value: GraphEvent[] }> };
      };
    };
  };
};

function mapBusyStatus(showAs: string): AppointmentBusyStatus {
  const map: Record<string, AppointmentBusyStatus> = {
    free: "free",
    tentative: "tentative",
    busy: "busy",
    oof: "outOfOffice",
    workingElsewhere: "workingElsewhere",
  };
  return map[showAs] ?? "busy";
}

function toAppointmentInfo(event: GraphEvent): AppointmentInfo {
  const start = new Date(event.start.dateTime);
  const end = new Date(event.end.dateTime);
  const durationMinutes = Math.round((end.getTime() - start.getTime()) / 60_000);
  return {
    id: event.id,
    title: event.subject ?? "",
    start: event.start.dateTime,
    end: event.end.dateTime,
    location: event.location?.displayName,
    busyStatus: mapBusyStatus(event.showAs),
    isAllDay: event.isAllDay,
    durationMinutes,
    categories: event.categories ?? [],
  };
}

/**
 * Fetches the next `count` upcoming calendar events from Graph.
 */
export async function getUpcomingAppointments(
  graphClient: GraphClient,
  count = 20
): Promise<AppointmentInfo[]> {
  const now = new Date().toISOString();
  const response = await graphClient
    .api("/me/calendarView")
    .select("id,subject,bodyPreview,location,organizer,categories,start,end,isAllDay,showAs")
    .filter(`start/dateTime ge '${now}'`)
    .top(count)
    .get();

  return response.value.map(toAppointmentInfo);
}

// ─── Rule matching helpers ────────────────────────────────────────────────────

function getFieldValue(appointment: AppointmentInfo, field: AppointmentMatchField): string {
  switch (field) {
    case "title":
      return appointment.title;
    case "body":
      return ""; // bodyPreview not in AppointmentInfo; extend if needed
    case "location":
      return appointment.location ?? "";
    case "organizer":
      return ""; // not in AppointmentInfo; extend if needed
    case "category":
      return appointment.categories.join(" ");
  }
}

function matchesKeywordRule(appointment: AppointmentInfo, rule: KeywordRule): boolean {
  const raw = getFieldValue(appointment, rule.field);
  const haystack = rule.caseSensitive ? raw : raw.toLowerCase();
  const needle = rule.caseSensitive ? rule.value : rule.value.toLowerCase();

  switch (rule.operator) {
    case "contains":
      return haystack.includes(needle);
    case "startsWith":
      return haystack.startsWith(needle);
    case "endsWith":
      return haystack.endsWith(needle);
    case "equals":
      return haystack === needle;
    case "regex":
      try {
        return new RegExp(rule.value, rule.caseSensitive ? undefined : "i").test(raw);
      } catch {
        return false;
      }
  }
}

/**
 * Returns true when the given appointment matches all rules in the profile.
 */
export function appointmentMatchesProfile(
  appointment: AppointmentInfo,
  profile: AutomationProfile
): boolean {
  const { matchRules, timingSettings: _ts } = profile;
  const results: boolean[] = [];

  // Keyword rules
  if (matchRules.keywordRules.length > 0) {
    const kwResults = matchRules.keywordRules.map((r) =>
      matchesKeywordRule(appointment, r)
    );
    const kwMatch =
      matchRules.combinator === "AND"
        ? kwResults.every(Boolean)
        : kwResults.some(Boolean);
    results.push(kwMatch);
  }

  // Duration rule
  if (matchRules.durationRule.enabled) {
    const { minMinutes, maxMinutes } = matchRules.durationRule;
    const dur = appointment.durationMinutes;
    const min = minMinutes !== undefined ? dur >= minMinutes : true;
    const max = maxMinutes !== undefined ? dur <= maxMinutes : true;
    results.push(min && max);
  }

  // Busy status rule
  if (matchRules.busyStatusRule.enabled && matchRules.busyStatusRule.statuses.length > 0) {
    results.push(matchRules.busyStatusRule.statuses.includes(appointment.busyStatus));
  }

  if (results.length === 0) return true; // no rules = always match
  return matchRules.combinator === "AND" ? results.every(Boolean) : results.some(Boolean);
}

/**
 * For a list of upcoming appointments, returns appointments that match at
 * least one enabled profile, paired with that profile.
 */
export function findMatchingAppointments(
  appointments: AppointmentInfo[],
  profiles: AutomationProfile[]
): Array<{ appointment: AppointmentInfo; profile: AutomationProfile }> {
  const enabledProfiles = profiles
    .filter((p) => p.enabled)
    .sort((a, b) => a.priority - b.priority);

  const results: Array<{ appointment: AppointmentInfo; profile: AutomationProfile }> = [];

  for (const appointment of appointments) {
    for (const profile of enabledProfiles) {
      if (appointmentMatchesProfile(appointment, profile)) {
        results.push({ appointment, profile });
        break; // use highest-priority matching profile only
      }
    }
  }

  return results;
}
