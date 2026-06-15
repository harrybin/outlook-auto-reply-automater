// ─── Auto-Reply Message Types ────────────────────────────────────────────────

export interface AutoReplyMessage {
  id: string;
  name: string;
  subject: string;
  body: string;
  isHtml: boolean;
  createdAt: string;
  updatedAt: string;
}

// ─── Appointment Matching Rules ───────────────────────────────────────────────

export type AppointmentMatchField =
  | "title"
  | "body"
  | "location"
  | "organizer"
  | "category";

export type AppointmentMatchOperator =
  | "contains"
  | "startsWith"
  | "endsWith"
  | "equals"
  | "regex";

export type AppointmentBusyStatus =
  | "free"
  | "tentative"
  | "busy"
  | "outOfOffice"
  | "workingElsewhere";

export interface KeywordRule {
  id: string;
  field: AppointmentMatchField;
  operator: AppointmentMatchOperator;
  value: string;
  caseSensitive: boolean;
}

export interface DurationRule {
  enabled: boolean;
  minMinutes?: number;
  maxMinutes?: number;
}

export interface BusyStatusRule {
  enabled: boolean;
  statuses: AppointmentBusyStatus[];
}

export interface AppointmentMatchRules {
  keywordRules: KeywordRule[];
  durationRule: DurationRule;
  busyStatusRule: BusyStatusRule;
  combinator: "AND" | "OR"; // how multiple rules are combined
}

// ─── Timing Settings ─────────────────────────────────────────────────────────

export interface TimingSettings {
  enableBefore: boolean;
  hoursBeforeAppointment: number;
  enableAfter: boolean;
  hoursAfterAppointment: number;
}

// ─── Location / Network Settings ─────────────────────────────────────────────

export type LocationTriggerType = "geofence" | "wifi" | "noInternet";

export interface GeofenceTrigger {
  type: "geofence";
  label: string;
  latitude: number;
  longitude: number;
  radiusMeters: number;
}

export interface WifiTrigger {
  type: "wifi";
  /** SSID or partial SSID to match */
  ssid: string;
}

export interface NoInternetTrigger {
  type: "noInternet";
}

export type LocationTrigger = GeofenceTrigger | WifiTrigger | NoInternetTrigger;

export interface LocationRule {
  id: string;
  enabled: boolean;
  trigger: LocationTrigger;
  autoReplyMessageId: string;
  teamsStatus?: TeamsPresenceStatus;
  teamsStatusMessage?: string;
}

export interface LocationSettings {
  enabled: boolean;
  pollIntervalSeconds: number;
  rules: LocationRule[];
}

// ─── MS Teams Presence ────────────────────────────────────────────────────────

export type TeamsPresenceStatus =
  | "Available"
  | "Busy"
  | "DoNotDisturb"
  | "BeRightBack"
  | "Away"
  | "Offline";

export interface TeamsStatusSettings {
  enabled: boolean;
  /** Set Teams status when auto-reply is active */
  statusWhenActive: TeamsPresenceStatus;
  statusMessageWhenActive: string;
  /** Restore original Teams status when auto-reply ends */
  restoreOnEnd: boolean;
}

// ─── Automation Profile ───────────────────────────────────────────────────────

/** A complete automation rule combining all conditions and actions */
export interface AutomationProfile {
  id: string;
  name: string;
  enabled: boolean;
  autoReplyMessageId: string;
  matchRules: AppointmentMatchRules;
  timingSettings: TimingSettings;
  teamsStatusSettings: TeamsStatusSettings;
  priority: number; // Lower = higher priority when multiple profiles match
  createdAt: string;
  updatedAt: string;
}

// ─── Copilot Suggestion ───────────────────────────────────────────────────────

export interface CopilotSuggestion {
  id: string;
  name: string;
  subject: string;
  body: string;
  matchedAppointment?: AppointmentInfo;
  confidence: number; // 0-1
}

export interface AppointmentInfo {
  id: string;
  title: string;
  start: string;
  end: string;
  location?: string;
  busyStatus: AppointmentBusyStatus;
  isAllDay: boolean;
  durationMinutes: number;
  categories: string[];
}

// ─── Active Auto-Reply State ──────────────────────────────────────────────────

export interface ActiveAutoReply {
  profileId: string;
  appointmentId?: string;
  message: AutoReplyMessage;
  activatedAt: string;
  scheduledEndAt?: string;
  trigger: "appointment" | "location" | "manual";
}

// ─── Application Settings ─────────────────────────────────────────────────────

export interface AppSettings {
  autoReplyMessages: AutoReplyMessage[];
  automationProfiles: AutomationProfile[];
  locationSettings: LocationSettings;
  /** ID of currently active auto-reply (null if none) */
  activeAutoReplyId: string | null;
}
