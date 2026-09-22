import type {
  AppointmentInfo,
  AppointmentMatchField,
  AutomationProfile,
  KeywordRule,
} from "../types";

function formatDate(value: string): string {
  return new Date(value).toLocaleDateString("de-DE", {
    day: "2-digit",
    month: "2-digit",
    year: "numeric",
  });
}

function nextWorkingDay(value: string): string {
  const date = new Date(value);
  date.setDate(date.getDate() + 1);

  while (date.getDay() === 0 || date.getDay() === 6) {
    date.setDate(date.getDate() + 1);
  }

  return formatDate(date.toISOString());
}

function getFieldValue(
  appointment: AppointmentInfo,
  field: AppointmentMatchField,
): string {
  switch (field) {
    case "title":
      return appointment.title;
    case "location":
      return appointment.location ?? "";
    case "category":
      return appointment.categories.join(" ");
    case "body":
    case "organizer":
      return "";
  }
}

function regexFromRule(rule: KeywordRule): RegExp | null {
  try {
    const literal = rule.value.match(/^\/([\s\S]*)\/([dgimsuvy]*)$/);
    const flags = literal ? literal[2] : rule.caseSensitive ? "" : "i";
    const normalizedFlags = rule.caseSensitive
      ? flags.replace(/i/g, "")
      : flags.includes("i")
        ? flags
        : `${flags}i`;
    return new RegExp(literal ? literal[1] : rule.value, normalizedFlags);
  } catch {
    return null;
  }
}

function getRegexRuleMatch(
  appointment: AppointmentInfo,
  profile?: AutomationProfile,
): string {
  if (!profile) return "";

  for (const rule of profile.matchRules.keywordRules) {
    if (rule.operator !== "regex") continue;
    const match = regexFromRule(rule)?.exec(
      getFieldValue(appointment, rule.field),
    );
    if (match) return match[1] ?? match[0];
  }

  return "";
}

/** Replaces supported appointment placeholders in an auto-reply message. */
export function renderMessageTemplate(
  template: string,
  appointment?: AppointmentInfo,
  profile?: AutomationProfile,
): string {
  if (!appointment) return template;

  const placeholders: Record<string, string> = {
    "{{appointment.title}}": appointment.title,
    "{{appointment.start}}": formatDate(appointment.start),
    "{{appointment.end}}": formatDate(appointment.end),
    "{{appointment.nextWorkingDayAfterEnd}}": nextWorkingDay(appointment.end),
    "{{appointment.location}}": appointment.location ?? "",
    "{{rule.match}}": getRegexRuleMatch(appointment, profile),
  };

  return Object.entries(placeholders).reduce(
    (result, [placeholder, value]) => result.split(placeholder).join(value),
    template,
  );
}
