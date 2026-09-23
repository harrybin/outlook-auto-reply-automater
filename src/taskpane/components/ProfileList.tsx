import { useEffect, useState } from "react";
import { motion } from "framer-motion";
import {
  Button,
  Checkbox,
  Combobox,
  Dialog,
  DialogActions,
  DialogBody,
  DialogContent,
  DialogSurface,
  DialogTitle,
  DialogTrigger,
  Dropdown,
  Field,
  Input,
  makeStyles,
  Option,
  shorthands,
  Switch,
  Text,
  tokens,
} from "@fluentui/react-components";
import {
  Add24Regular,
  Delete24Regular,
  Edit24Regular,
} from "@fluentui/react-icons";
import type {
  AutoReplyMessage,
  AutomationProfile,
  AppointmentBusyStatus,
  AppointmentMatchField,
  AppointmentMatchOperator,
  AutoReplyAudience,
  KeywordRule,
} from "../types";
import { useStore } from "../useStore";
import { nanoid } from "../utils/nanoid";
import { getAccount, getGraphClient, signIn } from "../services/authService";
import { clearTeamsPresence, setTeamsPresence } from "../services/teamsService";

const TEAMS_STATUS_TEST_DURATION_MS = 3000;

const useStyles = makeStyles({
  dialogSurface: {
    width: "calc(100vw - 32px)",
    maxWidth: "960px",
    maxHeight: "calc(100vh - 32px)",
    overflow: "visible",
  },
  dialogBody: {
    minHeight: 0,
    display: "flex",
    flexDirection: "column",
    overflow: "visible",
  },
  dialogContent: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalL,
    paddingTop: tokens.spacingVerticalS,
    overflowX: "visible",
    overflowY: "auto",
  },
  profileName: {
    paddingBottom: tokens.spacingVerticalXS,
  },
  section: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalM,
    backgroundColor: tokens.colorNeutralBackground2,
    ...shorthands.border("1px", "solid", tokens.colorNeutralStroke2),
    ...shorthands.borderRadius(tokens.borderRadiusLarge),
    ...shorthands.padding(
      tokens.spacingVerticalL,
      tokens.spacingHorizontalL,
    ),
  },
  rulesSection: {
    order: 1,
    borderTopColor: tokens.colorBrandStroke1,
    borderTopWidth: "3px",
  },
  emailSection: {
    order: 2,
    borderTopColor: tokens.colorPaletteGreenBorder1,
    borderTopWidth: "3px",
  },
  teamsSection: {
    order: 3,
    borderTopColor: tokens.colorPalettePurpleBorderActive,
    borderTopWidth: "3px",
  },
  sectionHeader: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalXXS,
  },
  sectionTitle: {
    fontSize: tokens.fontSizeBase400,
    fontWeight: tokens.fontWeightSemibold,
    lineHeight: tokens.lineHeightBase400,
  },
  sectionDescription: {
    color: tokens.colorNeutralForeground3,
  },
  sectionContent: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalM,
  },
  subsection: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalS,
    backgroundColor: tokens.colorNeutralBackground1,
    ...shorthands.borderRadius(tokens.borderRadiusMedium),
    ...shorthands.padding(
      tokens.spacingVerticalM,
      tokens.spacingHorizontalM,
    ),
  },
  subsectionHeader: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    gap: tokens.spacingHorizontalM,
    flexWrap: "wrap",
  },
  subsectionTitle: {
    fontWeight: tokens.fontWeightSemibold,
  },
  keywordList: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalS,
  },
  keywordRow: {
    position: "relative",
    display: "grid",
    gridTemplateColumns:
      "minmax(120px, 0.8fr) minmax(140px, 0.9fr) minmax(240px, 2fr) 32px",
    gap: tokens.spacingHorizontalS,
    alignItems: "end",
    backgroundColor: tokens.colorNeutralBackground1,
    ...shorthands.border("1px", "solid", tokens.colorNeutralStroke2),
    ...shorthands.borderRadius(tokens.borderRadiusMedium),
    ...shorthands.padding(
      tokens.spacingVerticalS,
      tokens.spacingHorizontalS,
    ),
    "@media (max-width: 720px)": {
      gridTemplateColumns: "minmax(0, 1fr) minmax(0, 1fr) 32px",
    },
    "@media (max-width: 500px)": {
      gridTemplateColumns: "minmax(0, 1fr) 32px",
    },
  },
  keywordField: {
    minWidth: 0,
  },
  keywordValue: {
    minWidth: 0,
    "@media (max-width: 720px)": {
      gridColumn: "1 / 3",
    },
    "@media (max-width: 500px)": {
      gridColumn: "1",
    },
  },
  deleteRuleButton: {
    color: tokens.colorPaletteRedForeground1,
  },
  addRuleButton: {
    width: "100%",
    ...shorthands.borderStyle("dashed"),
    backgroundColor: tokens.colorSubtleBackground,
  },
  twoColumnGrid: {
    display: "grid",
    gridTemplateColumns: "repeat(2, minmax(0, 1fr))",
    gap: tokens.spacingHorizontalM,
    "@media (max-width: 560px)": {
      gridTemplateColumns: "1fr",
    },
  },
  statusGrid: {
    display: "grid",
    gridTemplateColumns: "repeat(auto-fit, minmax(140px, 1fr))",
    gap: tokens.spacingHorizontalXS,
  },
  helperText: {
    color: tokens.colorNeutralForeground2,
    lineHeight: tokens.lineHeightBase300,
  },
  recommendation: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    gap: tokens.spacingHorizontalM,
    backgroundColor: tokens.colorNeutralBackground3,
    borderLeft: `4px solid ${tokens.colorBrandStroke1}`,
    ...shorthands.borderRadius(tokens.borderRadiusMedium),
    ...shorthands.padding(
      tokens.spacingVerticalS,
      tokens.spacingHorizontalM,
    ),
    marginTop: tokens.spacingVerticalXS,
    "@media (max-width: 560px)": {
      alignItems: "stretch",
      flexDirection: "column",
    },
  },
  dialogActions: {
    position: "sticky",
    bottom: 0,
    zIndex: 1,
    backgroundColor: tokens.colorNeutralBackground1,
    borderTop: `1px solid ${tokens.colorNeutralStroke2}`,
    paddingTop: tokens.spacingVerticalM,
    justifyContent: "flex-end",
    boxShadow: tokens.shadow4,
  },
});

const BUSY_STATUSES: AppointmentBusyStatus[] = [
  "free",
  "tentative",
  "busy",
  "outOfOffice",
  "workingElsewhere",
];
const FIELDS: AppointmentMatchField[] = [
  "title",
  "body",
  "location",
  "organizer",
  "category",
];
const OPERATORS: AppointmentMatchOperator[] = [
  "contains",
  "startsWith",
  "endsWith",
  "equals",
  "regex",
];

const FIELD_LABELS: Record<AppointmentMatchField, string> = {
  title: "Title",
  body: "Body",
  location: "Location",
  organizer: "Organizer",
  category: "Category",
};

const OPERATOR_LABELS: Record<AppointmentMatchOperator, string> = {
  contains: "Contains",
  startsWith: "Starts with",
  endsWith: "Ends with",
  equals: "Equals",
  regex: "Regular expression",
};

const BUSY_STATUS_LABELS: Record<AppointmentBusyStatus, string> = {
  free: "Free",
  tentative: "Tentative",
  busy: "Busy",
  outOfOffice: "Out of office",
  workingElsewhere: "Working elsewhere",
};

type ProfileDraft = Omit<AutomationProfile, "id" | "createdAt" | "updatedAt">;

export type ProfileRecommendation =
  | "teamsOnly"
  | "internalOnly"
  | "minimumEightHours";

interface OutlookMessageDraft {
  subject: string;
  htmlBody: string;
}

interface OutlookMailboxContext {
  displayNewMessageForm?: (message: OutlookMessageDraft) => void;
}

function defaultProfile(): ProfileDraft {
  return {
    name: "",
    enabled: true,
    enableAutoReply: true,
    autoReplyAudience: "both",
    autoReplyMessageId: "",
    priority: 50,
    matchRules: {
      keywordRules: [],
      durationRule: { enabled: false },
      busyStatusRule: { enabled: false, statuses: [] },
      combinator: "AND",
    },
    timingSettings: {
      enableBefore: false,
      hoursBeforeAppointment: 0,
      enableAfter: false,
      hoursAfterAppointment: 0,
    },
    teamsStatusSettings: {
      enabled: false,
      statusWhenActive: "Away",
      statusMessageWhenActive: "",
      restoreOnEnd: true,
    },
  };
}

export function getProfileRecommendation(
  profile: ProfileDraft,
): ProfileRecommendation | null {
  if (profile.enableAutoReply === false) {
    return null;
  }

  const duration = profile.matchRules.durationRule;
  if (!duration.enabled) {
    return "teamsOnly";
  }

  if (
    duration.maxHours !== undefined &&
    duration.maxHours < 8 &&
    profile.autoReplyAudience !== "internal"
  ) {
    return "internalOnly";
  }

  if (
    profile.autoReplyAudience !== "internal" &&
    (duration.minHours === undefined || duration.minHours < 8)
  ) {
    return "minimumEightHours";
  }

  return null;
}

function getAutoReplySummary(profile: AutomationProfile): string {
  if (profile.enableAutoReply === false) {
    return "Teams only";
  }

  switch (profile.autoReplyAudience) {
    case "internal":
      return "Internal auto-reply";
    case "external":
      return "External auto-reply";
    default:
      return "Internal + external auto-reply";
  }
}

function escapeHtml(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function formatMessageBody(message: AutoReplyMessage): string {
  if (message.isHtml) {
    return message.body;
  }

  return escapeHtml(message.body).replace(/\n/g, "<br/>");
}

export function buildCopilotDraftForRule(
  profile: AutomationProfile,
  message: AutoReplyMessage,
): OutlookMessageDraft {
  const keywordRules = profile.matchRules.keywordRules
    .map(
      (rule) =>
        `${escapeHtml(rule.field)} ${escapeHtml(rule.operator)} "${escapeHtml(rule.value)}"`,
    )
    .join(profile.matchRules.combinator === "AND" ? " and " : " or ");
  const hasDurationFilter = profile.matchRules.durationRule.enabled;
  const hasBusyStatusFilter = profile.matchRules.busyStatusRule.enabled;
  const ruleSummary = [
    keywordRules ? `Keyword rules: ${keywordRules}` : null,
    hasDurationFilter
      ? `Duration: ${profile.matchRules.durationRule.minHours ?? 0}-${profile.matchRules.durationRule.maxHours ?? "any"} hours`
      : null,
    hasBusyStatusFilter
      ? `Busy status: ${profile.matchRules.busyStatusRule.statuses.join(", ")}`
      : null,
  ]
    .filter((part): part is string => Boolean(part))
    .join("<br/>");

  return {
    subject: message.subject,
    htmlBody: `<p><strong>Copilot context for Outlook</strong></p>
<p>This draft was created from automation profile <strong>${escapeHtml(profile.name)}</strong>.</p>
<p>Please use Copilot in Outlook to suggest and refine a professional reply matching this rule intent.</p>
${ruleSummary ? `<p>${ruleSummary}</p>` : ""}
<hr/>
${formatMessageBody(message)}`,
  };
}

function getOutlookMailbox(): OutlookMailboxContext | undefined {
  const officeGlobal = (
    globalThis as { Office?: { context?: { mailbox?: OutlookMailboxContext } } }
  ).Office;
  return officeGlobal?.context?.mailbox as OutlookMailboxContext | undefined;
}

export function canCreateOutlookMessageForRule(): boolean {
  return Boolean(getOutlookMailbox()?.displayNewMessageForm);
}

export function ProfileList() {
  const classes = useStyles();
  const profiles = useStore((s) => s.automationProfiles);
  const messages = useStore((s) => s.autoReplyMessages);
  const addProfile = useStore((s) => s.addProfile);
  const updateProfile = useStore((s) => s.updateProfile);
  const reorderProfiles = useStore((s) => s.reorderProfiles);
  const deleteProfile = useStore((s) => s.deleteProfile);

  const [editing, setEditing] = useState<AutomationProfile | null>(null);
  const [draggedProfileId, setDraggedProfileId] = useState<string | null>(null);
  const [isNew, setIsNew] = useState(false);
  const [draft, setDraft] = useState<ProfileDraft>(defaultProfile());
  const [canCreateOutlookMessage, setCanCreateOutlookMessage] = useState(() =>
    canCreateOutlookMessageForRule(),
  );
  const [teamsStatusTestState, setTeamsStatusTestState] = useState<
    "idle" | "testing" | "error"
  >("idle");
  const [teamsStatusTestError, setTeamsStatusTestError] = useState<
    string | null
  >(null);

  useEffect(() => {
    if (typeof Office === "undefined") {
      return;
    }

    let isDisposed = false;
    const syncComposeCapability = () => {
      if (!isDisposed) {
        setCanCreateOutlookMessage(canCreateOutlookMessageForRule());
      }
    };

    syncComposeCapability();

    void Office.onReady().then(
      () => {
        syncComposeCapability();
      },
      () => undefined,
    );

    return () => {
      isDisposed = true;
    };
  }, []);

  const openNew = () => {
    setDraft(defaultProfile());
    setIsNew(true);
    setEditing(null);
    setTeamsStatusTestState("idle");
    setTeamsStatusTestError(null);
  };
  const openEdit = (p: AutomationProfile) => {
    // eslint-disable-next-line @typescript-eslint/no-unused-vars
    const { id: _id, createdAt: _ca, updatedAt: _ua, ...rest } = p;
    setDraft(rest);
    setEditing(p);
    setIsNew(false);
    setTeamsStatusTestState("idle");
    setTeamsStatusTestError(null);
  };

  const handleSave = () => {
    if (isNew) addProfile(draft);
    else if (editing) updateProfile(editing.id, draft);
    setEditing(null);
    setIsNew(false);
  };

  const handleTestTeamsStatus = async () => {
    setTeamsStatusTestError(null);
    setTeamsStatusTestState("testing");
    try {
      const account = await getAccount();
      if (!account) {
        await signIn();
      }
      const graphClient = getGraphClient();
      await setTeamsPresence(
        graphClient,
        draft.teamsStatusSettings.statusWhenActive,
        draft.teamsStatusSettings.statusMessageWhenActive,
      );
      setTimeout(() => {
        void clearTeamsPresence(getGraphClient())
          .catch((err) => {
            setTeamsStatusTestState("error");
            setTeamsStatusTestError(
              err instanceof Error ? err.message : String(err),
            );
          })
          .then(() => {
            setTeamsStatusTestState((prev) =>
              prev === "error" ? prev : "idle",
            );
          });
      }, TEAMS_STATUS_TEST_DURATION_MS);
    } catch (err) {
      setTeamsStatusTestState("error");
      setTeamsStatusTestError(err instanceof Error ? err.message : String(err));
    }
  };

  const addKeyword = () => {
    const rule: KeywordRule = {
      id: nanoid(),
      field: "title",
      operator: "contains",
      value: "",
      caseSensitive: false,
    };
    setDraft((p) => ({
      ...p,
      matchRules: {
        ...p.matchRules,
        keywordRules: [...p.matchRules.keywordRules, rule],
      },
    }));
  };

  const updateKeyword = (id: string, changes: Partial<KeywordRule>) => {
    setDraft((p) => ({
      ...p,
      matchRules: {
        ...p.matchRules,
        keywordRules: p.matchRules.keywordRules.map((r) =>
          r.id === id ? { ...r, ...changes } : r,
        ),
      },
    }));
  };

  const removeKeyword = (id: string) => {
    setDraft((p) => ({
      ...p,
      matchRules: {
        ...p.matchRules,
        keywordRules: p.matchRules.keywordRules.filter((r) => r.id !== id),
      },
    }));
  };

  const createMessageForRule = (profile: AutomationProfile) => {
    const message = messages.find((m) => m.id === profile.autoReplyMessageId);
    if (!message) return;

    const mailbox = getOutlookMailbox();
    if (!mailbox?.displayNewMessageForm) return;

    mailbox.displayNewMessageForm(buildCopilotDraftForRule(profile, message));
  };

  const recommendation = getProfileRecommendation(draft);
  const applyRecommendation = () => {
    if (recommendation === "teamsOnly") {
      setDraft((profile) => ({
        ...profile,
        enableAutoReply: false,
        teamsStatusSettings: {
          ...profile.teamsStatusSettings,
          enabled: true,
        },
      }));
      return;
    }

    if (recommendation === "internalOnly") {
      setDraft((profile) => ({
        ...profile,
        autoReplyAudience: "internal",
      }));
      return;
    }

    if (recommendation === "minimumEightHours") {
      setDraft((profile) => ({
        ...profile,
        matchRules: {
          ...profile.matchRules,
          durationRule: {
            ...profile.matchRules.durationRule,
            enabled: true,
            minHours: 8,
          },
        },
      }));
    }
  };

  const isOpen = isNew || editing !== null;

  return (
    <div
      style={{
        display: "flex",
        flexDirection: "column",
        gap: tokens.spacingVerticalM,
      }}
    >
      <div
        style={{
          display: "flex",
          justifyContent: "space-between",
          alignItems: "center",
        }}
      >
        <span style={{ fontWeight: "600", fontSize: tokens.fontSizeBase400 }}>
          Automation Profiles
        </span>
        <Button
          icon={<Add24Regular />}
          appearance="primary"
          size="small"
          onClick={openNew}
        >
          New Profile
        </Button>
      </div>

      {profiles.length === 0 && (
        <p style={{ color: tokens.colorNeutralForeground3 }}>
          No profiles yet. Create a profile to define when auto-reply activates.
        </p>
      )}

      {profiles.map((p, index) => {
        const hasMessage = messages.some((m) => m.id === p.autoReplyMessageId);
        const canCreateMessage = hasMessage && canCreateOutlookMessage;
        const isDragging = draggedProfileId === p.id;

        return (
          <motion.div
            key={p.id}
            layout
            draggable
            onDragStartCapture={(event) => {
              setDraggedProfileId(p.id);
              event.dataTransfer.effectAllowed = "move";
              event.dataTransfer.setData("text/plain", p.id);
            }}
            onDragEndCapture={() => {
              setDraggedProfileId(null);
            }}
            onDragEnterCapture={() => {
              if (draggedProfileId && draggedProfileId !== p.id) {
                reorderProfiles(draggedProfileId, p.id);
              }
            }}
            onDragOverCapture={(event) => {
              event.preventDefault();
              event.dataTransfer.dropEffect = "move";
            }}
            onDropCapture={(event) => {
              event.preventDefault();
              setDraggedProfileId(null);
            }}
            transition={{
              type: "spring",
              stiffness: 260,
              damping: 22,
              mass: 0.8,
            }}
            style={{
              border: `1px solid ${tokens.colorNeutralStroke2}`,
              borderRadius: tokens.borderRadiusMedium,
              padding: tokens.spacingVerticalS,
              display: "flex",
              justifyContent: "space-between",
              alignItems: "center",
              cursor: "grab",
              opacity: isDragging ? 0.5 : 1,
            }}
          >
            <div
              style={{
                display: "flex",
                alignItems: "center",
                gap: tokens.spacingHorizontalXS,
              }}
            >
              <div
                aria-label="Drag to reorder profile"
                title="Drag to reorder"
                style={{
                  display: "flex",
                  alignItems: "center",
                  justifyContent: "center",
                  width: 18,
                  height: 18,
                  color: tokens.colorNeutralForeground3,
                  cursor: "grab",
                  userSelect: "none",
                  opacity: 0.8,
                  lineHeight: 1,
                  fontSize: 16,
                  letterSpacing: "0.06em",
                }}
              >
                ⋮⋮
              </div>
              <div>
                <span style={{ fontWeight: "600" }}>{p.name}</span>
                <span
                  style={{
                    marginLeft: 8,
                    color: tokens.colorNeutralForeground3,
                    fontSize: tokens.fontSizeBase200,
                  }}
                >
                  Order {index + 1} · {p.enabled ? "Enabled" : "Disabled"} ·{" "}
                  {getAutoReplySummary(p)}
                </span>
              </div>
            </div>
            <div style={{ display: "flex", gap: tokens.spacingHorizontalXS }}>
              <Switch
                checked={p.enabled}
                onChange={(_e, d) =>
                  updateProfile(p.id, { enabled: d.checked })
                }
              />
              <Button
                appearance="secondary"
                size="small"
                onClick={() => createMessageForRule(p)}
                disabled={!canCreateMessage}
              >
                Preview
              </Button>
              <Button
                icon={<Edit24Regular />}
                appearance="primary"
                size="small"
                onClick={() => openEdit(p)}
                style={{ minWidth: "56px" }}
              />
              <Button
                icon={<Delete24Regular />}
                appearance="subtle"
                size="small"
                onClick={() => deleteProfile(p.id)}
              />
            </div>
          </motion.div>
        );
      })}

      <Dialog
        open={isOpen}
        onOpenChange={(_e, data) => {
          if (!data.open) {
            setEditing(null);
            setIsNew(false);
          }
        }}
      >
        <DialogSurface className={classes.dialogSurface}>
          <DialogBody className={classes.dialogBody}>
            <DialogTitle>
              {isNew ? "New Automation Profile" : "Edit Automation Profile"}
            </DialogTitle>
            <DialogContent className={classes.dialogContent}>
              {/* Basic settings */}
              <Field
                className={classes.profileName}
                label="Profile name"
                required
              >
                <Input
                  value={draft.name}
                  onChange={(_e, d) =>
                    setDraft((p) => ({ ...p, name: d.value }))
                  }
                />
              </Field>
              <fieldset
                className={`${classes.section} ${classes.emailSection}`}
              >
                <legend className={classes.sectionTitle}>
                  Email auto-reply
                </legend>
                <div className={classes.sectionHeader}>
                  <Text size={200} className={classes.sectionDescription}>
                    Choose who receives an Outlook reply and when it stays
                    active.
                  </Text>
                </div>
                <div className={classes.sectionContent}>
                  <Switch
                    checked={draft.enableAutoReply !== false}
                    onChange={(_e, data) =>
                      setDraft((profile) => ({
                        ...profile,
                        enableAutoReply: data.checked,
                      }))
                    }
                    label="Enable Outlook auto-reply"
                  />
                  {draft.enableAutoReply !== false && (
                    <div className={classes.subsection}>
                      <Field label="Recipients">
                        <Dropdown
                          inlinePopup
                          value={
                            draft.autoReplyAudience === "internal"
                              ? "Internal users only"
                              : draft.autoReplyAudience === "external"
                                ? "External users only"
                                : "Internal and external users"
                          }
                          onOptionSelect={(_e, data) =>
                            setDraft((profile) => ({
                              ...profile,
                              autoReplyAudience:
                                data.optionValue as AutoReplyAudience,
                            }))
                          }
                        >
                          <Option value="internal">Internal users only</Option>
                          <Option value="external">External users only</Option>
                          <Option value="both">
                            Internal and external users
                          </Option>
                        </Dropdown>
                      </Field>
                      <Field label="Auto-reply message" required>
                        <Dropdown
                          inlinePopup
                          value={
                            messages.find(
                              (m) => m.id === draft.autoReplyMessageId,
                            )?.name ?? "Select…"
                          }
                          onOptionSelect={(_e, d) =>
                            setDraft((p) => ({
                              ...p,
                              autoReplyMessageId: d.optionValue as string,
                            }))
                          }
                        >
                          {messages.map((m) => (
                            <Option key={m.id} value={m.id}>
                              {m.name}
                            </Option>
                          ))}
                        </Dropdown>
                      </Field>
                    </div>
                  )}
                  <div className={classes.subsection}>
                    <div className={classes.subsectionHeader}>
                      <Text className={classes.subsectionTitle}>Timing</Text>
                      <Text size={200} className={classes.sectionDescription}>
                        Optional offsets around the appointment
                      </Text>
                    </div>
                    <Switch
                      checked={draft.timingSettings.enableBefore}
                      onChange={(_e, d) =>
                        setDraft((p) => ({
                          ...p,
                          timingSettings: {
                            ...p.timingSettings,
                            enableBefore: d.checked,
                          },
                        }))
                      }
                      label="Activate before appointment"
                    />
                    {draft.timingSettings.enableBefore && (
                      <Field label="Hours before appointment">
                        <Input
                          type="number"
                          min="0"
                          step="0.25"
                          value={String(
                            draft.timingSettings.hoursBeforeAppointment,
                          )}
                          onChange={(_e, d) =>
                            setDraft((p) => ({
                              ...p,
                              timingSettings: {
                                ...p.timingSettings,
                                hoursBeforeAppointment: Number(d.value),
                              },
                            }))
                          }
                        />
                      </Field>
                    )}
                    <Switch
                      checked={draft.timingSettings.enableAfter}
                      onChange={(_e, d) =>
                        setDraft((p) => ({
                          ...p,
                          timingSettings: {
                            ...p.timingSettings,
                            enableAfter: d.checked,
                          },
                        }))
                      }
                      label="Keep active after appointment"
                    />
                    {draft.timingSettings.enableAfter && (
                      <Field label="Hours after appointment">
                        <Input
                          type="number"
                          min="0"
                          step="0.25"
                          value={String(
                            draft.timingSettings.hoursAfterAppointment,
                          )}
                          onChange={(_e, d) =>
                            setDraft((p) => ({
                              ...p,
                              timingSettings: {
                                ...p.timingSettings,
                                hoursAfterAppointment: Number(d.value),
                              },
                            }))
                          }
                        />
                      </Field>
                    )}
                  </div>
                </div>
              </fieldset>

              {/* Keyword rules */}
              <fieldset
                className={`${classes.section} ${classes.rulesSection}`}
              >
                <legend className={classes.sectionTitle}>
                  Appointment matching rules
                </legend>
                <div className={classes.sectionHeader}>
                  <Text size={200} className={classes.sectionDescription}>
                    Define which calendar appointments activate this profile.
                  </Text>
                </div>
                <div className={classes.sectionContent}>
                  <Field label="Combine rules with">
                    <Dropdown
                      inlinePopup
                      value={draft.matchRules.combinator}
                      onOptionSelect={(_e, d) =>
                        setDraft((p) => ({
                          ...p,
                          matchRules: {
                            ...p.matchRules,
                            combinator: d.optionValue as "AND" | "OR",
                          },
                        }))
                      }
                    >
                      <Option value="AND">AND (all must match)</Option>
                      <Option value="OR">OR (any must match)</Option>
                    </Dropdown>
                  </Field>
                  <div className={classes.keywordList}>
                    {draft.matchRules.keywordRules.map((rule) => (
                      <div key={rule.id} className={classes.keywordRow}>
                      <Field label="Field" className={classes.keywordField}>
                        <Combobox
                          inlinePopup
                          value={FIELD_LABELS[rule.field]}
                          onOptionSelect={(_e, d) =>
                            updateKeyword(rule.id, {
                              field: d.optionValue as AppointmentMatchField,
                            })
                          }
                        >
                          {FIELDS.map((f) => (
                            <Option key={f} value={f}>
                              {FIELD_LABELS[f]}
                            </Option>
                          ))}
                        </Combobox>
                      </Field>
                      <Field label="Operator" className={classes.keywordField}>
                        <Combobox
                          inlinePopup
                          value={OPERATOR_LABELS[rule.operator]}
                          onOptionSelect={(_e, d) =>
                            updateKeyword(rule.id, {
                              operator:
                                d.optionValue as AppointmentMatchOperator,
                            })
                          }
                        >
                          {OPERATORS.map((o) => (
                            <Option key={o} value={o}>
                              {OPERATOR_LABELS[o]}
                            </Option>
                          ))}
                        </Combobox>
                      </Field>
                      <Field label="Value" className={classes.keywordValue}>
                        <Input
                          value={rule.value}
                          placeholder={
                            rule.operator === "regex"
                              ? "e.g. /vacation|holiday/i"
                              : "Enter a value"
                          }
                          onChange={(_e, d) =>
                            updateKeyword(rule.id, { value: d.value })
                          }
                        />
                      </Field>
                      <Button
                        icon={<Delete24Regular />}
                        appearance="subtle"
                        size="small"
                        className={classes.deleteRuleButton}
                        aria-label="Delete keyword rule"
                        title="Delete keyword rule"
                        onClick={() => removeKeyword(rule.id)}
                      />
                      </div>
                    ))}
                  </div>
                  <Button
                    icon={<Add24Regular />}
                    appearance="outline"
                    size="small"
                    className={classes.addRuleButton}
                    onClick={addKeyword}
                  >
                    Add keyword rule
                  </Button>

                  {/* Duration rule */}
                  <div className={classes.subsection}>
                    <Switch
                      checked={draft.matchRules.durationRule.enabled}
                      onChange={(_e, d) =>
                        setDraft((p) => ({
                          ...p,
                          matchRules: {
                            ...p.matchRules,
                            durationRule: {
                              ...p.matchRules.durationRule,
                              enabled: d.checked,
                            },
                          },
                        }))
                      }
                      label="Filter by duration"
                    />
                    {draft.matchRules.durationRule.enabled && (
                      <div className={classes.twoColumnGrid}>
                      <Field label="Minimum hours">
                        <Input
                          type="number"
                          min="0"
                          value={String(
                            draft.matchRules.durationRule.minHours ?? "",
                          )}
                          onChange={(_e, d) =>
                            setDraft((p) => ({
                              ...p,
                              matchRules: {
                                ...p.matchRules,
                                durationRule: {
                                  ...p.matchRules.durationRule,
                                  minHours: d.value
                                    ? Number(d.value)
                                    : undefined,
                                },
                              },
                            }))
                          }
                        />
                      </Field>
                      <Field label="Maximum hours">
                        <Input
                          type="number"
                          min="0"
                          value={String(
                            draft.matchRules.durationRule.maxHours ?? "",
                          )}
                          onChange={(_e, d) =>
                            setDraft((p) => ({
                              ...p,
                              matchRules: {
                                ...p.matchRules,
                                durationRule: {
                                  ...p.matchRules.durationRule,
                                  maxHours: d.value
                                    ? Number(d.value)
                                    : undefined,
                                },
                              },
                            }))
                          }
                        />
                      </Field>
                      </div>
                    )}
                  </div>

                  {/* Busy status */}
                  <div className={classes.subsection}>
                    <Switch
                      checked={draft.matchRules.busyStatusRule.enabled}
                      onChange={(_e, d) =>
                        setDraft((p) => ({
                          ...p,
                          matchRules: {
                            ...p.matchRules,
                            busyStatusRule: {
                              ...p.matchRules.busyStatusRule,
                              enabled: d.checked,
                            },
                          },
                        }))
                      }
                      label="Filter by busy status"
                    />
                    {draft.matchRules.busyStatusRule.enabled && (
                      <div className={classes.statusGrid}>
                      {BUSY_STATUSES.map((s) => (
                        <Checkbox
                          key={s}
                          label={BUSY_STATUS_LABELS[s]}
                          checked={draft.matchRules.busyStatusRule.statuses.includes(
                            s,
                          )}
                          onChange={(_e, d) => {
                            const statuses = d.checked
                              ? [...draft.matchRules.busyStatusRule.statuses, s]
                              : draft.matchRules.busyStatusRule.statuses.filter(
                                  (x) => x !== s,
                                );
                            setDraft((p) => ({
                              ...p,
                              matchRules: {
                                ...p.matchRules,
                                busyStatusRule: {
                                  ...p.matchRules.busyStatusRule,
                                  statuses,
                                },
                              },
                            }));
                          }}
                        />
                      ))}
                      </div>
                    )}
                  </div>
                  <Text size={200} className={classes.helperText}>
                    Matching order is taken from the list order. Drag a profile
                    by its grip to change the precedence.
                  </Text>
                  {recommendation && (
                    <div
                      role="status"
                      className={classes.recommendation}
                    >
                      <Text>
                        {recommendation === "teamsOnly"
                          ? "No duration filter is set. Consider using only a Teams status for this rule."
                          : recommendation === "internalOnly"
                            ? "This rule targets appointments shorter than 8 hours. Consider replying to internal users only."
                            : "External auto-replies are best reserved for appointments lasting 8 hours or longer."}
                      </Text>
                      <Button
                        appearance="secondary"
                        size="small"
                        onClick={applyRecommendation}
                      >
                        {recommendation === "teamsOnly"
                          ? "Use Teams only"
                          : recommendation === "internalOnly"
                            ? "Internal only"
                            : "Set 8-hour minimum"}
                      </Button>
                    </div>
                  )}
                </div>
              </fieldset>

              {/* Teams status */}
              <fieldset
                className={`${classes.section} ${classes.teamsSection}`}
              >
                <legend className={classes.sectionTitle}>Teams status</legend>
                <div className={classes.sectionHeader}>
                  <Text size={200} className={classes.sectionDescription}>
                    Keep your Teams presence aligned with this calendar rule.
                  </Text>
                </div>
                <div className={classes.sectionContent}>
                  <div className={classes.subsectionHeader}>
                    <Switch
                      checked={draft.teamsStatusSettings.enabled}
                      onChange={(_e, d) =>
                        setDraft((p) => ({
                          ...p,
                          teamsStatusSettings: {
                            ...p.teamsStatusSettings,
                            enabled: d.checked,
                          },
                        }))
                      }
                      label="Set Teams status when active"
                    />
                    {draft.teamsStatusSettings.enabled && (
                      <Button
                        appearance="primary"
                        size="small"
                        onClick={() => void handleTestTeamsStatus()}
                        disabled={teamsStatusTestState === "testing"}
                      >
                        {teamsStatusTestState === "testing"
                          ? "Testing…"
                          : "Test Teams status"}
                      </Button>
                    )}
                    {teamsStatusTestState === "testing" && (
                      <Text size={200} className={classes.helperText}>
                        Status set to {draft.teamsStatusSettings.statusWhenActive}
                        {" – reverting in 3 seconds…"}
                      </Text>
                    )}
                    {teamsStatusTestState === "error" && (
                      <span
                        style={{
                          fontSize: tokens.fontSizeBase200,
                          color: tokens.colorPaletteRedForeground1,
                        }}
                      >
                        Test failed: {teamsStatusTestError}
                      </span>
                    )}
                  </div>
                  {draft.teamsStatusSettings.enabled && (
                    <div className={classes.subsection}>
                      <Field label="Teams status">
                        <Dropdown
                          inlinePopup
                          value={draft.teamsStatusSettings.statusWhenActive}
                          onOptionSelect={(_e, d) =>
                            setDraft((p) => ({
                              ...p,
                              teamsStatusSettings: {
                                ...p.teamsStatusSettings,
                                statusWhenActive: d.optionValue as never,
                              },
                            }))
                          }
                        >
                          {(
                            [
                              "Available",
                              "Busy",
                              "DoNotDisturb",
                              "BeRightBack",
                              "Away",
                              "Offline",
                            ] as const
                          ).map((s) => (
                            <Option key={s} value={s}>
                              {s}
                            </Option>
                          ))}
                        </Dropdown>
                      </Field>
                      <Field label="Status message">
                        <Input
                          value={
                            draft.teamsStatusSettings.statusMessageWhenActive
                          }
                          onChange={(_e, d) =>
                            setDraft((p) => ({
                              ...p,
                              teamsStatusSettings: {
                                ...p.teamsStatusSettings,
                                statusMessageWhenActive: d.value,
                              },
                            }))
                          }
                          placeholder="Optional status message…"
                        />
                      </Field>
                      <Switch
                        checked={draft.teamsStatusSettings.restoreOnEnd}
                        onChange={(_e, d) =>
                          setDraft((p) => ({
                            ...p,
                            teamsStatusSettings: {
                              ...p.teamsStatusSettings,
                              restoreOnEnd: d.checked,
                            },
                          }))
                        }
                        label="Restore Teams status when auto-reply ends"
                      />
                    </div>
                  )}
                </div>
              </fieldset>
            </DialogContent>

            <DialogActions className={classes.dialogActions}>
              <DialogTrigger disableButtonEnhancement>
                <Button appearance="secondary">Cancel</Button>
              </DialogTrigger>
              <Button
                appearance="primary"
                onClick={handleSave}
                disabled={
                  !draft.name.trim() ||
                  (draft.enableAutoReply !== false && !draft.autoReplyMessageId)
                }
              >
                Save
              </Button>
            </DialogActions>
          </DialogBody>
        </DialogSurface>
      </Dialog>
    </div>
  );
}
