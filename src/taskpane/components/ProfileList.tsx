import { useState } from "react";
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
  Option,
  Switch,
  tokens,
} from "@fluentui/react-components";
import {
  Add24Regular,
  Delete24Regular,
  Edit24Regular,
} from "@fluentui/react-icons";
import type {
  AutomationProfile,
  AppointmentBusyStatus,
  AppointmentMatchField,
  AppointmentMatchOperator,
  KeywordRule,
} from "../types";
import { useStore } from "../useStore";
import { nanoid } from "../utils/nanoid";

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

type ProfileDraft = Omit<AutomationProfile, "id" | "createdAt" | "updatedAt">;

function defaultProfile(): ProfileDraft {
  return {
    name: "",
    enabled: true,
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

export function ProfileList() {
  const profiles = useStore((s) => s.automationProfiles);
  const messages = useStore((s) => s.autoReplyMessages);
  const addProfile = useStore((s) => s.addProfile);
  const updateProfile = useStore((s) => s.updateProfile);
  const deleteProfile = useStore((s) => s.deleteProfile);

  const [editing, setEditing] = useState<AutomationProfile | null>(null);
  const [isNew, setIsNew] = useState(false);
  const [draft, setDraft] = useState<ProfileDraft>(defaultProfile());

  const openNew = () => {
    setDraft(defaultProfile());
    setIsNew(true);
    setEditing(null);
  };
  const openEdit = (p: AutomationProfile) => {
    // eslint-disable-next-line @typescript-eslint/no-unused-vars
    const { id: _id, createdAt: _ca, updatedAt: _ua, ...rest } = p;
    setDraft(rest);
    setEditing(p);
    setIsNew(false);
  };

  const handleSave = () => {
    if (isNew) addProfile(draft);
    else if (editing) updateProfile(editing.id, draft);
    setEditing(null);
    setIsNew(false);
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

      {profiles.map((p) => (
        <div
          key={p.id}
          style={{
            border: `1px solid ${tokens.colorNeutralStroke2}`,
            borderRadius: tokens.borderRadiusMedium,
            padding: tokens.spacingVerticalS,
            display: "flex",
            justifyContent: "space-between",
            alignItems: "center",
          }}
        >
          <div>
            <span style={{ fontWeight: "600" }}>{p.name}</span>
            <span
              style={{
                marginLeft: 8,
                color: tokens.colorNeutralForeground3,
                fontSize: tokens.fontSizeBase200,
              }}
            >
              Priority {p.priority} · {p.enabled ? "Enabled" : "Disabled"}
            </span>
          </div>
          <div style={{ display: "flex", gap: tokens.spacingHorizontalXS }}>
            <Switch
              checked={p.enabled}
              onChange={(_e, d) => updateProfile(p.id, { enabled: d.checked })}
            />
            <Button
              icon={<Edit24Regular />}
              appearance="subtle"
              size="small"
              onClick={() => openEdit(p)}
            />
            <Button
              icon={<Delete24Regular />}
              appearance="subtle"
              size="small"
              onClick={() => deleteProfile(p.id)}
            />
          </div>
        </div>
      ))}

      <Dialog
        open={isOpen}
        onOpenChange={(_e, data) => {
          if (!data.open) {
            setEditing(null);
            setIsNew(false);
          }
        }}
      >
        <DialogSurface style={{ maxWidth: "600px", width: "100%" }}>
          <DialogBody>
            <DialogTitle>
              {isNew ? "New Automation Profile" : "Edit Automation Profile"}
            </DialogTitle>
            <DialogContent
              style={{
                display: "flex",
                flexDirection: "column",
                gap: tokens.spacingVerticalM,
              }}
            >
              {/* Basic settings */}
              <Field label="Profile name" required>
                <Input
                  value={draft.name}
                  onChange={(_e, d) =>
                    setDraft((p) => ({ ...p, name: d.value }))
                  }
                />
              </Field>
              <Field label="Auto-reply message" required>
                <Dropdown
                  value={
                    messages.find((m) => m.id === draft.autoReplyMessageId)
                      ?.name ?? "Select…"
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
              <Field label="Priority (lower = higher priority)">
                <Input
                  type="number"
                  value={String(draft.priority)}
                  onChange={(_e, d) =>
                    setDraft((p) => ({ ...p, priority: Number(d.value) }))
                  }
                />
              </Field>

              {/* Timing */}
              <fieldset
                style={{
                  border: `1px solid ${tokens.colorNeutralStroke2}`,
                  borderRadius: tokens.borderRadiusMedium,
                  padding: tokens.spacingVerticalS,
                }}
              >
                <legend style={{ fontWeight: "600" }}>Timing</legend>
                <div
                  style={{
                    display: "flex",
                    flexDirection: "column",
                    gap: tokens.spacingVerticalS,
                  }}
                >
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
              </fieldset>

              {/* Keyword rules */}
              <fieldset
                style={{
                  border: `1px solid ${tokens.colorNeutralStroke2}`,
                  borderRadius: tokens.borderRadiusMedium,
                  padding: tokens.spacingVerticalS,
                }}
              >
                <legend style={{ fontWeight: "600" }}>
                  Appointment matching rules
                </legend>
                <div
                  style={{
                    display: "flex",
                    flexDirection: "column",
                    gap: tokens.spacingVerticalS,
                  }}
                >
                  <Field label="Combine rules with">
                    <Dropdown
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
                  {draft.matchRules.keywordRules.map((rule) => (
                    <div
                      key={rule.id}
                      style={{
                        display: "flex",
                        gap: tokens.spacingHorizontalXS,
                        alignItems: "flex-end",
                      }}
                    >
                      <Field label="Field" style={{ flex: 1 }}>
                        <Combobox
                          value={rule.field}
                          onOptionSelect={(_e, d) =>
                            updateKeyword(rule.id, {
                              field: d.optionValue as AppointmentMatchField,
                            })
                          }
                        >
                          {FIELDS.map((f) => (
                            <Option key={f} value={f}>
                              {f}
                            </Option>
                          ))}
                        </Combobox>
                      </Field>
                      <Field label="Operator" style={{ flex: 1 }}>
                        <Combobox
                          value={rule.operator}
                          onOptionSelect={(_e, d) =>
                            updateKeyword(rule.id, {
                              operator:
                                d.optionValue as AppointmentMatchOperator,
                            })
                          }
                        >
                          {OPERATORS.map((o) => (
                            <Option key={o} value={o}>
                              {o}
                            </Option>
                          ))}
                        </Combobox>
                      </Field>
                      <Field label="Value" style={{ flex: 2 }}>
                        <Input
                          value={rule.value}
                          onChange={(_e, d) =>
                            updateKeyword(rule.id, { value: d.value })
                          }
                        />
                      </Field>
                      <Button
                        icon={<Delete24Regular />}
                        appearance="subtle"
                        size="small"
                        onClick={() => removeKeyword(rule.id)}
                      />
                    </div>
                  ))}
                  <Button
                    icon={<Add24Regular />}
                    appearance="outline"
                    size="small"
                    onClick={addKeyword}
                  >
                    Add keyword rule
                  </Button>

                  {/* Duration rule */}
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
                    <div
                      style={{
                        display: "flex",
                        gap: tokens.spacingHorizontalM,
                      }}
                    >
                      <Field label="Min duration (min)">
                        <Input
                          type="number"
                          min="0"
                          value={String(
                            draft.matchRules.durationRule.minMinutes ?? "",
                          )}
                          onChange={(_e, d) =>
                            setDraft((p) => ({
                              ...p,
                              matchRules: {
                                ...p.matchRules,
                                durationRule: {
                                  ...p.matchRules.durationRule,
                                  minMinutes: d.value
                                    ? Number(d.value)
                                    : undefined,
                                },
                              },
                            }))
                          }
                        />
                      </Field>
                      <Field label="Max duration (min)">
                        <Input
                          type="number"
                          min="0"
                          value={String(
                            draft.matchRules.durationRule.maxMinutes ?? "",
                          )}
                          onChange={(_e, d) =>
                            setDraft((p) => ({
                              ...p,
                              matchRules: {
                                ...p.matchRules,
                                durationRule: {
                                  ...p.matchRules.durationRule,
                                  maxMinutes: d.value
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

                  {/* Busy status */}
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
                    <div
                      style={{
                        display: "flex",
                        flexWrap: "wrap",
                        gap: tokens.spacingHorizontalS,
                      }}
                    >
                      {BUSY_STATUSES.map((s) => (
                        <Checkbox
                          key={s}
                          label={s}
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
              </fieldset>

              {/* Teams status */}
              <fieldset
                style={{
                  border: `1px solid ${tokens.colorNeutralStroke2}`,
                  borderRadius: tokens.borderRadiusMedium,
                  padding: tokens.spacingVerticalS,
                }}
              >
                <legend style={{ fontWeight: "600" }}>Teams status</legend>
                <div
                  style={{
                    display: "flex",
                    flexDirection: "column",
                    gap: tokens.spacingVerticalS,
                  }}
                >
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
                    <>
                      <Field label="Teams status">
                        <Dropdown
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
                    </>
                  )}
                </div>
              </fieldset>
            </DialogContent>

            <DialogActions>
              <DialogTrigger disableButtonEnhancement>
                <Button appearance="secondary">Cancel</Button>
              </DialogTrigger>
              <Button
                appearance="primary"
                onClick={handleSave}
                disabled={!draft.name.trim() || !draft.autoReplyMessageId}
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
