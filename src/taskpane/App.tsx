import { useEffect, useRef, useState } from "react";
import type { ChangeEvent } from "react";
import {
  Button,
  Dialog,
  DialogActions,
  DialogBody,
  DialogContent,
  DialogSurface,
  DialogTitle,
  FluentProvider,
  type Theme,
  Title3,
  makeStyles,
  shorthands,
  tokens,
  webDarkTheme,
  webLightTheme,
} from "@fluentui/react-components";
import { Info24Regular } from "@fluentui/react-icons";
import { AutoReplyList } from "./components/AutoReplyList";
import { ProfileList } from "./components/ProfileList";
import { useStore } from "./useStore";

const useStyles = makeStyles({
  page: {
    minHeight: "100vh",
    backgroundColor: tokens.colorNeutralBackground2,
  },
  shell: {
    maxWidth: "960px",
    marginLeft: "auto",
    marginRight: "auto",
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalXL,
    ...shorthands.padding(
      tokens.spacingVerticalXXL,
      tokens.spacingHorizontalXXL,
    ),
  },
  header: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalXS,
  },
  headerTopRow: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    gap: tokens.spacingHorizontalM,
    flexWrap: "wrap",
  },
  headerActions: {
    display: "flex",
    alignItems: "center",
    gap: tokens.spacingHorizontalS,
  },
  subtitle: {
    color: tokens.colorNeutralForeground3,
    fontSize: tokens.fontSizeBase300,
    lineHeight: tokens.lineHeightBase300,
    margin: 0,
  },
  status: {
    margin: 0,
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground2,
  },
  section: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalM,
    backgroundColor: tokens.colorNeutralBackground1,
    ...shorthands.padding(tokens.spacingVerticalL, tokens.spacingHorizontalL),
    ...shorthands.borderRadius(tokens.borderRadiusLarge),
    boxShadow: tokens.shadow4,
  },
});

function parseHexChannel(value: string): number {
  return Number.parseInt(value, 16);
}

function getHexColorLuminance(hexColor: string): number | null {
  const normalized = hexColor.trim();
  const match = /^#([0-9a-f]{3}|[0-9a-f]{6})$/i.exec(normalized);
  if (!match) {
    return null;
  }

  const rawHex = match[1].toLowerCase();
  const expandedHex =
    rawHex.length === 3
      ? `${rawHex[0]}${rawHex[0]}${rawHex[1]}${rawHex[1]}${rawHex[2]}${rawHex[2]}`
      : rawHex;

  const red = parseHexChannel(expandedHex.slice(0, 2));
  const green = parseHexChannel(expandedHex.slice(2, 4));
  const blue = parseHexChannel(expandedHex.slice(4, 6));

  return (red * 299 + green * 587 + blue * 114) / 1000;
}

function resolveThemeFromOfficeTheme(officeTheme?: Office.OfficeTheme): Theme {
  const bodyBackground = officeTheme?.bodyBackgroundColor;
  const controlBackground = officeTheme?.controlBackgroundColor;
  const luminance = bodyBackground
    ? getHexColorLuminance(bodyBackground)
    : controlBackground
      ? getHexColorLuminance(controlBackground)
      : null;

  if (luminance === null) {
    return webLightTheme;
  }

  return luminance < 128 ? webDarkTheme : webLightTheme;
}

function getInitialTheme(): Theme {
  if (typeof Office === "undefined") {
    return webLightTheme;
  }

  return resolveThemeFromOfficeTheme(Office.context?.officeTheme);
}

export function App() {
  const classes = useStyles();
  const loadFromStorage = useStore((state) => state.loadFromStorage);
  const exportSettings = useStore((state) => state.exportSettings);
  const importSettings = useStore((state) => state.importSettings);
  const profiles = useStore((state) => state.automationProfiles);
  const hasHandledDefaultRulesPrompt = useStore(
    (state) => state.hasHandledDefaultRulesPrompt,
  );
  const handleDefaultRulesPrompt = useStore(
    (state) => state.handleDefaultRulesPrompt,
  );
  const importInputRef = useRef<HTMLInputElement | null>(null);
  const [statusMessage, setStatusMessage] = useState<string>("");
  const [theme, setTheme] = useState<Theme>(() => getInitialTheme());
  const [isHelpOpen, setIsHelpOpen] = useState(false);
  const showDefaultRulesPrompt =
    !hasHandledDefaultRulesPrompt && profiles.length === 0;

  useEffect(() => {
    loadFromStorage();
  }, [loadFromStorage]);

  useEffect(() => {
    if (typeof Office === "undefined") {
      return;
    }

    let isDisposed = false;
    const applyTheme = (officeTheme?: Office.OfficeTheme) => {
      if (isDisposed) {
        return;
      }
      const nextTheme = resolveThemeFromOfficeTheme(
        officeTheme ?? Office.context?.officeTheme,
      );
      setTheme(nextTheme);
    };

    const supportsThemeEvent =
      Office.context?.requirements?.isSetSupported("Mailbox", "1.14") ?? false;
    const onOfficeThemeChanged = (
      event: Office.OfficeThemeChangedEventArgs,
    ) => {
      applyTheme(event.officeTheme);
    };

    void Office.onReady().then(() => {
      applyTheme();

      if (!supportsThemeEvent) {
        return;
      }

      Office.context?.mailbox?.addHandlerAsync(
        Office.EventType.OfficeThemeChanged,
        onOfficeThemeChanged,
      );
    });

    return () => {
      isDisposed = true;

      if (!supportsThemeEvent) {
        return;
      }

      Office.context?.mailbox?.removeHandlerAsync(
        Office.EventType.OfficeThemeChanged,
      );
    };
  }, []);

  const handleExport = () => {
    const json = exportSettings();
    const blob = new Blob([json], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    const timestamp = new Date().toISOString().replace(/[:.]/g, "-");
    anchor.href = url;
    anchor.download = `outlook-auto-reply-settings-${timestamp}.json`;
    document.body.appendChild(anchor);
    anchor.click();
    document.body.removeChild(anchor);
    URL.revokeObjectURL(url);
    setStatusMessage("Settings exported.");
  };

  const handleImportFile = async (event: ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    const fileContents = await file.text();
    const result = importSettings(fileContents);

    if (result.success) {
      setStatusMessage("Settings imported successfully.");
    } else {
      setStatusMessage(`Import failed: ${result.error}`);
    }

    event.target.value = "";
  };

  return (
    <FluentProvider theme={theme} className={classes.page}>
      <Dialog open={showDefaultRulesPrompt} modalType="modal">
        <DialogSurface>
          <DialogBody>
            <DialogTitle>Standardregeln erstellen?</DialogTitle>
            <DialogContent>
              Create five English example messages and rules for vacation,
              meetings, public holidays, training, and traveling?
            </DialogContent>
            <DialogActions>
              <Button
                appearance="secondary"
                onClick={() => handleDefaultRulesPrompt(false)}
              >
                No, create my own
              </Button>
              <Button
                appearance="primary"
                onClick={() => handleDefaultRulesPrompt(true)}
              >
                Create default rules
              </Button>
            </DialogActions>
          </DialogBody>
        </DialogSurface>
      </Dialog>
      <main className={classes.shell}>
        <header className={classes.header}>
          <div className={classes.headerTopRow}>
            <Title3>Outlook Auto-Reply Automater</Title3>
            <div className={classes.headerActions}>
              <Button
                appearance="subtle"
                icon={<Info24Regular />}
                aria-label="Help with placeholders and rules"
                title="Help with placeholders and rules"
                onClick={() => setIsHelpOpen(true)}
              />
              <Button
                appearance="secondary"
                size="small"
                onClick={() => {
                  handleDefaultRulesPrompt(true);
                  setStatusMessage("Default rules created.");
                }}
              >
                Create Default Rules
              </Button>
              <Button
                appearance="secondary"
                size="small"
                onClick={handleExport}
              >
                Export Settings
              </Button>
              <Button
                appearance="primary"
                size="small"
                onClick={() => importInputRef.current?.click()}
              >
                Import Settings
              </Button>
              <input
                ref={importInputRef}
                type="file"
                accept="application/json"
                onChange={handleImportFile}
                style={{ display: "none" }}
              />
            </div>
          </div>
          <p className={classes.subtitle}>
            Manage reusable auto-reply messages and the rules that activate
            them.
          </p>
          {statusMessage ? (
            <p className={classes.status}>{statusMessage}</p>
          ) : null}
        </header>

        <section className={classes.section}>
          <AutoReplyList />
        </section>

        <section className={classes.section}>
          <ProfileList />
        </section>
      </main>
      <Dialog
        open={isHelpOpen}
        onOpenChange={(_event, data) => setIsHelpOpen(data.open)}
      >
        <DialogSurface style={{ maxWidth: "720px", width: "100%" }}>
          <DialogBody>
            <DialogTitle>Placeholders and Rules</DialogTitle>
            <DialogContent
              style={{
                display: "flex",
                flexDirection: "column",
                gap: tokens.spacingVerticalM,
              }}
            >
              <div>
                <strong>Message placeholders</strong>
                <ul>
                  <li>
                    <code>{"{{appointment.title}}"}</code>: Appointment subject
                  </li>
                  <li>
                    <code>{"{{appointment.start}}"}</code>: Start date
                  </li>
                  <li>
                    <code>{"{{appointment.end}}"}</code>: End date
                  </li>
                  <li>
                    <code>{"{{appointment.nextWorkingDayAfterEnd}}"}</code>:
                    Next working day after the appointment ends, skipping
                    Saturday and Sunday
                  </li>
                  <li>
                    <code>{"{{appointment.location}}"}</code>: Appointment
                    location
                  </li>
                  <li>
                    <code>{"{{rule.match}}"}</code>: First capture group from a
                    matching regex, or the complete first regex match
                  </li>
                </ul>
              </div>
              <div>
                <strong>Appointment rules</strong>
                <ul>
                  <li>
                    Profile order: the first matching profile in the list wins.
                  </li>
                  <li>
                    Auto-reply: Enables or disables Outlook out-of-office for
                    the profile. Teams can be configured independently.
                  </li>
                  <li>
                    Timing: Activate before an appointment and remain active
                    after it ends, in hours.
                  </li>
                  <li>
                    Combinator: <code>AND</code> requires all rules;{" "}
                    <code>OR</code> requires any rule.
                  </li>
                  <li>
                    Fields: Title, location, and category can be matched. Body
                    and organizer are available in the editor, but the current
                    calendar connection does not yet provide them.
                  </li>
                  <li>
                    Operators: contains, starts with, ends with, equals, and
                    regex. Regex can be entered as <code>/pattern/i</code>.
                  </li>
                  <li>Duration: minimum and maximum duration in hours.</li>
                  <li>
                    Busy status: free, tentative, busy, out of office, or
                    working elsewhere.
                  </li>
                  <li>
                    Teams: Status, status message, and restoring the prior
                    status after the appointment.
                  </li>
                </ul>
              </div>
            </DialogContent>
            <DialogActions>
              <Button appearance="primary" onClick={() => setIsHelpOpen(false)}>
                Close
              </Button>
            </DialogActions>
          </DialogBody>
        </DialogSurface>
      </Dialog>
    </FluentProvider>
  );
}
