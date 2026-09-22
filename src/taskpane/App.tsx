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
  Tooltip,
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
              Möchten Sie Nachrichten und Regeln für Reise, Training ab vier
              Stunden und Termine außer Haus beim Kunden erstellen?
            </DialogContent>
            <DialogActions>
              <Button
                appearance="secondary"
                onClick={() => handleDefaultRulesPrompt(false)}
              >
                Nein, selbst erstellen
              </Button>
              <Button
                appearance="primary"
                onClick={() => handleDefaultRulesPrompt(true)}
              >
                Standardregeln erstellen
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
              <Tooltip content="Hilfe zu Platzhaltern und Regeln" relationship="label">
                <Button
                  appearance="subtle"
                  icon={<Info24Regular />}
                  aria-label="Hilfe zu Platzhaltern und Regeln"
                  onClick={() => setIsHelpOpen(true)}
                />
              </Tooltip>
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
      <Dialog open={isHelpOpen} onOpenChange={(_event, data) => setIsHelpOpen(data.open)}>
        <DialogSurface style={{ maxWidth: "720px", width: "100%" }}>
          <DialogBody>
            <DialogTitle>Platzhalter und Regeln</DialogTitle>
            <DialogContent
              style={{
                display: "flex",
                flexDirection: "column",
                gap: tokens.spacingVerticalM,
              }}
            >
              <div>
                <strong>Nachrichten-Platzhalter</strong>
                <ul>
                  <li><code>{"{{appointment.title}}"}</code>: Betreff des Termins</li>
                  <li><code>{"{{appointment.start}}"}</code>: Startdatum</li>
                  <li><code>{"{{appointment.end}}"}</code>: Enddatum</li>
                  <li><code>{"{{appointment.nextWorkingDayAfterEnd}}"}</code>: Nächster Arbeitstag nach dem Terminende, ohne Samstag und Sonntag</li>
                  <li><code>{"{{appointment.location}}"}</code>: Ort des Termins</li>
                  <li><code>{"{{rule.match}}"}</code>: Erste Capture Group eines passenden Regex, sonst der vollständige erste Regex-Treffer</li>
                </ul>
              </div>
              <div>
                <strong>Terminregeln</strong>
                <ul>
                  <li>Profilpriorität: Bei mehreren Treffern gewinnt die kleinste Zahl.</li>
                  <li>Auto-Reply: Aktiviert oder deaktiviert die Outlook-Abwesenheitsnotiz für das Profil. Teams kann unabhängig davon gesetzt werden.</li>
                  <li>Zeitsteuerung: Aktivierung vor dem Termin und Fortsetzung nach seinem Ende, jeweils in Stunden.</li>
                  <li>Verknüpfung: <code>AND</code> verlangt alle Regeln, <code>OR</code> mindestens eine Regel.</li>
                  <li>Felder: Titel, Ort und Kategorie können abgeglichen werden. Body und Organisator stehen in der Oberfläche bereit, werden von der aktuellen Kalenderanbindung aber noch nicht geliefert.</li>
                  <li>Operatoren: enthält, beginnt mit, endet mit, exakt gleich sowie Regex. Regex kann als <code>/Muster/i</code> angegeben werden.</li>
                  <li>Dauer: Mindest- und Höchstdauer in Minuten.</li>
                  <li>Belegt-Status: Frei, vorläufig, gebucht, abwesend oder an anderem Ort tätig.</li>
                  <li>Teams: Status, Statusnachricht und Wiederherstellen des vorherigen Status nach dem Termin.</li>
                </ul>
              </div>
            </DialogContent>
            <DialogActions>
              <Button appearance="primary" onClick={() => setIsHelpOpen(false)}>
                Schließen
              </Button>
            </DialogActions>
          </DialogBody>
        </DialogSurface>
      </Dialog>
    </FluentProvider>
  );
}
