import { useEffect, useRef, useState } from "react";
import type { ChangeEvent } from "react";
import {
  Button,
  FluentProvider,
  Title3,
  makeStyles,
  shorthands,
  tokens,
  webLightTheme,
} from "@fluentui/react-components";
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

export function App() {
  const classes = useStyles();
  const loadFromStorage = useStore((state) => state.loadFromStorage);
  const exportSettings = useStore((state) => state.exportSettings);
  const importSettings = useStore((state) => state.importSettings);
  const importInputRef = useRef<HTMLInputElement | null>(null);
  const [statusMessage, setStatusMessage] = useState<string>("");

  useEffect(() => {
    loadFromStorage();
  }, [loadFromStorage]);

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

  const handleImportFile = async (
     event: ChangeEvent<HTMLInputElement>,
  ) => {
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
    <FluentProvider theme={webLightTheme} className={classes.page}>
      <main className={classes.shell}>
        <header className={classes.header}>
          <div className={classes.headerTopRow}>
            <Title3>Outlook Auto-Reply Automater</Title3>
            <div className={classes.headerActions}>
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
    </FluentProvider>
  );
}
