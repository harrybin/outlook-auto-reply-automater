/**
 * DemoApp.tsx – Wrapper that renders the real App inside a demo banner,
 * pre-seeding localStorage with sample data on first visit.
 */

import { useEffect, useCallback } from "react";
import {
  FluentProvider,
  webLightTheme,
  tokens,
  makeStyles,
  Button,
  Tag,
} from "@fluentui/react-components";
import { ArrowCounterclockwise24Regular } from "@fluentui/react-icons";
import { App } from "../taskpane/App";
import { useStore } from "../taskpane/useStore";
import { seedDemoDataIfNeeded, resetDemoData } from "./demoData";

const useStyles = makeStyles({
  banner: {
    backgroundColor: "#fff3cd",
    borderBottom: `2px solid #ffc107`,
    padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalXXL}`,
    display: "flex",
    alignItems: "center",
    gap: tokens.spacingHorizontalM,
    flexWrap: "wrap",
  },
  bannerText: {
    flex: 1,
    fontSize: tokens.fontSizeBase300,
    color: "#664d03",
    minWidth: "200px",
  },
  bannerActions: {
    display: "flex",
    alignItems: "center",
    gap: tokens.spacingHorizontalS,
    flexShrink: 0,
  },
  backLink: {
    fontSize: tokens.fontSizeBase200,
    color: "#0969da",
    textDecoration: "none",
    ":hover": {
      textDecoration: "underline",
    },
  },
});

function DemoBanner() {
  const classes = useStyles();
  const loadFromStorage = useStore((s) => s.loadFromStorage);

  const handleReset = useCallback(() => {
    resetDemoData();
    loadFromStorage();
  }, [loadFromStorage]);

  // Use Vite's injected base URL so this works both locally and on GitHub Pages
  const indexUrl = import.meta.env.BASE_URL ?? "/";

  return (
    <div className={classes.banner} role="banner">
      <Tag appearance="filled" color="warning" size="small">
        Demo Mode
      </Tag>
      <span className={classes.bannerText}>
        This is an interactive demo of the add-in UI. All data is stored
        locally in your browser – no Microsoft 365 account required.
      </span>
      <div className={classes.bannerActions}>
        <Button
          icon={<ArrowCounterclockwise24Regular />}
          size="small"
          appearance="subtle"
          onClick={handleReset}
          title="Reset demo data to defaults"
        >
          Reset
        </Button>
        <a className={classes.backLink} href={indexUrl}>
          ← Back to overview
        </a>
      </div>
    </div>
  );
}

export function DemoApp() {
  useEffect(() => {
    seedDemoDataIfNeeded();
    // Trigger store to pick up freshly seeded data
    useStore.getState().loadFromStorage();
  }, []);

  return (
    <FluentProvider theme={webLightTheme}>
      <DemoBanner />
      <App />
    </FluentProvider>
  );
}
