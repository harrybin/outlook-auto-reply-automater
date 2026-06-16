import { useEffect } from "react";
import {
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
  subtitle: {
    color: tokens.colorNeutralForeground3,
    fontSize: tokens.fontSizeBase300,
    lineHeight: tokens.lineHeightBase300,
    margin: 0,
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

  useEffect(() => {
    loadFromStorage();
  }, [loadFromStorage]);

  return (
    <FluentProvider theme={webLightTheme} className={classes.page}>
      <main className={classes.shell}>
        <header className={classes.header}>
          <Title3>Outlook Auto-Reply Automater</Title3>
          <p className={classes.subtitle}>
            Manage reusable auto-reply messages and the rules that activate
            them.
          </p>
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
