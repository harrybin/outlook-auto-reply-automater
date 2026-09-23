# Outlook Auto-Reply Automater

An Outlook add-in that automatically manages your out-of-office auto-reply messages based on calendar appointments, location, and Microsoft Teams status.

## ⬇️ Install the Add-in

The latest production manifest is always available directly from the hosted deployment:

**[📥 Download manifest.json](https://outlook-auto-reply-automater.harrybin.de/manifest.json)**

### How to install in Outlook

1. Click the download link above and save `manifest.json` to your computer.
2. Open **Outlook** (desktop or web).
3. Go to **Get Add-ins** / **Apps** (or **Manage Add-ins** in Outlook on the web).
4. Choose **My add-ins** → **Add a custom add-in** → **Add from file…**
5. Select the downloaded `manifest.json`.
6. The **Auto-Reply Automater** add-in will appear in your Outlook ribbon or Apps menu.

> **Note:** This hosted manifest is generated from the root `manifest.json` during deployment and points to `https://outlook-auto-reply-automater.harrybin.de` — no local server required.

---

## Features

- **Calendar-based automation** – Trigger auto-reply messages based on appointment keywords, duration, busy status, and categories.
- **Recipient targeting** – Send Outlook auto-replies to internal users, external users, or both, with duration-based safety recommendations.
- **Ready-to-use replies** – Includes preconfigured messages and matching rules for travel, training, and customer visits.
- **Flexible matching rules** – Match appointments by title, body, location, organizer, or category using operators like contains, startsWith, endsWith, equals, and regex.
- **Timing control** – Configure auto-reply to activate hours before an appointment and deactivate hours after it ends.
- **Location awareness** – Factor in geographic location and WiFi network to determine when to enable auto-reply.
- **Teams status integration** – Coordinate auto-reply with your Microsoft Teams presence status.
- **Multiple profiles** – Create multiple automation profiles with different rules and priorities.
- **Copilot-powered suggestions** – Get AI-powered suggestions for auto-reply messages based on your calendar context.
- **HTML & plain-text replies** – Compose auto-reply messages in HTML or plain text.

## Prerequisites

- [Node.js](https://nodejs.org/) (v18 or later recommended)
- npm (included with Node.js)
- Microsoft Outlook (desktop or web)
- A Microsoft 365 account

## Getting Started

### Installation

```bash
npm install
```

### Development

Start the local development server with HTTPS:

```bash
npm run dev
```

The add-in will be served at `https://localhost:3000`.

### Sideloading the Add-in

To load the add-in into Outlook for local development:

```bash
npm run start
```

This validates the manifest and launches Outlook with the add-in sideloaded against `https://localhost:3000`.

### Keeping the Local Server Running (Windows Startup)

The root [`manifest.json`](manifest.json) points at `https://localhost:3000`, so a sideloaded add-in only works while that dev server is running. To start it automatically at every logon:

```powershell
npm run startup:install
```

This registers a per-user Windows scheduled task named `OutlookAutoReplyAutomater-DevServer` that:

- Starts [`scripts/start-addin-server.ps1`](scripts/start-addin-server.ps1) hidden, 30 seconds after logon.
- Binds Vite to port 3000 with `--strictPort`, so the add-in URL never silently shifts to another port.
- Exits quietly when port 3000 is already served, so it never conflicts with a manual `npm run dev`.
- Installs dependencies automatically if `node_modules` is missing, and warns when the local HTTPS certificate is untrusted.
- Restarts up to three times if the server crashes, and runs without an execution time limit.
- Logs to `%LOCALAPPDATA%\outlook-auto-reply-automater\logs\dev-server.log` (rotated at 5 MB).

Administrator rights are not required. Manage the task with:

| Command | Description |
|---------|-------------|
| `npm run startup:install` | Register the logon task and start it immediately |
| `npm run startup:status` | Show task state, last result, and whether port 3000 is listening |
| `npm run startup:start` | Start the server now without logging off |
| `npm run startup:stop` | Stop the running server |
| `npm run startup:uninstall` | Remove the task so the server no longer starts at logon |
| `npm run startup:run` | Run the server in the foreground for troubleshooting |

To use a different port, pass it through to the install script and update `manifest.json` to match:

```powershell
powershell -ExecutionPolicy Bypass -File scripts\install-startup-task.ps1 -Port 3000 -DelaySeconds 60 -StartNow
```

> **Note:** Run `npx office-addin-dev-certs install` once before relying on the startup task. Without a trusted certificate, Outlook refuses to load the task pane even though the server is reachable.
>
> **Tip:** If you prefer not to depend on a local server, install the hosted production manifest described at the top of this file instead.

## How to Use the Extension

After sideloading, use this flow to configure and run automation:

### Where to Find It in Outlook

- **New Outlook (Windows/Mac) and Outlook on the web**
    - Open any email item.
    - Go to **Apps** (or **More actions** > **Apps**) and select **Auto-Reply Automater**.
    - You can pin the add-in/task pane for quicker access.

- **Classic Outlook (Windows desktop)**
    - Open any email item.
    - In the ribbon, open the **Auto-Reply** tab.
    - In the **Auto-Reply** group, click **Open Auto-Reply**.

### Where Configuration Lives

- Configuration is done inside the add-in task pane UI.
- **Auto-Reply Messages** section: create/edit message templates (name, subject, body, HTML/plain text).
- **Automation Profiles** section: define rules, timing, priority, and enable/disable behavior.
- Settings are persisted locally in browser storage under the key `outlookAutoReplyAutomater_settings`.

1. **Open the add-in in Outlook**
    - In Outlook, open the add-in taskpane from the ribbon/app launcher.
    - Sign in when prompted so the add-in can access Microsoft Graph data (calendar/presence) with the permissions in the manifest.

2. **Create one or more auto-reply messages**
    - Add a message template with a subject and body.
    - Choose plain text or HTML format depending on your needs.

3. **Create an automation profile**
    - Link the profile to one of your message templates.
    - Choose whether Outlook should reply to internal users, external users, or both.
    - The initial configuration includes profiles that recognize `Reise`, `Training`, and `außer Haus beim Kunden` in the appointment title. Training requires a minimum duration of four hours, so multi-day events also match.
    - Add matching rules for appointment fields such as title, location, organizer, or categories.
    - Optionally constrain by busy status, all-day flag, and minimum or maximum duration in hours. Both limits can be set together.

4. **Set timing and priority**
    - Configure how many hours before an event to enable auto-reply.
    - Configure how many hours after an event to disable auto-reply.
    - Set profile priority (lower number = higher priority) to resolve conflicts when multiple profiles match.

5. **Optionally add context signals**
    - Add location-based conditions.
    - Add Teams presence conditions.
    - Use Copilot suggestions to draft or improve response text.

6. **Save and verify behavior**
    - Save your profiles and keep the add-in running.
    - Create a test calendar event that matches your rules.
    - Confirm Outlook automatic replies are enabled/disabled at the expected times.

### Tips

- Start with one simple profile and one test event before adding complex rule combinations.
- If multiple profiles can match, double-check priority order first.
- If automation does not trigger, verify sign-in/consent and calendar rule conditions.

## Available Scripts

| Command | Description |
|---------|-------------|
| `npm run dev` | Start Vite dev server with HTTPS |
| `npm run build` | Type-check and build for production |
| `npm run preview` | Preview the production build locally |
| `npm run lint` | Run ESLint on source and test files |
| `npm run lint:fix` | Auto-fix lint issues |
| `npm run test` | Run tests with Vitest |
| `npm run test:watch` | Run tests in watch mode |
| `npm run test:coverage` | Run tests with coverage report |
| `npm run type-check` | Run TypeScript type checking without emitting |
| `npm run validate` | Validate the Office add-in manifest |
| `npm run startup:install` | Register a Windows logon task that keeps `https://localhost:3000` running |
| `npm run startup:status` | Show the startup task state and whether port 3000 is listening |
| `npm run startup:start` | Start the local server via the scheduled task |
| `npm run startup:stop` | Stop the local server started by the scheduled task |
| `npm run startup:uninstall` | Remove the Windows logon task |
| `npm run startup:run` | Run the local server in the foreground (troubleshooting) |
| `npm run pack -- https://your-host.example.com/app` | Create a Windows/Mac deployment bundle in `dist/deployment` |
| `npm run pack:manifest` | Export the raw manifest without building a deployment bundle |
| `npm run version:bump` | Bump the patch version |
| `npm run version:bump:minor` | Bump the minor version |
| `npm run version:bump:major` | Bump the major version |

## Deployment

For Outlook on Windows and Mac, deployment requires two pieces:

- A hosted HTTPS web app for the add-in UI
- A generated manifest that points to that hosted URL

Create a deployment bundle with:

```bash
npm run build
npm run pack -- https://your-host.example.com/outlook-auto-reply-automater
```

This writes a ready-to-deploy bundle to `dist/deployment`:

- `dist/deployment/web/` contains the static web assets you must host over HTTPS
- `dist/deployment/manifest.json` contains a production manifest generated from the root `manifest.json`
- `dist/deployment/DEPLOY.md` contains the deployment steps

After hosting `dist/deployment/web/`, sideload or deploy `dist/deployment/manifest.json` into Outlook.

### GitHub Pages Deployment (Automated)

This repository includes a GitHub Actions workflow at `.github/workflows/deploy-github-pages.yml`.

- On every push to `main` (or manual trigger), it:
    - Builds the web app for the custom-domain root path.
    - Creates a production deployment bundle using `https://outlook-auto-reply-automater.harrybin.de` as `DEPLOY_BASE_URL`.
    - Publishes `dist/deployment/web/` to GitHub Pages.
    - Uploads `dist/deployment/manifest.json` as a workflow artifact named `outlook-addin-manifest`.

Before using production sign-in on GitHub Pages, add this redirect URI to your Azure app registration:

- `https://outlook-auto-reply-automater.harrybin.de/src/taskpane/index.html`

If needed, you can also override the redirect URI with `VITE_AAD_REDIRECT_URI`.

## Project Structure

```
src/
└── taskpane/
    ├── components/        # React UI components
    │   ├── AutoReplyList.tsx   # Manage auto-reply message templates
    │   └── ProfileList.tsx     # Manage automation profiles
    ├── services/          # Business logic & API integrations
    │   ├── authService.ts      # MSAL authentication
    │   ├── autoReplyService.ts # Outlook auto-reply management
    │   ├── calendarService.ts  # Calendar event access
    │   ├── copilotService.ts   # Copilot AI suggestions
    │   ├── locationService.ts  # Location/WiFi detection
    │   ├── storageService.ts   # Local storage persistence
    │   └── teamsService.ts     # Teams presence integration
    ├── types/             # TypeScript type definitions
    ├── utils/             # Utility functions
    └── useStore.ts        # Zustand state management
scripts/
└── bump-version.mjs     # Version bumping utility
```

## How It Works

1. **Define auto-reply messages** – Create one or more message templates with a subject and body.
2. **Create automation profiles** – Set up profiles that specify when each message should be activated, based on:
   - Appointment keyword/field matching rules
   - Appointment duration thresholds
   - Calendar busy status (free, tentative, busy, out-of-office, working elsewhere)
   - All-day event detection
   - Internal, external, or combined auto-reply recipients
3. **Configure timing** – Optionally activate the auto-reply before an appointment starts and deactivate it after it ends, with configurable hour offsets for both.
4. **Set priorities** – When multiple profiles match, the one with the lowest priority number wins.
5. **Let it run** – Run the independent Windows automation runner so rules continue to apply when the Outlook task pane is closed.

## When Rules Are Applied

The task pane is configuration UI only. The independent Windows runner (`npm run runner`) evaluates rules immediately and then polls every 60 seconds by default. It runs outside Outlook's task pane and can remain active for the whole Windows session.

To use it:

1. Export settings from the task pane and save the JSON as `%APPDATA%\outlook-auto-reply-automater\settings.json`.
2. Set `AAD_CLIENT_ID` (and optionally `AAD_TENANT_ID`) in the runner's environment. The app registration must allow public-client device-code flow and grant `User.Read`, `Calendars.Read`, `MailboxSettings.ReadWrite`, and `Presence.ReadWrite`.
3. Run `npm run runner` once and complete the device-code sign-in.
4. Install the logon task with `npm run runner:install`. Use `npm run runner:status` and `npm run runner:uninstall` to manage it.

### Calendar-based rules

Calendar matching considers events from the previous 24 hours through the next 24 hours. A matching profile is active from its appointment start, optionally shifted earlier by `hoursBeforeAppointment`, until its appointment end, optionally shifted later by `hoursAfterAppointment`. The first matching profile in the list wins.

### Location-based rules

The local runner currently applies calendar profiles. Browser-only location signals (geofence and Wi-Fi) are not available to the Node process; no-internet can be added separately using the runner host's network state.

### When the active reply message may change

When the desired rule changes, the runner disables the previously managed Outlook auto-reply, clears the managed Teams presence when configured to restore it, and applies the new message/status. When no rule matches, it clears the managed state. It avoids repeating Graph writes when the active rule has not changed.

> **Tip:** If multiple profiles could match the same appointment data, the profile with the **lowest priority number** wins in the current matching logic.

## Technology Stack

- [React 18](https://react.dev/) – UI framework
- [Fluent UI React v9](https://react.fluentui.dev/) – Microsoft design system components
- [Vite](https://vite.dev/) – Build tool and dev server
- [Zustand](https://zustand.docs.pmnd.rs/) – Lightweight state management
- [MSAL Browser](https://github.com/AzureAD/microsoft-authentication-library-for-js) – Microsoft identity authentication
- [Microsoft Graph](https://learn.microsoft.com/en-us/graph/overview) – API for Outlook and Teams integration
- [Vitest](https://vitest.dev/) – Testing framework
- [TypeScript](https://www.typescriptlang.org/) – Type safety

## License

This project is private. See the repository settings for access details.
