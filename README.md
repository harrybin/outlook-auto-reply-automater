# Outlook Auto-Reply Automater

An Outlook add-in that automatically manages your out-of-office auto-reply messages based on calendar appointments, location, and Microsoft Teams status.

## Features

- **Calendar-based automation** – Trigger auto-reply messages based on appointment keywords, duration, busy status, and categories.
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

This validates the manifest and launches Outlook with the add-in sideloaded.

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
    - Add matching rules for appointment fields such as title, location, organizer, or categories.
    - Optionally constrain by busy status, all-day flag, and minimum duration.

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
| `npm run version:bump` | Bump the patch version |
| `npm run version:bump:minor` | Bump the minor version |
| `npm run version:bump:major` | Bump the major version |

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
3. **Configure timing** – Optionally activate the auto-reply before an appointment starts and deactivate it after it ends, with configurable hour offsets for both.
4. **Set priorities** – When multiple profiles match, the one with the lowest priority number wins.
5. **Let it run** – The add-in monitors your calendar and automatically enables/disables your Outlook auto-reply based on your configured rules.

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
