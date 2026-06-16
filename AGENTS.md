# AGENTS.md

Guidance for AI coding agents working in this repository.

## Project At A Glance
- Outlook add-in taskpane app for auto-reply automation.
- React + TypeScript + Vite frontend.
- State management with Zustand.
- Microsoft Graph + Outlook/Teams/calendar integrations in service modules.

## Start Here
- Primary docs: [README.md](README.md)
- Build/test script source of truth: [package.json](package.json)
- Add-in metadata and permissions: [manifest.json](manifest.json)
- App entry and composition: [src/taskpane/App.tsx](src/taskpane/App.tsx)
- Global store and persisted state actions: [src/taskpane/useStore.ts](src/taskpane/useStore.ts)
- Domain contracts: [src/taskpane/types/index.ts](src/taskpane/types/index.ts)

## Common Commands
- Install deps: `npm install`
- Local UI dev: `npm run dev`
- Start add-in (sideload flow): `npm run start`
- Stop add-in flow: `npm run stop`
- Type check only: `npm run type-check`
- Lint: `npm run lint`
- Tests: `npm run test`
- Build: `npm run build`

## Important Build Behavior
- `npm run build` triggers `prebuild`, which runs version bump logic in [scripts/bump-version.mjs](scripts/bump-version.mjs).
- This modifies [package.json](package.json) and [manifest.json](manifest.json), creating git diffs even when code is unchanged.
- Prefer `npm run type-check`, `npm run lint`, and `npm run test` during normal iteration unless a build artifact is required.

## Architecture Boundaries
- UI components stay in [src/taskpane/components](src/taskpane/components).
- App-level state and CRUD actions stay in [src/taskpane/useStore.ts](src/taskpane/useStore.ts).
- External integrations and business logic stay in [src/taskpane/services](src/taskpane/services).
- Shared model types stay in [src/taskpane/types/index.ts](src/taskpane/types/index.ts).
- Utilities stay focused and small in [src/taskpane/utils](src/taskpane/utils).

## Service Ownership
- Auth/Graph client: [src/taskpane/services/authService.ts](src/taskpane/services/authService.ts)
- Calendar fetch and rule matching: [src/taskpane/services/calendarService.ts](src/taskpane/services/calendarService.ts)
- Outlook automatic replies: [src/taskpane/services/autoReplyService.ts](src/taskpane/services/autoReplyService.ts)
- Teams status/presence: [src/taskpane/services/teamsService.ts](src/taskpane/services/teamsService.ts)
- Location trigger evaluation: [src/taskpane/services/locationService.ts](src/taskpane/services/locationService.ts)
- AI suggestions/fallback logic: [src/taskpane/services/copilotService.ts](src/taskpane/services/copilotService.ts)

## Conventions To Follow
- Keep strict TypeScript compliance (see [tsconfig.json](tsconfig.json)).
- Use the `@` alias for imports from `src` (see [vite.config.ts](vite.config.ts)).
- Preserve existing functional module style in services (small exported functions, explicit types).
- Maintain immutable store updates in Zustand actions.
- Keep comments concise and only where logic is non-obvious.

## Known Pitfalls
- Missing AAD env configuration can break auth flows (see [src/taskpane/services/authService.ts](src/taskpane/services/authService.ts)).
- Local HTTPS cert expectations come from Office add-in tooling (see [vite.config.ts](vite.config.ts)).
- Some rule fields (`body`, `organizer`) are currently mapped as empty in calendar normalization and may not match as expected (see [src/taskpane/services/calendarService.ts](src/taskpane/services/calendarService.ts)).
- Wi-Fi SSID detection is effectively unavailable in browser runtime and returns `null` in the current implementation (see [src/taskpane/services/locationService.ts](src/taskpane/services/locationService.ts)).

## Change Scope Guidance
- Keep edits targeted; do not refactor unrelated modules.
- Do not silently change manifest permissions or app IDs without explicit request.
- If touching service contracts, update corresponding types in [src/taskpane/types/index.ts](src/taskpane/types/index.ts).
- If changing user-visible behavior, add or update tests where practical.

## Validation Checklist Before Finishing
- Run `npm run type-check`
- Run `npm run lint`
- Run `npm run test`
- Run `npm run build` only when build output/version bump validation is explicitly needed
