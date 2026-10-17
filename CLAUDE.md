# CLAUDE.md — email-mcp (Jon's local fork of @marlinjai/email-mcp)

This checkout is a FORK deployed live on this machine: Claude Code's `email-mcp` MCP server runs from `dist/` here (launched via `wrapper.mjs` → `dist/index.js`; `wrapper.mjs` hard-codes absolute paths and appends debug output to `mcp-debug.log`).

## Build & deploy
- Package manager is **npm** (`package-lock.json`), not pnpm. Build: `npm run build` (esbuild via `build.mjs`). Tests: `npm test` (vitest).
- After ANY rebuild the MCP server needs a **FULL Claude Code restart** — an MCP reconnect is NOT sufficient (the old process keeps serving the stale dist).
- MCP config lives in BOTH `~/.claude.json` AND `~/.claude/.mcp.json` — keep them in sync.

## Fork-only changes (past upstream v1.2.3 / CHANGELOG top)
- `8afe9f8` — `email_reply` recipient overrides (`to`/`cc`/`bcc`/`additionalRecipients`), threaded drafts (`email_draft_create` + `inReplyToEmailId`), forward preserves original attachments.
- `b2a4031` — headless `watcher-fetch` CLI (`src/cli/watcher-fetch.ts`) for the BCC inbox-watcher Phase B.
- `b896120` — Gmail adapter: attachments, HTML bodies and reply threading were silently dropped; fixed.
- 2026-09-25 (uncommitted) — `email_draft_create` takes `attachmentPaths`; Outlook fresh (non-reply) drafts now carry attachments too.
- 2026-09-30 (uncommitted): `email_calendar_event_create` (Outlook only) creates an event with a Teams link. It uses `isOnlineMeeting` + `teamsForBusiness` via `src/providers/outlook/calendar.ts` and `src/tools/calendar.ts`.
  - `Calendars.ReadWrite` is requested by INCREMENTAL consent (`OutlookAuth.CALENDAR_SCOPES`), so mail scopes and existing tokens are untouched.
  - Grant it once with `node dist/cli/grant-calendar.js`, which runs a browser sign-in and writes only `msal-cache.json`.
- CHANGELOG.md stops at upstream 1.2.3 — git log is the truth for fork history.

## Behavioral traps (each backed by a real incident)
- **`email_calendar_event_create` WITH attendees sends the invitations immediately** (events have no draft state). Without attendees it only books the organizer's calendar and returns the Teams `joinUrl`.
- **`email_reply` ALWAYS SENDS — it never drafts.** Anything needing review first must use `email_draft_create` (+`inReplyToEmailId`).
- `email_search` returns `attachments: []` even when attachments exist — verify with `email_get`.
- Accounts drift to `connected: false` in the cached account list; `email_test_account` is the live truth.
- Outlook chokes on complex/date search filters — prefer subject search.

## House rules
- Never send email without Jon's explicit "send". No em dashes in drafted email text.
