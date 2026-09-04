# PGN WhatsApp QA Harness

This harness drives the consumer WhatsApp Web UI with Playwright. Messages follow the real route through the PGN WhatsApp Business number, LivePerson, and the PGN bot; no direct bot or messaging API is used.

## Getting Started

Install the project, run the guided setup, and open the operator control panel:

```bash
npm install
npm run setup
npm run pgn
```

The setup wizard inspects Node.js, npm, dependencies, Playwright Chromium, the PGN workbooks, `.env`, Google Drive configuration, optional Discord notifications, browser display support, and the saved WhatsApp profile. It can safely create or update `.env`, validate a service-account file, configure a Discord Incoming Webhook through a masked prompt, install Chromium, and optionally open WhatsApp login. Credential values, webhook URLs, and profile contents are never displayed.

`npm run pgn` provides one interactive entry point for:

- full and filtered PGN runs
- approved retest validation, execution, and resume
- workbook, evidence, and Drive validation
- evidence migration
- Discord notification status, testing, and configuration
- WhatsApp login, verification, and explicit authentication recreation
- setup and diagnostics
- TypeScript checks and safe regression tests

Execution, fresh-run preparation, evidence migration, and authentication recreation require explicit confirmation. Cancelling a prompt or selecting Back does not launch the selected action.

Run non-interactive prerequisite checks at any time:

```bash
npm run doctor
```

### Browser Support

Browser commands run directly on Windows and macOS. Linux runs directly when `DISPLAY` is available or Chromium is configured headless. On headless Linux and GitHub Codespaces, the launcher uses `xvfb-run` only when a headed browser needs it. If no display provider is available, the command exits with a concrete setup instruction instead of relying on shell-specific syntax.

### Advanced Commands

The direct commands remain available for automation and experienced operators:

- `npm run whatsapp:login` opens Playwright Chromium, writes only a validated QR to `artifacts/whatsapp-login.png`, and waits for login.
- `npm run whatsapp:verify` verifies that `.whatsapp-profile/` opens without another QR scan.
- `npm run test:single` sends `Halo`, captures all new incoming messages through the configured quiet window, and writes `reports/PGN_Single_Test_Result.xlsx`.
- `npm run test:pgn:validate` validates the real PGN workbook without opening WhatsApp or sending messages.
- `npm run test:pgn` executes the real PGN workbook and progressively saves the executed copy.
- `npm run test:pgn:fresh` archives the current report and prepares a clean full-run workbook without opening WhatsApp.
- `npm run test:pgn:retest:validate` lists approved retest scenarios and readiness without opening WhatsApp.
- `npm run test:pgn:retest` executes only approved retest scenarios.
- `npm run evidence:validate` validates existing local evidence, creates three cleaned previews, and optionally checks Drive access without contacting WhatsApp.
- `npm run evidence:migrate` backs up the completed workbook, cleans and uploads existing evidence, and writes hyperlinks without contacting WhatsApp.
- `npm run discord:validate` safely inspects the configured Discord webhook with `GET` and does not post a message.
- `npm run discord:validate -- --send-test` explicitly posts one Discord test notification.
- `npm run discord:demo` sends a short fake lifecycle through the existing notifier without running WhatsApp, workbook, evidence, or Drive operations.
- `npm run test:response-collector` runs deterministic delayed-message and hard-timeout tests without opening WhatsApp.
- `npm run test:session-reset` tests reset confirmation and failure handling without opening WhatsApp.
- `npm run test:workbook-writer` verifies result writes preserve the source XLSX table metadata.
- `npm run test:config` verifies `.env`, credential-source precedence, safe credential parsing, and fresh-run preservation without contacting WhatsApp or Drive.
- `npm run test:retest` verifies retest selection and history behavior with temporary workbook fixtures.
- `npm run test:evidence` verifies evidence migration with temporary files and mocked Drive operations.
- `npm run test:operator` verifies setup, diagnostics, menus, cancellation, and platform behavior with mocks only.
- `npm run test:discord` verifies webhook lifecycle, throttling, retries, redaction, and signal handling with mocked HTTP only.
- `npm run data:template` creates a starter legacy test-case workbook without overwriting an existing one.
- `npm run check` runs the TypeScript compiler.

## Configuration

All commands load the repository-root `.env` through the central configuration module. A missing `.env` is valid. Existing process environment variables, including Codespaces Secrets and CI variables, are never overwritten by `.env` values with the same name.

Configure either `PGN_WHATSAPP_PHONE` (international digits, without `+`) or `PGN_WHATSAPP_CHAT`. A phone number is preferred when both are present because the direct WhatsApp chat URL avoids ambiguous chat-name matches. The harness verifies the open conversation header against that target and confirms each outgoing WhatsApp message before collecting a response.

The authenticated profile is fixed at `.whatsapp-profile/`. Data files are restricted to `data/`, and generated reports are restricted to `reports/`. The profile, `.env`, `.secrets/`, service-account JSON files, QR images, evidence, diagnostics, and executed workbooks are gitignored. Treat the profile and credentials as secrets and do not share or commit them. Fresh-run preparation only archives and replaces generated test workbooks; it never removes `.env` or `.secrets/`.

### Discord Notifications

Discord notifications are optional and disabled by default. They use only a Discord Incoming Webhook; no bot token, Gateway connection, Discord application, OAuth flow, or Discord SDK is used. Treat `DISCORD_WEBHOOK_URL` as a secret and store it only in the gitignored `.env`, a Codespaces Secret, or another secret environment variable. Setup accepts it through a masked prompt and never displays an existing value. For Discord values supplied by the higher-precedence process environment or Codespaces Secrets, Setup leaves those keys unchanged, directs you to update their source, and can still configure independent `.env` settings.

The Notifications submenu in `npm run pgn` can show Discord status, send one explicitly confirmed test message, or update notification settings. `npm run doctor` reports only whether notifications are enabled and whether the webhook is configured; it does not contact Discord. `npm run discord:validate` performs a non-posting webhook inspection. A visible test message is sent only when `--send-test` is supplied or an operator explicitly confirms the test in Setup or the Notifications submenu.

The available settings are:

```dotenv
DISCORD_NOTIFICATIONS_ENABLED=false
DISCORD_WEBHOOK_URL=
DISCORD_PROGRESS_EVERY=5
DISCORD_PROGRESS_MINUTES=2
DISCORD_NOTIFY_START=true
DISCORD_NOTIFY_PROGRESS=true
DISCORD_NOTIFY_COMPLETE=true
DISCORD_NOTIFY_FAILURE=true
```

An active full or Ready-for-Retest run creates one live status message, edits it at the configured scenario or time interval, finalizes it, and posts a fresh completion or technical-failure event. An incomplete retest batch is labeled as a checkpoint instead of a completed retest. Setup offers progress every 5 scenarios, every 10 scenarios, final only, or a custom scenario/time cadence; the final-only preset disables start and progress messages. Start, progress, completion, and failure events can also be controlled independently with the advanced flags above.

Notification delivery is fail-open. Timeouts, rate limits, deleted webhooks, malformed responses, and other Discord failures produce a redacted warning but never stop execution or alter workbook results. Discord payloads contain only operational identifiers, counts, timing, technical status, evidence counts, and the executed workbook basename. They never include WhatsApp messages, bot responses, phone numbers, screenshots, credentials, semantic Pass/Fail decisions, or automatic mentions.

#### Discord Verification

Validate configuration without posting a message:

```bash
npm run discord:validate
```

Send exactly one basic test notification:

```bash
npm run discord:validate -- --send-test
```

Simulate a complete start, progress, and completion lifecycle:

```bash
npm run discord:demo
```

Simulate failure or interruption notifications:

```bash
npm run discord:demo -- --fail
npm run discord:demo -- --interrupt
```

`discord:demo` uses only fake operational data and the existing Discord notifier. It does not launch WhatsApp or Playwright, send PGN testcases, read or modify a workbook, upload evidence, access Google Drive, alter run metadata, or perform fresh-run cleanup. The demo uses one-second event delays and a short bounded notification deadline rather than the production progress cadence.

### Google Drive Evidence

Drive evidence is disabled by default. It continues to use Google Drive API v3, Shared Drive support, inherited permissions, and one evidence subfolder per run.

#### Local Development

The recommended local setup is:

```bash
cp .env.example .env
mkdir -p .secrets
```

Place the downloaded Service Account JSON at:

`.secrets/google-service-account.json`

Then edit `.env`:

```dotenv
GOOGLE_DRIVE_EVIDENCE_ENABLED=true
GOOGLE_DRIVE_EVIDENCE_PARENT_FOLDER=<Shared Drive folder URL>
GOOGLE_SERVICE_ACCOUNT_FILE=.secrets/google-service-account.json
```

Validate configuration before migration or execution:

```bash
npm run evidence:validate
```

Both `.env` and `.secrets/` are gitignored. A relative `GOOGLE_SERVICE_ACCOUNT_FILE` path is resolved from the repository root, even if a command is launched from another working directory.

#### Codespaces / CI

Environment variables and GitHub Codespaces Secrets remain fully supported. Set these without creating `.env`:

```text
GOOGLE_DRIVE_EVIDENCE_ENABLED
GOOGLE_DRIVE_EVIDENCE_PARENT_FOLDER
GOOGLE_SERVICE_ACCOUNT_JSON
```

Process environment values take precedence over `.env` values with the same name. The supported credential methods, in priority order, are `GOOGLE_SERVICE_ACCOUNT_JSON`, `GOOGLE_SERVICE_ACCOUNT_JSON_BASE64`, and `GOOGLE_SERVICE_ACCOUNT_FILE`. Only one is required. Validation reports the selected source and warns if multiple methods are configured, but never prints raw or decoded credentials.

`GOOGLE_DRIVE_EVIDENCE_PARENT_FOLDER` accepts either a raw folder ID or a `https://drive.google.com/drive/folders/...` URL. Use a Shared Drive and add the Service Account `client_email` as Content manager. The harness verifies parent-folder and write access, creates or reuses one deterministic subfolder per run, and lets files inherit its access; it never creates `anyoneWithLink` permissions. Credentials, private keys, access tokens, and refresh tokens are never written to the workbook or logs. An optional `LEGACY_EVIDENCE_CROP_LEFT` can override legacy crop auto-detection after manual verification.

## Real PGN Workbook

The source workbook is:

`data/PGN AI Assistant - Knowledge Base Testing Report - User Inputs.xlsx`

The source is never overwritten. The first execution copies it to:

`reports/PGN AI Assistant - Knowledge Base Testing Report - Executed.xlsx`

Later runs resume from that executed copy. Existing Bot Response cells are skipped by default. Existing expected responses, semantic statuses, notes, references, styles, dimensions, merged cells, and completed evidence are preserved. The runner does not generate User Input or assign Passed or Failed.

`Test Case Knowledge Base` reads User Input from column H and writes Bot Response, response time, and test date to I-K. A blank Test Case ID row with a populated Turn continues the preceding scenario.

`Negative Case` reads input from column E and writes results to H-J. Explicit `Turn 1:`, `Turn 2:`, and later markers inside one cell execute sequentially in the same scenario context. Their responses and timings are combined with turn labels.

The executed copy adds `Execution Transcript` with per-turn user and bot messages, timestamps, first-response timing, total timing, technical status, and evidence paths. Every bot bubble has its own transcript row. Multiple bubbles are labeled `Message 1:`, `Message 2:`, and later in the logical Bot Response cell. The workbook is atomically replaced after every attempted turn.

Evidence hyperlinks use column N on both report sheets. This column remains outside the immutable `Negative Case` table at A:M. `Execution Transcript` keeps local paths in M, adds hyperlink column N and evidence-only status column O, and `Execution Metadata` stores run-folder and Drive file IDs without using an Excel table or storing credentials.

ExcelJS misreads this source workbook's table defaults and otherwise emits invalid AutoFilter metadata. Each atomic save therefore restores every immutable source table definition byte-for-byte into the table targets generated by ExcelJS, validates the table relationships and content types, and only then replaces the executed workbook. Resume is refused if the executed workbook's table structure or source-owned input cells no longer match the source.

## Existing Run Evidence Backfill

Configure Drive, then run:

```bash
npm run evidence:validate
npm run evidence:migrate
```

Validation reads the completed workbook and existing `artifacts/evidence/*.png`, checks transcript mapping and missing screenshots, and creates representative previews for `PGN-KB-003`, `PGN-KB-031` turn 2, and `PGN-KB-075`. The legacy crop boundary is derived and cross-checked from image pixels instead of accepting an arbitrary coordinate.

Migration never imports the WhatsApp runner, opens Playwright, sends `reset`, or sends testcase messages. It reads only the active evidence directory, does not import pre-fix archives, preserves original screenshots, writes cleaned files under `artifacts/evidence/clean/<RUN_ID>/`, creates an exact `reports/archive/*-before-evidence*.xlsx` backup for every invocation, uploads PNGs with up to three attempts, and saves the workbook after every evidence record. Missing or unusable evidence is reported as `EVIDENCE_MISSING` or `EVIDENCE_REQUIRES_RERUN`; it is never regenerated automatically. Upload failures retain the cleaned local file, record `EVIDENCE_UPLOAD_ERROR`, clear any stale current-result link, and continue when safe.

Migration is resumable. Stored Drive file IDs are reused, exact-name files are found before upload, and run folders are recovered from `Execution Metadata` or by deterministic name. Local crop bytes and Drive checksums are revalidated on every invocation, so stale cached or remote content is updated without creating duplicates. Runs recorded as `FUTURE` or `RETEST` retain their already scoped conversation-pane PNGs and are never passed through the legacy crop.

## Future Run Evidence

Normal PGN execution waits for response settlement, scrolls the active conversation to the bottom, and captures a Playwright `Locator.screenshot()` of the visible conversation pane selected from `#main` or the semantic conversation wrapper. The full WhatsApp page, navigation rail, chat list, search area, and unrelated contacts are excluded. Full-page screenshots remain available only for local failure diagnostics and are not uploaded as normal evidence.

When Drive evidence is enabled, the runner validates the parent and creates the run folder before opening WhatsApp. Each attempted turn is captured locally when possible, uploaded, linked in Excel, and atomically saved with its current result before the next turn or reset. A rerun clears stale result values and links that are not produced by the new attempt. Evidence upload status remains separate from `CAPTURED`, `TIMEOUT`, `SEND_ERROR`, and `CHAT_ERROR`, so a Drive failure does not invalidate a chatbot response.

## Filters And Resume

```bash
npm run test:pgn -- --limit 5
npm run test:pgn -- --sheet kb
npm run test:pgn -- --sheet negative
npm run test:pgn -- --test PGN-KB-031
npm run test:pgn -- --rerun PGN-KB-003
npm run test:pgn -- --test PGN-KB-031 --rerun PGN-NEG-018
```

`--rerun` without an ID reruns the selected set. `--rerun ID` selects and reruns that scenario. Without rerun, completed scenarios are skipped. A partially completed multi-turn scenario is skipped because continuing it later would not guarantee the original turn context.

## Full New Run

Prepare, validate, and launch these as separate commands:

```bash
npm run test:pgn:fresh
npm run test:pgn:validate
npm run test:pgn
```

`test:pgn:fresh` never starts WhatsApp. It creates an exact timestamped archive of the previous executed workbook, creates a new executed copy from the immutable source, clears generated Bot Response, Response Time, Test Date, Evidence, transcript, evidence metadata, and retest sheets, then exits. Prepared User Input, Expected Bot Response, Status, and Notes remain unchanged. Normal full-run selection ignores semantic Status values, so every scenario is runnable after fresh preparation.

## Retest Fixed Cases

Set the primary scenario row in the executed workbook to `Ready for Re-test`, then validate and launch separately:

```bash
npm run test:pgn:retest:validate
npm run test:pgn:retest
```

Retest selection is trimmed and case-insensitive. It accepts only `Ready for Re-test` and the exact alias `Ready for Retest`; values such as `Failed`, `Blocked`, `Review`, `ready`, and `retest` are not automatic candidates. Recognized statuses are `Passed`, `Failed`, `Blocked`, `Review`, `Ready for Re-test`, and `Pending Evaluation`. Unknown non-empty statuses are reported as workbook validation warnings and are never interpreted as retest approval.

The complete selection is frozen before execution starts. A positive multi-turn scenario selected from its primary row executes every continuation turn, and a negative multi-turn scenario executes every parsed turn without a reset between turns. Unselected result rows are not changed.

Before the first reset for each selected scenario, the runner appends an idempotent snapshot to `Retest History`. Positive multi-turn scenarios receive one history row per result row. The snapshot retains the previous transcript Run ID, semantic Status, Bot Response, Response Time, Test Date, and Evidence URL. New results and evidence are written back to the same history row as the retest progresses, while the active report points to the latest execution.

Every batch receives a new ID such as `RETEST-20260902T053000Z`. The ID is used by `Execution Transcript`, `Retest Metadata`, evidence metadata, and the dedicated Drive folder `PGN-WhatsApp-Retest-20260902T053000Z`. Previous Drive files and folders are never deleted or replaced by another run.

After every turn, transcript, evidence metadata, active results, history, and resume state are atomically saved. When all turns in a scenario have technical status `CAPTURED`, its semantic Status becomes `Pending Evaluation`; automation never assigns `Passed` or `Failed`. A timeout or send/chat error keeps the previous semantic Status and records the technical failure separately.

Useful retest filters are:

```bash
npm run test:pgn:retest -- --limit 3
npm run test:pgn:retest -- --test PGN-KB-075
npm run test:pgn:retest -- --resume RETEST-20260902T053000Z
```

`--test` explicitly selects only the named scenario and prints a warning if its current Status is not `Ready for Re-test`. `--resume` reloads the immutable selected-ID set from `Retest Metadata`, skips scenarios already completed successfully in that same run, retries prior technical failures, reuses its Drive folder, and deduplicates history rows. If all scenarios were saved but final session cleanup failed, resume retries that cleanup before marking the run complete. `--resume` cannot be combined with `--test` or `--sheet`, though `--limit` can constrain the remaining resumed scenarios. If a new selection is empty, the retest command exits successfully before Drive setup or WhatsApp startup.

Selected retests require Google Drive evidence to be configured. Parent-folder authentication and retest-folder creation are completed before WhatsApp opens; invalid or disabled Drive configuration aborts without sending a testcase message. Per-file upload failures after startup remain non-fatal and are recorded separately from chatbot technical status.

## Response Completion

The runner identifies response ownership from the confirmed outgoing message ID, excludes pre-existing and outgoing messages, and captures every new incoming bubble in DOM order. After the first bubble, every new or updated bubble restarts `WHATSAPP_RESPONSE_IDLE_MS`, which defaults to 10000 ms. `WHATSAPP_RESPONSE_TIMEOUT_MS` is the hard response limit and defaults to 60000 ms. A visible `typing` or `mengetik` state holds and then restarts the quiet timer, but typing detection is only an additional signal.

`firstResponseMs` ends at the first captured bubble. `totalResponseMs` ends at the last captured bubble and excludes the final idle confirmation period. Each turn in a multi-turn scenario independently waits for complete response settlement before the next turn is sent.

## Session Isolation

Independent scenarios are isolated with the deployed Conversation Builder debug command `reset`. Before every runnable scenario, including the first remaining scenario after resume, the runner snapshots WhatsApp, sends `reset`, and waits only for a new incoming response containing `Session deleted`. `PGN_RESET_COMMAND`, `PGN_RESET_CONFIRMATION`, and `PGN_RESET_TIMEOUT_MS` configure this contract and default to `reset`, `Session deleted`, and 30000 ms.

After reset confirmation, the runner requires `POST_RESET_QUIET_MS` of silence, defaulting to 10000 ms. Any new or changed incoming message is recorded as `STALE_BOT` and restarts that timer, so it cannot be assigned to the next testcase. A visible typing state also holds the drain and restarts the quiet timer when it clears, but message arrival remains authoritative. The reset occurs outside the scenario turn loop. Multi-turn scenarios therefore retain context across every turn, and the next reset is attempted only after the completed scenario has been written and atomically saved. A final reset and drain also run after the last selected scenario.

Reset traffic never enters User Input, Bot Response, expected handling, or semantic Status cells. It is recorded in `Execution Transcript` as `CONTROL_USER`, `CONTROL_BOT`, `CONTROL_SYSTEM`, or `STALE_BOT`.

If the expected confirmation is not captured before the reset timeout, the runner saves a reset-failure screenshot and diagnostics, records the failed control attempt, saves the executed workbook, aborts all remaining scenarios, and exits non-zero. This fail-safe prevents results from being collected under uncertain bot context.

Isolation depends on the deployed bot continuing to allow the Conversation Builder reset/debug command. LivePerson recommends disabling debug commands in production, so this deployment capability must remain enabled for the QA harness.

`CAPTURED`, `TIMEOUT`, `SEND_ERROR`, and `CHAT_ERROR` are technical states only. They are written to the transcript and technical failures are appended to Notes without replacing existing notes. They never produce a semantic Passed or Failed decision.
