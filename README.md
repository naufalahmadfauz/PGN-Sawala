# PGN Sawala

This harness runs the PGN workbook through either the consumer WhatsApp Web UI with Playwright or an opt-in LivePerson REST transport. Both use the current shared workbook execution engine; transport and session mode are selected per run.

**v1.0.0** is the first stable release. See the [release notes](RELEASE_NOTES_v1.0.0.md) and [changelog](CHANGELOG.md) for the V1 overview and verification results.

## Test Transports

| Transport | Purpose | Channel And Evidence |
| --- | --- | --- |
| WhatsApp (default) | Real-channel acceptance testing | Consumer WhatsApp Web -> PGN WhatsApp Business number -> LivePerson -> PGN bot. Browser screenshots and optional Drive evidence. |
| REST (opt-in) | Bulk bot testing without a browser | Direct LivePerson messaging requests. Workbook results and transcripts, **no screenshots**, no WhatsApp profile, and no Drive requirement or uploads. |

REST results do not prove WhatsApp delivery, rendering, or end-to-end channel acceptance. Keep WhatsApp acceptance runs for those checks. REST is not a parallel grading engine: source inputs, selection, output workbook, retest history, technical status, and semantic evaluation rules stay in the shared runner. Neither transport automatically grades responses as Passed or Failed.

Open `npm run pgn` -> **Run tests** -> **REST Bulk Test**, choose **Isolated** (default) or **Continuous**, then choose a full run, IDs, a sheet, retest, validation, or Back. Execution requires default-No confirmation. The existing WhatsApp choices remain available. Neither enabling REST in Setup nor choosing a session mode changes the default transport for later runs.

REST commands are explicit alternatives to the WhatsApp commands:

```bash
npm run test:pgn:rest:validate
npm run test:pgn:rest
npm run test:pgn:rest -- --sheet kb
npm run test:pgn:rest -- --test PGN-KB-031 --rerun
npm run test:pgn:rest -- --session=continuous
npm run test:pgn:rest:retest
```

The REST execution/retest entrypoints force REST and call the shared `runPgnWorkbook`; `--transport=rest` and `--transport=whatsapp` select transport on the shared commands. REST **Full run** preserves normal source-workbook selection and skips existing completed results. It does not silently clear output or force a rerun. For a fresh bulk run, use the existing separate fresh-workbook preparation action, or explicitly choose filtered IDs/`--rerun` as appropriate. Both transports share the configured executed workbook, so existing results can affect selection when switching transports.

`test:pgn:rest:validate` checks workbook, configuration, and output readiness before making **real but harmless domain/authentication requests**: domain discovery, app JWT authentication, and synthetic consumer JWS validation only. It never creates conversations or sends testcase messages. The REST validation menu and Setup authentication check both explain these requests and require explicit confirmation. Default diagnostics and Setup inspection do not contact LivePerson.

The smoke command is a separate **real integration action**, never an automatic setup, validation, regression-test, or acceptance step:

```bash
# Explicitly opt into a real LivePerson conversation and smoke message.
npm run rest:smoke
```

Do not run the smoke command or REST execution commands as part of safe mock-only verification. Domain/auth validation is not a smoke conversation and does not establish bot-response readiness.

## Getting Started

With Node.js 20.12 or newer and npm installed, run the guided setup and open the operator control panel:

```bash
npm install
npm run setup
npm run pgn
```

The setup wizard inspects Node.js, npm, dependencies, Playwright Chromium, the PGN workbooks, `.env`, Google Drive configuration, optional Discord notifications, optional LivePerson REST configuration, browser display support, and the saved WhatsApp profile. It can safely create or update `.env`, validate a service-account file, configure Discord and REST credentials through masked prompts, install Chromium, and optionally open WhatsApp login. REST domain/auth validation is separately offered with default-No confirmation; enabling REST alone never opens a conversation. Credential values, webhook URLs, and profile contents are never displayed.

`npm run pgn` provides one interactive entry point for:

- full and filtered PGN runs through WhatsApp or REST
- approved retest validation, execution, and resume
- interrupted-run discovery and guided recovery
- header-based workbook schema inspection and mapping review
- workbook, evidence, and Drive validation
- evidence migration
- Discord notification status, testing, and configuration
- WhatsApp login, verification, and explicit authentication recreation
- setup and diagnostics
- TypeScript checks and safe regression tests

Execution, fresh-run preparation, evidence migration, and authentication recreation require explicit confirmation. Cancelling a prompt or selecting Back does not launch the selected action.

Each new full, filtered, or retest execution asks for a Session Mode, defaulting to **Isolated**, including full tests launched after Setup. **Continuous** requires an explicit per-run choice, displays the shared-context warning, and keeps execution confirmation defaulted to No. The choice is not a permanent setting or an environment variable. Preparing a fresh workbook remains a separate action and never starts execution.

Run non-interactive prerequisite checks at any time:

```bash
npm run doctor
npm run doctor -- --transport=rest
```

Default doctor checks WhatsApp prerequisites and may validate configured Google Drive access; it does not send WhatsApp testcases or Discord notifications. Its LivePerson REST check is configuration-only by default: disabled, configured/domain override status, or configuration errors, with authentication explicitly marked not checked. REST-only doctor still checks common configuration and workbook schema but does not require or probe Playwright, Chromium, Xvfb, a WhatsApp target/profile, or Drive. LivePerson network checks require an explicit access-check request; use `test:pgn:rest:validate` for the dedicated workbook/domain/auth check. The control panel's recovery demo uses local synthetic artifacts and suppresses external checks.

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
- `npm run test:pgn:rest` executes the same workbook through LivePerson REST without screenshots or a browser.
- `npm run test:pgn:rest:validate` checks workbook/config/output readiness and makes real domain/auth requests only, never conversations or testcase messages.
- `npm run test:pgn:rest:retest` retests through REST using shared selection and history rules, without Drive requirements.
- `npm run rest:smoke` is an explicitly invoked real REST conversation/message check, never part of safe regression tests.
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

For WhatsApp execution, configure either `PGN_WHATSAPP_PHONE` (international digits, without `+`) or `PGN_WHATSAPP_CHAT`. A phone number is preferred when both are present because the direct WhatsApp chat URL avoids ambiguous chat-name matches. The harness verifies the open conversation header against that target and confirms each outgoing WhatsApp message before collecting a response. REST does not require either setting.

The authenticated profile is fixed at `.whatsapp-profile/`. Data files are restricted to `data/`, and generated reports are restricted to `reports/`. The profile, `.env`, `.secrets/`, service-account JSON files, QR images, evidence, diagnostics, and executed workbooks are gitignored. Treat the profile and credentials as secrets and do not share or commit them. Fresh-run preparation only archives and replaces generated test workbooks; it never removes `.env` or `.secrets/`.

### LivePerson REST

REST is disabled by default. Use Setup's opt-in step or configure these values in the gitignored `.env` or process-managed secrets. The values below are placeholders, not working credentials or tokens:

```dotenv
LIVEPERSON_REST_ENABLED=false
LIVEPERSON_ACCOUNT_ID=
LIVEPERSON_CLIENT_ID=
LIVEPERSON_CLIENT_SECRET=
LIVEPERSON_SKILL_ID=
LIVEPERSON_SENTINEL_DOMAIN=
LIVEPERSON_IDP_DOMAIN=
LIVEPERSON_ASYNC_MESSAGING_DOMAIN=
LIVEPERSON_MESSAGING_REST_DOMAIN=
REST_RESPONSE_IDLE_MS=3000
REST_RESPONSE_TIMEOUT_MS=60000
REST_POLL_INTERVAL_MS=750
REST_REQUEST_TIMEOUT_MS=15000
```

Set `LIVEPERSON_REST_ENABLED=true` only when opting into REST. Account ID, client ID, client secret, and the intended skill ID are required. Domain overrides are optional: leave them blank for account domain discovery, or provide the correct HTTPS service hosts for your account. Do not put credentials, JWTs, or access tokens in domain values. App JWTs and synthetic consumer JWS values are obtained by the client, not pasted into configuration, workbooks, or documentation.

Setup masks client ID/secret entry, offers to retain existing credentials without showing them, and preserves process-managed values instead of copying or shadowing them in `.env`. Cancelling credential entry writes none of the pending setup changes. The optional auth/domain check uses no browser and never automatically creates a live conversation. Enablement is configuration, not a persistent transport or session preference.

The client discovers `sentinel`, `idp`, `asyncMessagingEnt`, and `messagingRestApiDomain` once per run unless reviewed overrides are supplied. It obtains AppJWT through the Sentinel client-credentials endpoint and ConsumerJWS through IDP. Tokens stay in private in-memory caches, refresh near expiry, and get one authentication refresh after an HTTP 401. Isolated scenarios and recovery attempts use distinct synthetic consumer identities; no customer identifiers are used.

Conversation creation sends synthetic profile metadata followed by `cm.ConsumerRequestConversation` with the configured brand/skill, NORMAL TTR, and MESSAGING channel. Replies are correlated by request ID; a missing dialog ID is resolved by selecting the open MAIN dialog from conversation metadata. Text input uses `ms.PublishEvent` with `text/plain` ContentEvent. Cleanup uses `cm.UpdateConversationField` with ConversationStateField `CLOSE`; failures are reported without deleting captured results.

Polling uses the Messaging REST messages endpoint with ascending sequence order and an inclusive `newerThanSequence` cursor advanced by one. A pre-send high-water mark and publish acknowledgement prevent old history from becoming a new answer. Consumer echoes, receipts, chat-state events, internal-audience messages, and duplicate sequences do not become bot text. Multiple bot/agent text messages are joined with blank lines; each new bot message restarts the idle window. Polling stops after a settled response or the hard deadline.

Retries are bounded to three attempts for safe auth/read/close requests and HTTP 429 rejections, with exponential backoff and Retry-After support. Excessive Retry-After waits fail rather than being shortened. Conversation creation and text publication are **not** retried after ambiguous network/5xx outcomes, because the operation may already have reached LivePerson; this avoids duplicate testcase messages. A response timeout or scenario-specific rejection is recorded and isolated bulk execution continues. Permanent authentication/permission failures, invalid global configuration/protocol responses, and exhausted infrastructure retries abort safely with a checkpoint. Cleanup warnings preserve results but can require operator follow-up in LivePerson.

REST uses the configured executed-workbook path and existing selection/skip/rerun rules, not a second testcase format or runner. Use an intentional output path or the existing fresh/rerun workflow when comparing channels; a prior captured result is not automatically re-executed just because the transport changes. `npm run test:pgn:fresh -- --transport=rest` prepares workbook results without modifying WhatsApp profiles or contacting LivePerson. REST Run Configuration records transport, session mode, and polling timings; transcript rows record conversation/dialog IDs but never tokens. REST evidence is marked not applicable and is excluded from screenshot migration.

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

Start, resume, running, completion, failure, and interruption cards include **Transport: WhatsApp/REST**, **Session Mode: Isolated/Continuous**, and **Context isolation: Enabled/Disabled**. Final and failed/interrupted cards include session reset attempts when supplied by the WhatsApp runner; REST never performs debug resets. This context is retained even with start notifications disabled; it adds fields to the existing cards, not extra posts. Full/Ready-for-Retest/Discord Demo remains a separate run classification.

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

Drive evidence applies to WhatsApp execution and is disabled by default. It continues to use Google Drive API v3, Shared Drive support, inherited permissions, and one evidence subfolder per run. REST execution has no screenshots and does not require or upload Drive evidence.

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

`Test Case Knowledge Base` reads and writes fields by their headers, not fixed Excel letters. A blank Test Case ID row with a populated Turn continues the preceding scenario.

`Negative Case` independently resolves `User Input / Test Steps` and its result fields. Explicit `Turn 1:`, `Turn 2:`, and later markers inside one cell execute sequentially in the same scenario context. Their responses and timings are combined with turn labels.

The executed copy adds `Execution Transcript` with per-turn user and bot messages, timestamps, first-response timing, total timing, technical status, and evidence paths. Every bot bubble has its own transcript row. Multiple bubbles are labeled `Message 1:`, `Message 2:`, and later in the logical Bot Response cell. New transcript rows can include optional `Session Mode` and `Transport` columns; legacy transcripts without them remain readable. The workbook is atomically replaced after every attempted turn.

The new `Run Configuration` metadata sheet records per-run execution context, including Run ID, Session Mode, and Transport, separately from evidence metadata and the full/retest classification. New runs persist `sessionMode` and `transport` in their recovery state and manifest as well. Missing context on legacy runs means isolated WhatsApp execution; the metadata does not create a persistent default for later runs.

Evidence hyperlinks use the detected `Evidence` or `Evidence URL` header. When absent, preparing the executed workbook appends the Evidence column beyond all used columns and Excel tables; validation alone never adds it. `Execution Transcript` resolves `Evidence Path`, `Evidence URL`, and `Evidence Status` from its controlled schema. `Execution Metadata` stores run-folder and Drive file IDs in separate run/file groups without storing credentials.

ExcelJS misreads this source workbook's table defaults and otherwise emits invalid AutoFilter metadata. Each atomic save therefore restores every immutable source table definition byte-for-byte into the table targets generated by ExcelJS, validates the table relationships and content types, and only then replaces the executed workbook. Resume is refused if the executed workbook's table structure or source-owned input cells no longer match the source.

### Workbook Mapping

Workbook columns may be inserted or reordered. Headers define meaning; letters only describe their current location. For example, inserting a Reviewer column may move `User Input` from H to I and `Bot Response` from I to J without changing the runtime field identities.

```bash
npm run test:pgn:validate
npm run workbook:schema
npm run workbook:schema -- --review
```

The current workbook format requires headers in row 1. Matching tolerates case, leading/trailing whitespace, repeated spaces, and newlines. Conservative aliases include `User Question`, `Test Input`, and `User Message` for User Input; `Actual Bot Response`, `Actual Response`, and `Bot Answer` for Bot Response. `Negative Case` also accepts `User Input` and `Test Steps` for its combined input field. There are no fuzzy matches or vague Input/Output/Result aliases.

Test Case ID, User Input, Bot Response, and Status are required on each main sheet; the Knowledge Base sheet also requires Turn. Missing optional description/reporting fields produce warnings; absent Response Time, Test Date, or Notes are not written. Existing Evidence links, formulas, formatting, widths, filters, and table definitions are retained. Internally generated transcript/history/metadata sheets are resolved centrally without asking operators to map them; malformed internal schemas fail safely. Execution Metadata retains its two controlled groups, delimited by Evidence Key, to distinguish their duplicate Run ID headers. Evaluation Summary and other unrelated supporting sheets are not remapped or recreated.

For an ambiguous or unknown header, open `npm run pgn` -> **Workbook** -> **Review column mapping**. Setup also offers mapping review. Both canonical and alias candidates appearing together require a choice; identical duplicate headers must be renamed in Excel because a positional override would be unsafe. Accepted overrides are workbook-specific, never global aliases.

Approved mappings are stored in the gitignored repository-local `.workbook-mappings.json`, keyed by workbook path, sheet, semantic field, and selected header. Saved letters are display-only cache values: every load resolves current headers again. Validation/re-detection can refresh an already approved, unambiguous cache and report moved fields; first-time review requires confirmation. Schema status and doctor are read-only. Main-menu execution offers review when mappings are unapproved or invalid; direct non-interactive execution requires a valid schema but cannot resolve ambiguity interactively.

Recovery content hashes remain separate from mapping fingerprints. New recovery manifests record column-schema fingerprints; moved/renamed columns during a recoverable run block resume rather than silently remapping partial work. Older checkpoints retain formatting-only drift warnings when their original canonical field layout and selected inputs can still be verified; changed layouts require conservative recovery review. Mapping review is blocked while a real active/recoverable run exists. A fresh output path may be needed if an older executed workbook's table structure no longer matches an edited source.

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

WhatsApp PGN execution waits for response settlement, scrolls the active conversation to the bottom, and captures a Playwright `Locator.screenshot()` of the visible conversation pane selected from `#main` or the semantic conversation wrapper. The full WhatsApp page, navigation rail, chat list, search area, and unrelated contacts are excluded. Full-page screenshots remain available only for local failure diagnostics and are not uploaded as normal evidence. REST writes results and transcripts without screenshots; it does not fabricate evidence images or use browser evidence backfill.

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

## Interrupted Runs

Open `npm run pgn` after an interruption to inspect the recoverable run before starting new work. For isolated runs, the recovery menu offers validation, explicit resume confirmation, scenario skip, reconciliation of conflicting progress, and abandonment while preserving history. Atomic checkpoints, process locks, and heartbeats prevent concurrent runners from overwriting progress.

Isolated resume preserves the Run ID, completed scenarios, original selection, and recorded transport. An incomplete multi-turn scenario restarts from Turn 1: after a session reset for WhatsApp, or in a **new conversation for REST**. It never resumes midway through an old conversation. WhatsApp recovery reuses the existing Drive folder; REST recovery has no browser or Drive requirement. Source-content and column-schema drift are checked before execution, and unsafe structural changes block resume.

**Continuous runs cannot resume mid-stream.** After an interruption, the original shared conversation context may no longer be reliable. The same recovery menu instead offers **Inspect details**, **Restart continuous run from beginning**, **Abandon**, **Main menu**, and **Exit**. Partial resume, skip, and progress repair are disabled. Any continuous demo encountered here is preview-only and can never reach a live executor.

A full continuous restart validates the original inputs and requires full-restart readiness, explicit acceptance of formatting-only source drift when present, and a strong default-No execution confirmation. It reruns **all originally selected scenarios, in their original order**, including previously completed or skipped scenarios, and preserves the original full/retest, transport, and session modes. It creates a **new Run ID**, plus a new evidence/Drive folder for WhatsApp or **one new conversation for REST**, not a continuation of the old conversation. Only after the new checkpoint exists is the old run marked `ABANDONED` with a link to its replacement; old checkpoints, workbook history, transcripts, and evidence remain preserved.

The direct equivalents are:

```bash
npm run test:pgn -- --restart-run PGN-ORIGINAL-RUN-ID
npm run test:pgn:retest -- --restart-run RETEST-ORIGINAL-RUN-ID
npm run test:pgn:rest -- --restart-run REST-ORIGINAL-RUN-ID
```

`--restart-run` is for a full continuous restart, not partial recovery or a fresh selection. It cannot be combined with `--resume` or selection filters. Use `--accept-source-drift` only after explicitly reviewing a formatting-only drift warning; it does not permit changed testcase inputs or unsafe schema drift. Recovery restores the recorded session mode and transport rather than switching them. REST continuous recovery is always a total restart or abandonment, never a partial resume, skip, or debug reset.

Use `npm run test:pgn:resume:validate` to inspect readiness without sending testcase messages. For real runs this can check Drive folder access; demo runs skip external checks. If no interrupted run exists, the command reports `No recoverable PGN run was found`. Recovery state is stored locally in the gitignored `.runtime/` directory and is not included in release archives.

## Recovery Demo

Create a safe local interrupted-run fixture, then use the same recovery menu as a real run:

```bash
npm run recovery:demo
npm run pgn
npm run test:pgn:resume:validate
npm run recovery:demo:reset
```

The menu shows **Recoverable run found [DEMO]**. Resume and **Restart interrupted scenario** display validation, completed/remaining counts, and a Turn 1 restart preview without executing or changing progress. A confirmed skip uses the normal skip action on the demo checkpoint, then previews the remaining work; cancellation changes nothing. Abandon uses the normal action and preserves demo artifacts and recovery history.

Demo workbooks are isolated in `.runtime/pgn/demos/<Run ID>/`; checkpoints use the normal `.runtime/pgn/runs/<Run ID>/` history. Demo generation, validation, and recovery previews do not launch WhatsApp or Playwright, contact Drive or Discord, or read/write the real workbooks. Main-menu actions outside demo recovery remain real, including WhatsApp, evidence, and notification actions.

Use `npm run recovery:demo -- --mode=retest` for a retest fixture, `--source-drift` for formatting-only source drift, or `--mismatch` for conflicting recovery progress. Options may be combined. Duplicate demo generation refuses to overwrite an existing demo; run `npm run recovery:demo:reset` before generating another variant. Reset removes only demo fixtures/history, not real runs, real workbooks, credentials, or the WhatsApp profile.

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

Before executing each selected scenario, the runner appends an idempotent snapshot to `Retest History`. Positive multi-turn scenarios receive one history row per result row. The snapshot retains the previous transcript Run ID, semantic Status, Bot Response, Response Time, Test Date, and Evidence URL. New results and evidence are written back to the same history row as the retest progresses, while the active report points to the latest execution.

Every batch receives a new ID such as `RETEST-20260902T053000Z`. The ID is used by `Execution Transcript` and `Retest Metadata`; WhatsApp also uses evidence metadata and a dedicated Drive folder such as `PGN-WhatsApp-Retest-20260902T053000Z`. Previous Drive files and folders are never deleted or replaced by another run. REST retests use the same workbook/history engine without screenshots or Drive folders.

After every turn, transcript, evidence metadata, active results, history, and resume state are atomically saved. When all turns in a scenario have technical status `CAPTURED`, its semantic Status becomes `Pending Evaluation`; automation never assigns `Passed` or `Failed`. A timeout or send/chat error keeps the previous semantic Status and records the technical failure separately.

Useful retest filters are:

```bash
npm run test:pgn:retest -- --limit 3
npm run test:pgn:retest -- --test PGN-KB-075
npm run test:pgn:retest -- --resume RETEST-20260902T053000Z
```

`--test` explicitly selects only the named scenario and prints a warning if its current Status is not `Ready for Re-test`. For isolated retests, `--resume` reloads the immutable selected-ID set from `Retest Metadata`, skips scenarios already completed successfully in that same run, retries prior technical failures, reuses its Drive folder, and deduplicates history rows. If all scenarios were saved but final session cleanup failed, isolated resume retries that cleanup before marking the run complete. `--resume` cannot be combined with `--limit`, `--sheet`, `--test`, or `--rerun`: recovery preserves the original selection snapshot. If a new selection is empty, the retest command exits successfully before Drive setup or WhatsApp startup.

Fresh retest runs default to isolated even when the previous full run or retest used continuous mode. Continuous retesting is opt-in for each new run via the selector or `--session=continuous`/`--fast`. Use the same mode when validating and executing:

```bash
npm run test:pgn:retest:validate -- --session=continuous
npm run test:pgn:retest -- --session=continuous
```

Continuous retests share context across selected scenarios, perform no final cleanup reset, and follow the full-restart-only recovery policy above. Retest approval, history, evidence requirements, and semantic evaluation rules are otherwise unchanged.

Selected WhatsApp retests require Google Drive evidence to be configured. Parent-folder authentication and retest-folder creation are completed before WhatsApp opens; invalid or disabled Drive configuration aborts without sending a testcase message. Per-file upload failures after startup remain non-fatal and are recorded separately from chatbot technical status. REST retests instead require REST configuration/authentication, never Drive or browser prerequisites.

## Response Completion

The WhatsApp transport identifies response ownership from the confirmed outgoing message ID, excludes pre-existing and outgoing messages, and captures every new incoming bubble in DOM order. After the first bubble, every new or updated bubble restarts `WHATSAPP_RESPONSE_IDLE_MS`, which defaults to 10000 ms. `WHATSAPP_RESPONSE_TIMEOUT_MS` is the hard response limit and defaults to 60000 ms. A visible `typing` or `mengetik` state holds and then restarts the quiet timer, but typing detection is only an additional signal. REST uses conversation-scoped polling and its own response/request timing configuration, not DOM or typing detection.

`firstResponseMs` ends at the first captured bubble. `totalResponseMs` ends at the last captured bubble and excludes the final idle confirmation period. Each turn in a multi-turn scenario independently waits for complete response settlement before the next turn is sent.

## Session Modes

Transport and Session Mode are independent of the separate full/retest run classification. Session Mode controls WhatsApp reset boundaries or REST conversation boundaries, not response timing. The default is **Isolated** for both transports. Select the mode per invocation, not through `.env` or another persistent setting:

```bash
npm run test:pgn:validate -- --session=isolated
npm run test:pgn -- --session isolated
npm run test:pgn:validate -- --session=continuous
npm run test:pgn -- --session continuous
npm run test:pgn -- --fast
```

`--session=isolated` and `--session isolated` are equivalent; likewise for continuous. `--fast` is an alias for `--session=continuous`, not a shortcut around reset confirmation, response settlement, or evidence checks. Conflicting session flags are rejected. `test:pgn:validate` accepts only these session-related flags, rejects execution/recovery filters, and only describes the selected policy; it never opens WhatsApp or sends a reset.

The WhatsApp policy is:

| Policy | Isolated (Default) | Continuous (Opt-In) |
| --- | --- | --- |
| Initial reset and quiet drain | Required | Required |
| Between-scenario reset | Enabled | Disabled |
| Final cleanup reset and drain | Enabled | Disabled |
| Context across scenarios | Independent after reset | Shared across the whole selection |
| Interrupted-run recovery | Safe scenario-level resume | Full restart under a new Run ID or abandon |

**Continuous Session Mode warning:** exactly one clean initial reset is required before the first scenario. There are no resets between scenarios and **no final reset**. All scenarios share the same bot conversation, so previous testcase context may influence later responses. Results are context-dependent, not independent testcase outcomes. This mode is useful for rapid development checks, exploratory testing, and context-stress testing, but is not recommended as the only final acceptance run.

The REST policy is:

| Policy | Isolated (Default) | Continuous (Opt-In) |
| --- | --- | --- |
| Conversation creation | Fresh conversation per scenario | One fresh initial conversation per run |
| Multi-turn scenario | All turns use that scenario's conversation | All turns use the run's conversation |
| Context across scenarios | Independent conversations | Shared across the whole selection |
| Debug `reset` commands | Never | Never |
| Interrupted-run recovery | New conversation; interrupted scenario starts at Turn 1 | Total restart with a new Run ID/conversation or abandon |

**REST continuous warning:** one fresh initial conversation is created; all selected scenarios share it. Previous testcase context may influence later responses, so results are not independent testcase outcomes. There is no initial, between-scenario, or final debug reset. REST isolation comes from new conversations, not the Conversation Builder debug/reset facility.

### Reset Safety

This section applies only to WhatsApp. In isolated mode, independent scenarios are isolated with the deployed Conversation Builder debug command `reset`. Before every runnable scenario, including the first remaining scenario after isolated resume, the runner snapshots WhatsApp, sends `reset`, and waits only for a new incoming response containing `Session deleted`. Continuous mode uses this same contract exactly once at the beginning. `PGN_RESET_COMMAND`, `PGN_RESET_CONFIRMATION`, and `PGN_RESET_TIMEOUT_MS` configure the WhatsApp reset contract, not the session mode, and default to `reset`, `Session deleted`, and 30000 ms. REST never sends this debug command, including during recovery.

After reset confirmation, the runner requires `POST_RESET_QUIET_MS` of silence, defaulting to 10000 ms. Any new or changed incoming message is recorded as `STALE_BOT` and restarts that timer, so it cannot be assigned to the next testcase. A visible typing state also holds the drain and restarts the quiet timer when it clears, but message arrival remains authoritative. Resets occur outside the scenario turn loop, so both modes retain context within a multi-turn scenario. In isolated mode, the next reset is attempted only after the completed scenario has been written and atomically saved, and a final reset and drain run after the last selected scenario. Continuous mode retains the initial reset/drain safety but skips all between-scenario and final cleanup resets.

Reset traffic never enters User Input, Bot Response, expected handling, or semantic Status cells. It is recorded in `Execution Transcript` as `CONTROL_USER`, `CONTROL_BOT`, `CONTROL_SYSTEM`, or `STALE_BOT`.

If the expected confirmation is not captured before the reset timeout, the runner saves a reset-failure screenshot and diagnostics, records the failed control attempt, saves the executed workbook, aborts all remaining scenarios, and exits non-zero. This fail-safe prevents results from being collected under uncertain bot context.

WhatsApp isolation depends on the deployed bot continuing to allow the Conversation Builder reset/debug command. LivePerson recommends disabling debug commands in production, so this deployment capability must remain enabled for WhatsApp QA. REST isolation does not depend on debug commands.

`CAPTURED`, `TIMEOUT`, `SEND_ERROR`, and `CHAT_ERROR` are technical states only. They are written to the transcript and technical failures are appended to Notes without replacing existing notes. They never produce a semantic Passed or Failed decision.
