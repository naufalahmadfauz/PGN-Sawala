# PGN Sawala v1.0.0

PGN Sawala v1.0.0 is the first stable release of the automated PGN WhatsApp QA runner. It brings workbook-driven execution, evidence collection, guided operations, and interruption recovery into one workflow.

Released on **2026-09-06**.

## Highlights

### Automated WhatsApp QA

Execute workbook-defined Knowledge Base and Negative Case scenarios through the real WhatsApp Web channel. Full and filtered runs support multi-turn conversations, collect multiple response bubbles, and progressively save technical results and an Execution Transcript. Session resets isolate independent scenarios without breaking context between turns of the same scenario.

### Operator Experience

Start with `npm run setup`, operate through `npm run pgn`, and inspect prerequisites with `npm run doctor`. Guided configuration covers the test target, browser/session readiness, Google Drive, optional Discord notifications, and workbook mapping. Browser-launch handling supports Windows, macOS, Linux, and GitHub Codespaces, including Xvfb where needed.

### Evidence Pipeline

Capture conversation-pane screenshots rather than full WhatsApp pages. Keep evidence locally or upload it to run-specific Google Drive folders using a service account, with evidence hyperlinks in the executed workbook. The migration workflow cleans and backfills older evidence while preserving report results and source table metadata.

### Retesting

Select cases marked `Ready for Re-test`, preserve their prior results in run-specific Retest History, and move successfully captured cases to `Pending Evaluation`. Recovery retains the original selected-ID snapshot, reuses the same run and evidence folder, and avoids selecting unrelated cases whose status changes later.

### Discord Notifications

Optional lifecycle notifications report started, progress, completed, interrupted/failed, and resumed runs. Notification failures are fail-open: telemetry never controls testcase execution. Dedicated validation and simulated-notification commands help operators check their configuration without running PGN testcases.

### Resilient Runs

Interrupted-run discovery, atomic checkpoints, process locks, and heartbeats protect durable progress. Resume keeps the original Run ID, preserves completed scenarios, and restarts an incomplete multi-turn scenario from Turn 1. Source-content and schema-drift checks block unsafe recovery, while conflicting artifacts require explicit reconciliation.

Local recovery demos exercise the same discovery, validation, and operator menu using synthetic workbooks. Demo execution is blocked before WhatsApp, Playwright, Drive, or Discord activity.

### Flexible Workbook Layouts

Headers define semantic fields, not fixed Excel letters. Knowledge Base and Negative Case mappings are resolved independently, so columns can move without changing runtime field identities. Conservative aliases and workbook-specific overrides support controlled header changes; missing required or ambiguous fields fail safely and require review. There are no silent reads or writes to the historical H/I/J positions.

## Main Commands

| Command | Purpose |
| --- | --- |
| `npm run setup` | Guided environment and workbook configuration. |
| `npm run pgn` | Interactive operator control panel and recovery workflow. |
| `npm run doctor` | Prerequisite diagnostics; configured Drive access may be checked. |
| `npm run test:pgn` | Execute real PGN WhatsApp testcases from the workbook. |
| `npm run test:pgn:retest` | Execute the approved Ready-for-Retest selection. |
| `npm run discord:validate` | Inspect the configured Discord webhook without posting a notification. |
| `npm run discord:demo` | Send a simulated lifecycle to the configured Discord webhook. |
| `npm run recovery:demo` | Create a synthetic, local-only interrupted-run fixture. |
| `npm run test:pgn:resume:validate` | Inspect recovery readiness without sending testcase messages. |
| `npm run workbook:schema` | Inspect detected workbook mappings without changing workbook files. |

The direct test commands send real messages. `discord:demo` sends real Discord notifications, unlike `recovery:demo`, which is local-only. Use the interactive control panel when you want guided confirmation and prerequisite review. Real recovery validation may check access to an existing Drive folder; demo recovery skips all external checks.

## Installation

Use Node.js 20.12 or newer and npm. For a checkout pinned to this release:

```bash
git clone --branch v1.0.0 --depth 1 https://github.com/naufalahmadfauz/PGN-Sawala.git
cd PGN-Sawala
npm install
npm run setup
npm run pgn
```

Setup guides browser installation/authentication and local configuration. Keep `.env`, service-account credentials, and the WhatsApp profile private. For repeatable dependency installation in an existing release checkout, use `npm ci` with the included lockfile.

See the [README](https://github.com/naufalahmadfauz/PGN-Sawala/blob/v1.0.0/README.md) for full-run preparation, retesting, evidence configuration, recovery, and schema review. GitHub's source ZIP/tarball contains the tagged repository; no local runtime bundles or evidence archives are attached to this release.

## Known Limitations

- Technical response capture is not semantic correctness evaluation. Passed/Failed assessment remains a separate QA responsibility.
- Test coverage comes from the supplied workbook scenarios; automatic paraphrase generation and broader robustness expansion are not part of V1.
- Supported main sheets require recognizable headers in row 1. Ambiguous or unknown mappings need operator review; incompatible source/executed table structures can require a fresh output path.
- Recovery intentionally rejects structural input or schema drift instead of guessing how to continue partial conversations.
- Live execution depends on WhatsApp authentication, the deployed bot's reset capability, and valid external-service configuration. Selected retests require Google Drive evidence configuration.
- WhatsApp Web UI changes and external-service availability can affect live execution. Offline release verification is not a new live integration certification.

## Validation

Release verification on Linux/Codespaces with Node.js 24.14.0 and npm 11.9.0 completed successfully:

- TypeScript: `npm run check` passed.
- **233 offline regression tests passed across ten npm test commands, with zero failures or skips.** Coverage includes configuration, workbook persistence, response collection, session isolation, retest workflows, evidence migration, Discord notifications, operator UX, recovery/demo safety, and workbook schema mapping.
- Recovery coverage: **78 tests passed**. Workbook-schema coverage: **39 tests passed**. Discord coverage: **33 tests passed using mocked HTTP**, not a live webhook.
- The full retest suite includes a zero-candidate CLI check that exited before WhatsApp startup.
- Workbook validation passed with expected warnings for the optional Evidence headers absent from the immutable source template; the executed workbook's Evidence headers resolved successfully.
- Tracked-file and XLSX-content secret scans found no likely real secrets or unexpected private/runtime artifacts. The local recovery demo was cleaned through the official reset command.
- Hash verification confirmed **1,480 protected real files unchanged** after cleanup and testing.

No real PGN WhatsApp testcase, evidence upload, or Discord test notification was triggered for this release. No product behavior was changed during release preparation.
