# Changelog

Notable user-facing changes to PGN Sawala are documented here.

## [1.0.0] - 2026-09-06

### Added

- Automated PGN WhatsApp Web QA execution from workbook-defined Knowledge Base and Negative Case scenarios, including multi-turn conversations.
- Full-run, filtered-run, and Ready-for-Retest workflows with Execution Transcript, Retest History, and technical outcome reporting.
- Conversation-pane screenshots, local evidence, Google Drive service-account integration, workbook evidence hyperlinks, and legacy evidence migration.
- Optional Discord start, progress, completion, interruption, failure, and resume notifications, plus notification validation and demo commands.
- Interactive setup and operator control panel, `.env` configuration, and diagnostics for Windows, macOS, Linux, and GitHub Codespaces.
- Interrupted-run discovery, guided recovery, and local recovery demos with retest, source-drift, and reconciliation-mismatch variants.
- Per-sheet header-based workbook mapping, conservative aliases, workbook-specific overrides, and operator schema review.

### Reliability

- Session reset and quiet-period isolation before independent scenarios, with stale-message exclusion and multi-bubble response collection.
- Progressive workbook persistence, atomic replacement, external-write detection, and preservation of source Excel table metadata.
- Atomic recovery checkpoints, process locks, heartbeats, and graceful interruption handling.
- Same Run ID and existing evidence-folder reuse on resume; incomplete multi-turn scenarios restart from Turn 1.
- Source-content and schema-drift checks, explicit reconciliation of conflicting artifacts, and preserved original retest selection snapshots.

### Safety

- Discord delivery failures remain fail-open and cannot abort testcase execution.
- Credential values are redacted from supported error/diagnostic paths; credentials, browser profiles, generated reports, evidence, and runtime state remain excluded from Git.
- Recovery demos use synthetic local artifacts and block WhatsApp, Playwright, Drive, and Discord execution. The separate Discord demo intentionally sends simulated notifications when configured.
- Missing required headers and ambiguous mappings block execution instead of silently reading or writing historical column positions.
- Operator confirmation protects live execution and destructive workflow actions; demo reset removes only verified demo artifacts.

[1.0.0]: https://github.com/naufalahmadfauz/PGN-Sawala/releases/tag/v1.0.0
