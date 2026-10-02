# PGN Sawala

Workbook-driven QA harness for the externally hosted PGN chatbot. The main workbook execution path is `src/pgn-runner.ts`.

## Codebase context

- **Architecture:** before changing execution, workbook persistence, recovery, or integrations, read [docs/architecture.md](docs/architecture.md) for the owning modules and cross-module contracts.
- **Development:** before choosing verification commands or adding tests, read [docs/development.md](docs/development.md). Code regression tests live in `scripts/*.test.ts`; similarly named operator commands can perform real chatbot runs.
- **Operation:** for installation, configuration, and operator procedures, use [README.md](README.md).
- **Continuation:** when finishing or transferring a task, follow [docs/development.md#continuing-work](docs/development.md#continuing-work) so the next agent has the current task state and verification evidence.

## Agent skills

### Issue tracker

Use GitHub Issues in `naufalahmadfauz/PGN-Sawala` for specs, tickets, and triage. Before tracker operations, read `docs/agents/issue-tracker.md`.

### Triage labels

Use the five default triage labels. Before triaging or changing triage labels, read `docs/agents/triage-labels.md`.

### Domain docs

Single-context: root `GLOSSARY.md` and `docs/adr/`. Before exploring code or proposing changes, read `docs/agents/domain.md`.
