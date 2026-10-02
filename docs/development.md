# Development and continuation

Use the current issue's acceptance criteria and the [architecture guide](architecture.md) to locate the behavior under change. [package.json](../package.json), [tsconfig.json](../tsconfig.json), and [biome.json](../biome.json) are the sources of truth for scripts, TypeScript settings, and formatting policy.

## Verification

For TypeScript changes, run `npm run check` and the regression suites covering the changed behavior. Select regression files explicitly: maintained suites are `scripts/*.test.ts`, run with Node's test runner and `tsx`. Choose their named npm commands from `package.json`.

For a cross-cutting change, all current regression files can be selected from the repository root in a POSIX shell:

```sh
node --import tsx --test scripts/*.test.ts
```

This explicit selection matters: test discovery also recognizes `test-*` filenames, and `scripts/test-single.ts` is a live WhatsApp command. The operator menu's regression action uses a separate explicit `SAFE_TEST_FILES` list in [default-actions.ts](../src/operator/default-actions.ts); update that list when adding a suite intended for the menu.

### Choose coverage by behavior

Paths in this table are relative to `scripts/`.

| Changed behavior | Regression starting points |
| --- | --- |
| Response ownership, delayed bubbles, typing/idle timing | `response-collector.test.ts`, `session-reset.test.ts` |
| Browser choice and platform launch | `whatsapp-browser.test.ts`, `operator.test.ts` |
| Shared execution lifecycle, session policy, transport context | `session-mode.test.ts`, `rest.test.ts`, `recovery.test.ts` |
| Headers, field mapping, table preservation, result writes | `workbook-schema.test.ts`, `workbook-writer.test.ts`; include `retest.test.ts` for history/status changes |
| Retest selection, snapshots, evaluation transitions | `retest.test.ts`, `session-mode.test.ts` |
| Checkpoints, drift, reconciliation, demo isolation | `recovery.test.ts`, `recovery-demo.test.ts`, `session-mode.test.ts` |
| Environment, setup, menu actions, cancellation | `config.test.ts`, `operator.test.ts`; include `rest.test.ts` for REST configuration |
| Evidence migration or notification delivery | `evidence-migration.test.ts`, `discord.test.ts`, respectively |

Workbook tests include reads of the source XLSX under `data/`; result writes use temporary copies. Regression success establishes the behavior exercised by those fixtures. Record any separately performed live verification with its transport, session mode, selection, and outcome.

### Existing test patterns

- **Time-dependent collection:** [response-collector.test.ts](../scripts/response-collector.test.ts) supplies a virtual `ResponseCollectorEnvironment`; delayed events advance a controlled clock.
- **REST requests:** [rest.test.ts](../scripts/rest.test.ts) provides a fake service and clock, including sequence-boundary, retry, cleanup, and error-redaction cases. `RestDependencies` accepts `fetch`, `now`, and `sleep`.
- **Runner/persistence integration:** [session-mode.test.ts](../scripts/session-mode.test.ts) and [workbook-schema.test.ts](../scripts/workbook-schema.test.ts) create temporary project roots, synthetic workbooks, mocked external boundaries, and source-preservation assertions.
- **Operator interaction:** [operator.test.ts](../scripts/operator.test.ts) uses scripted UI answers and injected actions/dependencies to assert dispatch and cancellation.

For configuration-sensitive fixtures, provide both a temporary `repositoryRoot` and an isolated `environment` to `loadConfig`. The loader still reads `.env` at the selected root, so an isolated environment object alone does not isolate a test from workstation configuration.

## Entry-point conventions

`scripts/` contains both executable commands and importable command handlers. Some operator actions import handlers from scripts directly. Follow [validate-rest.ts](../scripts/validate-rest.ts): export the callable function, guard CLI invocation with `isEntrypoint(import.meta.url)`, and use `runCliMain` for CLI error handling. Executable wrappers such as `run-pgn.ts` invoke their main function at module load; tests should call the underlying handler instead.

The control panel owns interactive confirmations through `OperatorUi` and `OperatorActions`. Direct execution scripts dispatch immediately. Keep menu cancellation/confirmation behavior covered when adding an action.

## Operator commands and their effects

Use the [README](../README.md) for operational procedures. Command names containing `test`, `validate`, or `demo` can represent real integration work:

| Command family | Effect |
| --- | --- |
| `test:pgn`, `test:pgn:retest`, REST execution/retest, `test:single`, `rest:smoke` | Send real messages; workbook runs also save execution state/results |
| `test:pgn:rest:validate` | Makes domain/authentication requests; creates no conversation |
| `doctor`, real-run `test:pgn:resume:validate` | May check configured Drive access |
| `evidence:validate` | Creates local previews and may check Drive access |
| `test:pgn:fresh`, `evidence:migrate` | Replace/archive generated workbook content or migrate/upload evidence |
| `discord:demo`, `discord:validate -- --send-test` | Send real Discord notifications |
| `recovery:demo` | Creates synthetic local recovery artifacts; recovery execution is preview-only |

Use the explicit regression suites above for routine code verification. Operational runs are separate task actions with their actual effects recorded. Local credentials, profiles, generated reports, evidence, and recovery state follow the exclusions in [.gitignore](../.gitignore); share redacted reproduction details and fixture data in issues.

## Production context

The maintainer confirmed that production is operated manually on a workstation. The [getting-started procedure](../README.md#getting-started) documents the supported launch workflow.

Context still to confirm with the maintainer: workstation OS, deployed commit/tag, and the actual update, backup, and rollback procedure. Record those facts here when supplied; repository release notes describe historical releases rather than identifying the workstation's installed revision.

## Continuing work

Keep task-specific progress in the active GitHub issue, using [the tracker conventions](agents/issue-tracker.md). A fresh agent starts with `AGENTS.md`, the issue, and the current working-tree diff.

When finishing a session or preparing a portable handoff, record:

```text
Task / issue and acceptance criteria:
Branch, commit, and outstanding working-tree changes:
Completed changes and relevant file/symbol pointers:
Verification: exact command -> result; regression or live scope:
Open questions, blockers, and remaining work:
Next concrete action:
```

Update durable behavior in the architecture guide, operator procedures in the README, and domain terms in the glossary alongside the change. Keep per-session logs and transient status in the issue/handoff so the entry-point documentation stays current and short.
