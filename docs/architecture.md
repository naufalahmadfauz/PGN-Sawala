# Architecture

PGN Sawala owns workbook-driven QA execution, response capture, evidence, and recovery for an externally hosted chatbot. Domain terms are defined in [GLOSSARY.md](../GLOSSARY.md); operator procedures are in [README.md](../README.md).

## Execution path

```text
operator menu ----+
                  +--> shared runner --> WhatsApp or REST
direct CLI -------+          |
                             +--> workbook + recovery state
```

The menu starts at `scripts/pgn.ts` and dispatches through `src/operator/control-panel.ts` and `default-actions.ts`. Direct workbook commands start at `scripts/run-pgn*.ts` or `scripts/retest-pgn*.ts`.

[runPgnWorkbook](../src/pgn-runner.ts) coordinates the lifecycle:

1. Parse invocation options; discover existing recovery state and restore the recorded execution context when recovering.
2. Acquire the project run lock and output-workbook lock; load and validate the source and executed workbooks.
3. Select scenarios and record the selection in a recovery manifest/checkpoint. Save preparation metadata before opening a conversation or sending testcase messages.
4. Initialize the chosen transport. For each scenario, record its attempt, snapshot prior retest results when applicable, and call `beginScenario`.
5. Send turns sequentially. After each returned turn result, update transcript/results/history, save the workbook, then update checkpoint progress. A non-`CAPTURED` turn stops the remaining turns in that scenario.
6. Record the scenario outcome, call `endScenario`, then finalize/close the transport and save the final run audit. Failure and interruption paths retain recovery state; locks are released on exit.

The lifecycle interface is [TestTransport](../src/transports/test-transport.ts). Selection, report writing, retest history, and recovery remain shared across transports.

## Where to change behavior

| Concern | Starting points |
| --- | --- |
| Menus, confirmation, setup, and diagnostics | [control-panel.ts](../src/operator/control-panel.ts), [default-actions.ts](../src/operator/default-actions.ts), [setup.ts](../src/operator/setup.ts), [diagnostics.ts](../src/operator/diagnostics.ts) |
| CLI options and ordinary selection | [pgn-cli.ts](../src/pgn-cli.ts), [pgn-selection.ts](../src/pgn-selection.ts), [pgn-preflight.ts](../src/operator/pgn-preflight.ts) |
| Transport/session context and persisted defaults | [session-mode.ts](../src/session-mode.ts), [run-configuration.ts](../src/excel/run-configuration.ts), [run-state.ts](../src/recovery/run-state.ts) |
| Scenario parsing and header resolution | [pgn-workbook-loader.ts](../src/excel/pgn-workbook-loader.ts), [multi-turn-parser.ts](../src/excel/multi-turn-parser.ts), [workbook-schema.ts](../src/excel/workbook-schema.ts), [workbook-mapping.ts](../src/excel/workbook-mapping.ts) |
| Result persistence and retest policy | [pgn-workbook-writer.ts](../src/excel/pgn-workbook-writer.ts), [retest-selection.ts](../src/retest/retest-selection.ts), [retest-workbook.ts](../src/excel/retest-workbook.ts) |
| Interrupted runs and reconciliation | [run-state.ts](../src/recovery/run-state.ts), [recovery-service.ts](../src/recovery/recovery-service.ts), [demo-safety.ts](../src/recovery/demo-safety.ts) |
| WhatsApp message ownership, reset, and browser behavior | [WhatsApp transport](../src/transports/whatsapp.ts), [client.ts](../src/whatsapp/client.ts), [response-collector.ts](../src/whatsapp/response-collector.ts), [session-reset.ts](../src/whatsapp/session-reset.ts), [browser-runtime.ts](../src/operator/browser-runtime.ts) |
| REST conversations and requests | [REST transport](../src/transports/rest.ts), [liveperson-client.ts](../src/rest/liveperson-client.ts), [errors.ts](../src/rest/errors.ts) |
| Evidence and notification side effects | [evidence-migration.ts](../src/evidence/evidence-migration.ts), [google-drive.ts](../src/evidence/google-drive.ts), [discord.ts](../src/notifications/discord.ts) |
| Configuration and environment precedence | [config.ts](../src/config.ts), [environment.ts](../src/environment.ts), [REST config](../src/rest/config.ts) |

`test:single` follows a separate path: [scripts/test-single.ts](../scripts/test-single.ts) calls [src/runner.ts](../src/runner.ts), which uses the simpler `src/excel/loader.ts` and `src/excel/writer.ts`. Workbook-runner changes belong in the PGN path above; the single-message path has its own result model.

## Workbook persistence

- **Test-case sheet availability:** a workbook must contain `Test Case Knowledge Base`, `Negative Case`, or both. `parsePgnWorkbook` validates every present sheet and records its presence in `availableSheets` independently of scenario counts. `assertPgnWorkbookValid` also rejects an explicitly requested absent category at preflight, execution, and retest validation boundaries. Mapping and evidence preparation operate on present test-case sheets; executed copies retain the source's sheet set. Header-only recognized sheets are valid and follow normal empty-selection behavior.
- **Schema-aware access:** main-sheet fields are resolved independently by header through `workbook-schema.ts`. Use `fieldCell`/`optionalFieldCell` and the resolved schema when changing reads or writes. Mapping overrides are workbook-specific; the mapping-review workflow checks for active/recoverable runs before changing them. See [Workbook Mapping](../README.md#workbook-mapping).
- **Source preservation:** `openExecutedPgnWorkbook` creates or validates the results copy, checks source-owned cells, and retains the source's table definitions. Execution writes the copy; prepared inputs and expected responses remain source-owned.
- **Excel compatibility:** `saveExecutedPgnWorkbook` writes a temporary XLSX, restores source table XML into the generated table targets, verifies table topology, then renames the file. This preserves the workaround for the ExcelJS table/AutoFilter behavior described in [Real PGN Workbook](../README.md#real-pgn-workbook). A direct `workbook.xlsx.writeFile` bypasses these guarantees.
- **Concurrent edits:** workbook locks and expected-output hashes protect the executed workbook. The saver checks for external changes before writing and before replacement.

The workbook and JSON checkpoint are independently persisted artifacts. Result progress advances after the workbook save; they do not form one cross-file transaction. `recovery-service.ts` reconciles checkpoint, workbook, transcript, and evidence when they disagree.

## Results and recovery

- Technical, semantic, and evidence statuses have distinct owners. Full execution preserves semantic evaluation; `applyRetestStatusTransition` changes a fully captured retest to `Pending Evaluation`. Passed/Failed decisions remain with the evaluator.
- Ordinary selection skips existing complete or partial responses unless explicitly rerun. Its completion checks in `pgn-workbook-loader.ts` differ from checkpoint completion.
- In **full** execution, an attempted scenario with a saved technical failure can count as completed for recovery. In **retest** execution, automatic completion requires every turn to be captured; technical failures remain recoverable. A `COMPLETED` run therefore does not assert that the bot passed evaluation.
- Recovery retains the original selection and execution context. Isolated recovery restarts an incomplete scenario from Turn 1; continuous recovery replays the full original selection under a new Run ID. See [Interrupted Runs](../README.md#interrupted-runs).
- Older persisted records with absent transport/session fields resolve to WhatsApp/isolated through `session-mode.ts`. Preserve that compatibility when extending metadata readers.

The executed workbook holds transcript, run configuration, evidence metadata, and retest history. Local checkpoints/manifests live under `.runtime/pgn/runs/<Run ID>/`; `run-state.ts` also manages the active-run pointer and process lock. Recovery demos use separately scoped workbooks and execution guards in `demo-safety.ts` and `run-state.ts`.

## Transport and integration boundaries

**WhatsApp:** the adapter owns browser/client startup, reset boundaries, response collection, screenshot capture, and optional evidence upload. It also writes reset/drain control transcript records through the shared workbook writer. Keep this existing workbook coupling in mind when changing the adapter. Response ownership follows the confirmed outgoing message; delayed/reset traffic is handled by the collector and reset/drain logic. The exact policy is in [Response Completion](../README.md#response-completion) and [Session Modes](../README.md#session-modes).

**REST:** `LivePersonClient` owns discovery, authentication, request deadlines/retries, and protocol parsing. `RestTransport` owns conversation boundaries, the pre-send response boundary, sequence deduplication, and response settlement. Isolation uses fresh synthetic consumer/conversation identities. Polling uses an inclusive sequence cursor; local deduplication removes previously seen events. Ambiguous send/create outcomes are not retried because the request may already have taken effect. Preserve these contracts when changing [LivePerson REST](../README.md#liveperson-rest) behavior.

**Evidence and notifications:** Drive initialization is checked before WhatsApp execution when enabled; WhatsApp retests require it. Per-file upload failure is recorded independently of response capture. REST marks screenshot evidence not applicable. Discord delivery is fail-open and carries operational metadata; notification failures do not determine scenario results.

**Configuration:** `loadEnvironment` resolves `.env` from the repository root and gives pre-existing process values precedence. `loadConfig({ transport: "rest" })` excludes irrelevant WhatsApp/Drive settings from REST-only validation. Consult the [configuration guide](../README.md#configuration) for values; keep environment parsing and error redaction in the existing configuration/integration modules.
