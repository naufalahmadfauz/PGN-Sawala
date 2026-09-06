import assert from "node:assert/strict";
import { access, mkdir, mkdtemp, readFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import test, { type TestContext } from "node:test";
import ExcelJS, { type Worksheet } from "exceljs";
import JSZip from "jszip";
import { chromium } from "playwright";
import { loadConfig } from "../src/config";
import { GoogleDriveEvidencePublisher } from "../src/evidence/google-drive";
import { discoverEvidenceInventory } from "../src/evidence/evidence-migration";
import { readEvidenceHyperlink, upsertEvidenceFileMetadata, upsertEvidenceRunMetadata, getEvidenceFileMetadata, getEvidenceRunMetadata, removeEvidenceFileMetadata } from "../src/excel/evidence-workbook";
import { KB_HEADERS, NEGATIVE_HEADERS, loadPgnWorkbook, parsePgnWorkbook } from "../src/excel/pgn-workbook-loader";
import { assertPgnWorkbookValid } from "../src/excel/pgn-workbook-validator";
import { appendRecoveryTranscriptEvent, applyScenarioExecution, applyScenarioResults, openExecutedPgnWorkbook, saveExecutedPgnWorkbook } from "../src/excel/pgn-workbook-writer";
import { KB_SHEET_NAME, NEGATIVE_SHEET_NAME, type ExecutedTurn, type PgnTestScenario } from "../src/excel/pgn-types";
import { applyRetestStatusTransition, snapshotRetestHistory, updateRetestHistory, upsertRetestRunMetadata, getRetestRunMetadata } from "../src/excel/retest-workbook";
import { EVIDENCE_FILE_SCHEMA, EVIDENCE_RUN_SCHEMA, fieldCell, fieldColumn, getWorksheetSchema, mainSchemaFingerprint, optionalFieldCell, resolveWorksheetSchema, setWorkbookSchemaOverrides } from "../src/excel/workbook-schema";
import { readWorkbookMappingStore, saveWorkbookMappingStore, workbookMappingKey, WORKBOOK_MAPPING_FILE } from "../src/excel/workbook-mapping";
import { collectDiagnostics } from "../src/operator/diagnostics";
import { runControlPanel, type OperatorActions } from "../src/operator/control-panel";
import { runSetupWizard } from "../src/operator/setup";
import type { OperatorUi } from "../src/operator/ui";
import { ensureReviewedWorkbookMapping, formatWorkbookMappings, inspectWorkbookMappings, reviewWorkbookMapping } from "../src/operator/workbook-configuration";
import { runPgnWorkbook } from "../src/pgn-runner";
import { hashFile, initializeRecoveryCheckpoint } from "../src/recovery/run-state";
import { reconcileRecoveryArtifacts, validateRecoveryRun } from "../src/recovery/recovery-service";
import { selectRetestScenarios } from "../src/retest/retest-selection";
import { WhatsAppClient } from "../src/whatsapp/client";
import { validatePgnWorkbook } from "./validate-pgn-workbook";

function fixtureWorkbook(): ExcelJS.Workbook {
  const workbook = new ExcelJS.Workbook();
  const kb = workbook.addWorksheet(KB_SHEET_NAME);
  kb.addRow(KB_HEADERS);
  kb.addRow([1, "Synthetic article", "SCHEMA-KB-001", "Operator", "Objective", "Expected", 1, "Synthetic input", null, null, null, "Ready for Re-test", "Preserve note"]);
  kb.addRow([2, "Synthetic article", "SCHEMA-KB-002", "Operator", "Objective", "Expected", 1, "First turn", null, null, null, "Ready for Re-test"]);
  kb.addRow([null, null, null, null, null, null, 2, "Second turn"]);
  const negative = workbook.addWorksheet(NEGATIVE_SHEET_NAME);
  negative.addRow(NEGATIVE_HEADERS);
  negative.addRow([1, "Synthetic category", "SCHEMA-NEG-001", "Objective", "Turn 1: First negative input\nTurn 2: Second negative input", "Negative condition", "Expected handling", null, null, null, "Ready for Re-test", "Negative note", "Reference"]);
  return workbook;
}
function executions(scenario: PgnTestScenario): ExecutedTurn[] {
  return scenario.turns.map((turn) => ({
    turn, technicalStatus: "CAPTURED", completedAt: new Date("2026-09-05T10:00:00Z"),
    botMessages: [{ sequence: 1, message: `Synthetic response ${turn.turnNumber}`, timestamp: new Date("2026-09-05T10:00:00Z") }],
    combinedResponse: `Synthetic response ${turn.turnNumber}`, totalResponseMs: 1_000,
    evidenceUrl: `https://drive.google.com/file/d/fixture-${scenario.testCaseId}-${turn.turnNumber}/view`,
    evidenceStatus: "EVIDENCE_SYNCED",
  }));
}
function reorder(sheet: Worksheet, order: number[]): void {
  const rows = sheet.getRows(1, sheet.rowCount)!.map((row) => order.map((column) => row.getCell(column).value));
  rows.forEach((values, index) => { sheet.getRow(index + 1).values = values; });
}
async function project(context: TestContext, workbook = fixtureWorkbook()) {
  const root = await mkdtemp(path.join(tmpdir(), "pgn-schema-"));
  context.after(() => rm(root, { recursive: true, force: true }));
  const forbid = async (): Promise<never> => { throw new Error("Unexpected live boundary"); };
  const guards = [
    context.mock.method(WhatsAppClient.prototype, "open", forbid),
    context.mock.method(chromium, "launchPersistentContext", forbid),
    context.mock.method(globalThis, "fetch", forbid),
    ...(["validateParentFolder", "validateRunFolder", "ensureRunFolder", "uploadPng"] as const).map((method) => context.mock.method(GoogleDriveEvidencePublisher.prototype, method, forbid)),
  ];
  context.after(() => guards.forEach((guard) => assert.equal(guard.mock.callCount(), 0)));
  const config = loadConfig({ repositoryRoot: root, environment: { GOOGLE_DRIVE_EVIDENCE_ENABLED: "false", DISCORD_NOTIFICATIONS_ENABLED: "false", WHATSAPP_HEADLESS: "true" } });
  await mkdir(path.dirname(config.pgnSourceWorkbookPath), { recursive: true });
  await mkdir(path.dirname(config.pgnExecutedWorkbookPath), { recursive: true });
  await workbook.xlsx.writeFile(config.pgnSourceWorkbookPath);
  return { root, config, workbook };
}
function scriptedUi(answers: unknown[]) {
  const events: string[] = [];
  const ui: OperatorUi = {
    intro: (message) => { events.push(message); }, outro: (message) => { events.push(message); }, cancel: (message) => { events.push(message); },
    note: (message, title) => { events.push(`${title}: ${message}`); }, info: (message) => { events.push(message); },
    success: (message) => { events.push(message); }, warn: (message) => { events.push(message); }, error: (message) => { events.push(message); },
    select: async (prompt) => { events.push(prompt.message); return answers.shift() as never; },
    confirm: async (prompt) => { events.push(prompt.message); return answers.shift() as boolean | undefined; },
    text: async () => answers.shift() as string | undefined,
    secret: async () => { throw new Error("No secret prompt expected"); },
    task: async (_, operation) => operation(),
  };
  return { ui, events, answers };
}

test("current repository workbook resolves headers without changing its bytes", async () => {
  const source = path.resolve("data/PGN AI Assistant - Knowledge Base Testing Report - User Inputs.xlsx");
  const before = await hashFile(source);
  const loaded = await loadPgnWorkbook(source);
  assertPgnWorkbookValid(loaded.parsed);
  const kb = getWorksheetSchema(loaded.workbook.getWorksheet(KB_SHEET_NAME)!);
  assert.equal(kb.fields.userInput?.columnLetter, "H");
  assert.equal(kb.fields.botResponse?.columnLetter, "I");
  assert.equal(getWorksheetSchema(loaded.workbook.getWorksheet(NEGATIVE_SHEET_NAME)!).fields.userInput?.columnLetter, "E");
  assert.equal(await hashFile(source), before);
});

for (const kind of ["insert", "swap", "reverse", "multiple"] as const) {
  test(`${kind} columns resolves independent input/output/status fields and preserves multi-turn grouping`, async (context) => {
    const workbook = fixtureWorkbook();
    const kb = workbook.getWorksheet(KB_SHEET_NAME)!;
    const negative = workbook.getWorksheet(NEGATIVE_SHEET_NAME)!;
    if (kind === "insert") {
      kb.spliceColumns(8, 0, ["Reviewer", "Keep reviewer", "Keep reviewer"]);
      assert.equal(getWorksheetSchema(kb).fields.userInput?.columnLetter, "I");
      assert.equal(getWorksheetSchema(kb).fields.botResponse?.columnLetter, "J");
    } else if (kind === "swap") reorder(kb, [1, 2, 3, 4, 5, 6, 7, 9, 8, 10, 11, 12, 13]);
    else if (kind === "reverse") reorder(kb, Array.from({ length: 13 }, (_, index) => 13 - index));
    else kb.spliceColumns(3, 0, ["Reviewer"], ["Approval Owner"]);
    negative.spliceColumns(5, 0, ["Negative Reviewer", "Keep negative reviewer"]);
    const { config } = await project(context, workbook);
    const opened = await openExecutedPgnWorkbook(config.pgnSourceWorkbookPath, config.pgnExecutedWorkbookPath);
    const parsed = opened.parsed;
    assertPgnWorkbookValid(parsed);
    assert.equal(parsed.scenarios.length, 3);
    const multi = parsed.scenarios.find((scenario) => scenario.testCaseId === "SCHEMA-KB-002")!;
    assert.deepEqual(multi.turns.map((turn) => [turn.turnNumber, turn.userInput]), [[1, "First turn"], [2, "Second turn"]]);
    for (const scenario of parsed.scenarios) {
      const values = executions(scenario);
      snapshotRetestHistory(opened.workbook, "RETEST-SCHEMA", scenario, new Date());
      applyScenarioExecution(opened.workbook, "RETEST-SCHEMA", scenario, values);
      assert.equal(applyRetestStatusTransition(opened.workbook, scenario, values), true);
      updateRetestHistory(opened.workbook, "RETEST-SCHEMA", scenario, values);
    }
    await saveExecutedPgnWorkbook(opened.workbook, config.pgnExecutedWorkbookPath);
    const reloaded = await loadPgnWorkbook(config.pgnExecutedWorkbookPath);
    for (const scenario of reloaded.parsed.scenarios) {
      const sheet = reloaded.workbook.getWorksheet(scenario.sheetName)!;
      assert.match(fieldCell(sheet, scenario.sourceRowNumber, "botResponse").text, /Synthetic response/);
      assert.equal(fieldCell(sheet, scenario.sourceRowNumber, "status").text, "Pending Evaluation");
      assert(readEvidenceHyperlink(fieldCell(sheet, scenario.turns.at(-1)!.rowNumber, "evidence")));
    }
    assert.equal(fieldCell(reloaded.workbook.getWorksheet(KB_SHEET_NAME)!, 2, "notes").text, "Preserve note");
    assert.equal(negative.getCell(2, 5).text, "Keep negative reviewer");
    if (kind === "insert") assert.equal(reloaded.workbook.getWorksheet(KB_SHEET_NAME)!.getCell(2, 8).text, "Keep reviewer");
    assert.equal(selectRetestScenarios(reloaded.parsed.scenarios).selected.length, 0);
  });
}

for (const header of ["user input", " User Input ", "User   Input", "User\nInput", "User Question", "Test Input", "User Message"]) {
  test(`canonical normalization or approved alias resolves ${JSON.stringify(header)}`, () => {
    const workbook = fixtureWorkbook();
    const sheet = workbook.getWorksheet(KB_SHEET_NAME)!;
    sheet.getCell(1, 8).value = header;
    const mapping = resolveWorksheetSchema(sheet);
    assert.equal(mapping.valid, true);
    assert.equal(mapping.fields.userInput?.columnIndex, 8);
    assert.equal(parsePgnWorkbook(workbook).scenarios[0].turns[0].userInput, "Synthetic input");
  });
}

for (const header of ["Actual Bot Response", "Actual Response", "Bot Answer"]) {
  test(`approved response alias ${header} is writable`, () => {
    const workbook = fixtureWorkbook();
    const sheet = workbook.getWorksheet(KB_SHEET_NAME)!;
    sheet.getCell(1, 9).value = header;
    const scenario = parsePgnWorkbook(workbook).scenarios[0];
    applyScenarioResults(workbook, scenario, executions(scenario));
    assert.equal(sheet.getCell(2, 9).text, "Synthetic response 1");
  });
}

for (const fault of ["unknown", "ambiguous", "missing-response", "missing-status", "missing-turn", "header-row"] as const) {
  test(`${fault} fails before workbook writes, WhatsApp, Playwright, or notifications`, async (context) => {
    const workbook = fixtureWorkbook();
    const sheet = workbook.getWorksheet(KB_SHEET_NAME)!;
    if (fault === "unknown") sheet.getCell(1, 8).value = "Customer Prompt";
    if (fault === "ambiguous") sheet.getCell(1, 4).value = "User Question";
    if (fault === "missing-response") sheet.spliceColumns(9, 1);
    if (fault === "missing-status") sheet.spliceColumns(12, 1);
    if (fault === "missing-turn") sheet.spliceColumns(7, 1);
    if (fault === "header-row") sheet.spliceRows(1, 0, ["Title, not a supported header row"]);
    assert.equal(resolveWorksheetSchema(sheet).valid, false);
    const { config } = await project(context, workbook);
    const before = await hashFile(config.pgnSourceWorkbookPath);
    await assert.rejects(openExecutedPgnWorkbook(config.pgnSourceWorkbookPath, config.pgnExecutedWorkbookPath), /validation failed/i);
    await assert.rejects(runPgnWorkbook([], "full", config), /validation failed/i);
    await assert.rejects(access(config.pgnExecutedWorkbookPath));
    assert.equal(await hashFile(config.pgnSourceWorkbookPath), before);
  });
}

test("optional fields can be removed without positional writes", async (context) => {
  const workbook = fixtureWorkbook();
  for (const name of [KB_SHEET_NAME, NEGATIVE_SHEET_NAME]) {
    const sheet = workbook.getWorksheet(name)!;
    for (const field of ["notes", "responseTime", "testDate"] as const) sheet.spliceColumns(fieldColumn(sheet, field), 1);
  }
  const { config } = await project(context, workbook);
  const opened = await openExecutedPgnWorkbook(config.pgnSourceWorkbookPath, config.pgnExecutedWorkbookPath);
  assert(opened.parsed.issues.some((issue) => issue.severity === "WARNING" && issue.message.includes("Notes")));
  for (const scenario of opened.parsed.scenarios) {
    applyScenarioResults(opened.workbook, scenario, executions(scenario));
    assert.equal(optionalFieldCell(opened.workbook.getWorksheet(scenario.sheetName)!, scenario.sourceRowNumber, "notes"), undefined);
  }
});

for (const duplicate of [false, true]) {
  test(`${duplicate ? "ambiguous alias choice" : "unknown header override"} persists by header and follows movement after restart`, async (context) => {
    const workbook = fixtureWorkbook();
    const sheet = workbook.getWorksheet(KB_SHEET_NAME)!;
    sheet.getCell(1, duplicate ? 4 : 8).value = duplicate ? "User Question" : "Customer Prompt";
    const { config, root } = await project(context, workbook);
    const { ui, events } = scriptedUi(["change", "0:kb", "userInput", duplicate ? "4" : "8", "accept"]);
    const before = await hashFile(config.pgnSourceWorkbookPath);
    assert.equal(await reviewWorkbookMapping(ui, config), true);
    assert.equal(await hashFile(config.pgnSourceWorkbookPath), before);
    const store = await readWorkbookMappingStore(root);
    const key = workbookMappingKey(root, config.pgnSourceWorkbookPath);
    assert.deepEqual(store.workbooks[key].overrides, [{ sheet: KB_SHEET_NAME, field: "userInput", header: duplicate ? "User Question" : "Customer Prompt" }]);
    const loaded = await loadPgnWorkbook(config.pgnSourceWorkbookPath);
    const moved = loaded.workbook.getWorksheet(KB_SHEET_NAME)!;
    moved.spliceColumns(1, 0, ["Inserted reviewer"]);
    await loaded.workbook.xlsx.writeFile(config.pgnSourceWorkbookPath);
    const inspection = await inspectWorkbookMappings(config, true);
    assert.equal(inspection.ready, true);
    assert.match(formatWorkbookMappings(inspection), /userInput moved/);
    const restarted = await loadPgnWorkbook(config.pgnSourceWorkbookPath);
    assert.equal(fieldColumn(restarted.workbook.getWorksheet(KB_SHEET_NAME)!, "userInput"), duplicate ? 5 : 9);
    const saved = await readWorkbookMappingStore(root);
    assert.equal(saved.workbooks[key].sheets[KB_SHEET_NAME].fields.userInput.columnLetter, duplicate ? "E" : "I");
    assert.match(events.join("\n"), /Workbook-specific mapping saved/);
    const quiet = scriptedUi([]);
    assert.equal(await ensureReviewedWorkbookMapping(quiet.ui, config), true);
    assert.equal(quiet.events.length, 0);
  });
}

test("review cancellation never saves mappings or modifies workbooks", async (context) => {
  const { config, root } = await project(context);
  const before = await hashFile(config.pgnSourceWorkbookPath);
  assert.equal(await reviewWorkbookMapping(scriptedUi(["cancel"]).ui, config), false);
  await assert.rejects(access(path.join(root, WORKBOOK_MAPPING_FILE)));
  assert.equal(await hashFile(config.pgnSourceWorkbookPath), before);
});

for (const omitted of ["notes", "testDate"] as const) {
  test(`completed technical attempts remain recoverable without optional ${omitted}`, async (context) => {
    const workbook = fixtureWorkbook();
    const sheet = workbook.getWorksheet(KB_SHEET_NAME)!;
    sheet.spliceColumns(fieldColumn(sheet, omitted), 1);
    const { config, root } = await project(context, workbook);
    const opened = await openExecutedPgnWorkbook(config.pgnSourceWorkbookPath, config.pgnExecutedWorkbookPath);
    const { checkpoint, manifest } = await initializeRecoveryCheckpoint({
      projectRoot: root, runId: "OPTIONAL-RECOVERY", mode: "full",
      sourceWorkbookPath: config.pgnSourceWorkbookPath, executedWorkbookPath: config.pgnExecutedWorkbookPath,
      sourceWorkbookHash: await hashFile(config.pgnSourceWorkbookPath), scenarios: opened.parsed.scenarios,
    });
    const scenario = opened.parsed.scenarios[0];
    applyScenarioExecution(opened.workbook, "OPTIONAL-RECOVERY", scenario, [{
      ...executions(scenario)[0], technicalStatus: "TIMEOUT", combinedResponse: "", botMessages: [], error: "Synthetic timeout",
    }]);
    appendRecoveryTranscriptEvent(opened.workbook, { runId: "OPTIONAL-RECOVERY", scenario, event: "SCENARIO_ATTEMPT_FAILED", message: "Checkpointed technical attempt" });
    await saveExecutedPgnWorkbook(opened.workbook, config.pgnExecutedWorkbookPath);
    await checkpoint.update((state) => { state.status = "INTERRUPTED"; state.completedScenarioIds = [scenario.testCaseId]; });
    const reloaded = await loadPgnWorkbook(config.pgnExecutedWorkbookPath);
    const reconciled = reconcileRecoveryArtifacts(reloaded.workbook, reloaded.parsed.scenarios, checkpoint.snapshot(), manifest);
    assert.deepEqual(reconciled.mismatchedScenarioIds, []);
    assert.deepEqual(reconciled.safeCompletedIds, [scenario.testCaseId]);
  });
}

test("stale header overrides and colliding field assignments cannot fall back to letters", () => {
  const workbook = fixtureWorkbook();
  const sheet = workbook.getWorksheet(KB_SHEET_NAME)!;
  setWorkbookSchemaOverrides(workbook, [{ sheet: KB_SHEET_NAME, field: "userInput", header: "Removed custom header" }]);
  assert.equal(resolveWorksheetSchema(sheet).valid, false);
  assert.throws(() => fieldCell(sheet, 2, "userInput"), /Saved header override/);
  setWorkbookSchemaOverrides(workbook, [{ sheet: KB_SHEET_NAME, field: "botResponse", header: "User Input" }]);
  assert.equal(resolveWorksheetSchema(sheet).valid, false);
});

test("evidence is discovered outside its old location and appends beyond populated cells and tables", async (context) => {
  const workbook = fixtureWorkbook();
  const sheet = workbook.getWorksheet(KB_SHEET_NAME)!;
  sheet.getCell(1, 14).value = "Reviewer";
  sheet.getCell(2, 14).value = "Never overwrite";
  sheet.addTable({ name: "FixtureFarTable", ref: "Q1", columns: [{ name: "Far value" }], rows: [["Preserve far value"]] });
  sheet.getCell(2, 5).value = { formula: "1+1", result: 2 };
  sheet.getCell(2, 5).font = { bold: true, color: { argb: "FF123456" } };
  sheet.getColumn(5).width = 43;
  sheet.addConditionalFormatting({ ref: "E2:E4", rules: [{ type: "cellIs", operator: "greaterThan", formulae: ["0"], priority: 1, style: { font: { italic: true } } }] });
  const negative = workbook.getWorksheet(NEGATIVE_SHEET_NAME)!;
  negative.getCell(1, 19).value = "Evidence URL";
  negative.getCell(2, 19).value = { text: "Existing", hyperlink: "https://example.invalid/preserved-evidence" };
  const { config } = await project(context, workbook);
  const sourceZip = await JSZip.loadAsync(await readFile(config.pgnSourceWorkbookPath));
  const sourceTable = await sourceZip.file("xl/tables/table1.xml")!.async("string");
  const opened = await openExecutedPgnWorkbook(config.pgnSourceWorkbookPath, config.pgnExecutedWorkbookPath);
  const mapped = opened.workbook.getWorksheet(KB_SHEET_NAME)!;
  assert.equal(fieldColumn(mapped, "evidence"), 18);
  assert.equal(fieldColumn(opened.workbook.getWorksheet(NEGATIVE_SHEET_NAME)!, "evidence"), 19);
  const scenario = opened.parsed.scenarios[0];
  applyScenarioExecution(opened.workbook, "SCHEMA-RUN", scenario, executions(scenario));
  await saveExecutedPgnWorkbook(opened.workbook, config.pgnExecutedWorkbookPath);
  const reloaded = await loadPgnWorkbook(config.pgnExecutedWorkbookPath);
  const preserved = reloaded.workbook.getWorksheet(KB_SHEET_NAME)!;
  assert.equal(preserved.getCell(2, 14).text, "Never overwrite");
  assert.deepEqual(preserved.getCell(2, 5).value, { formula: "1+1", result: 2 });
  assert.equal(preserved.getCell(2, 5).font.color?.argb, "FF123456");
  assert.equal(preserved.getColumn(5).width, 43);
  assert.equal((preserved.model as unknown as { conditionalFormattings: unknown[] }).conditionalFormattings.length, 1);
  assert.equal(readEvidenceHyperlink(fieldCell(reloaded.workbook.getWorksheet(NEGATIVE_SHEET_NAME)!, 2, "evidence")), "https://example.invalid/preserved-evidence");
  const outputZip = await JSZip.loadAsync(await readFile(config.pgnExecutedWorkbookPath));
  assert.equal(await outputZip.file("xl/tables/table1.xml")!.async("string"), sourceTable);
});

test("reordered controlled transcript/history/metadata fields are read and written semantically", async (context) => {
  const { config } = await project(context);
  const opened = await openExecutedPgnWorkbook(config.pgnSourceWorkbookPath, config.pgnExecutedWorkbookPath);
  const scenario = opened.parsed.scenarios[0];
  applyScenarioExecution(opened.workbook, "RUN-A", scenario, executions(scenario));
  const transcript = opened.workbook.getWorksheet("Execution Transcript")!;
  reorder(transcript, Array.from({ length: 15 }, (_, index) => 15 - index));
  snapshotRetestHistory(opened.workbook, "RETEST-A", scenario, new Date());
  const history = opened.workbook.getWorksheet("Retest History")!;
  reorder(history, Array.from({ length: 18 }, (_, index) => 18 - index));
  updateRetestHistory(opened.workbook, "RETEST-A", scenario, executions(scenario));
  assert.equal(fieldCell(history, 2, "runId").text, "RUN-A");
  assert.equal(fieldCell(history, 2, "newBotResponse").text, "Synthetic response 1");
  upsertRetestRunMetadata(opened.workbook, { runId: "RETEST-A", startedAt: new Date(), updatedAt: new Date(), state: "IN_PROGRESS", selectedIds: [scenario.testCaseId], finishedIds: [] });
  const retestMetadata = opened.workbook.getWorksheet("Retest Metadata")!;
  reorder(retestMetadata, [8, 7, 6, 5, 4, 3, 2, 1]);
  assert.deepEqual(getRetestRunMetadata(opened.workbook, "RETEST-A")?.selectedIds, [scenario.testCaseId]);
  const metadata = opened.workbook.getWorksheet("Execution Metadata")!;
  reorder(metadata, [6, 5, 4, 3, 2, 1, 7, 8, 16, 15, 14, 13, 12, 11, 10, 9]);
  upsertEvidenceRunMetadata(opened.workbook, { runId: "RUN-A", folderId: "folder-fixture", folderUrl: "", migrationVersion: "1", timestamp: new Date(), mode: "FUTURE" });
  upsertEvidenceFileMetadata(opened.workbook, { evidenceKey: "key-fixture", runId: "RUN-A", testCaseId: scenario.testCaseId, turnNumber: 1, driveFileName: "fixture.png", status: "EVIDENCE_MISSING" });
  assert.equal(getEvidenceFileMetadata(opened.workbook, "key-fixture")?.runId, "RUN-A");
  assert.equal(getEvidenceRunMetadata(opened.workbook, "RUN-A")?.folderId, "folder-fixture");
  assert.notEqual(fieldColumn(metadata, "runId", EVIDENCE_RUN_SCHEMA), fieldColumn(metadata, "runId", EVIDENCE_FILE_SCHEMA));
  removeEvidenceFileMetadata(opened.workbook, "key-fixture");
  assert.equal(getEvidenceRunMetadata(opened.workbook, "RUN-A")?.folderId, "folder-fixture");
  appendRecoveryTranscriptEvent(opened.workbook, { runId: "RUN-A", scenario, event: "RUN_INTERRUPTED", message: "Synthetic audit" });
  assert.equal(fieldCell(transcript, transcript.rowCount, "status").text, "RUN_INTERRUPTED");
  assert.equal(discoverEvidenceInventory(opened.workbook).records[0]?.testCaseId, scenario.testCaseId);
});

for (const change of ["formatting", "columns", "alias"] as const) {
  test(`recovery preserves content hashing and rejects structural ${change} drift safely`, async (context) => {
    const { config, root } = await project(context);
    const opened = await openExecutedPgnWorkbook(config.pgnSourceWorkbookPath, config.pgnExecutedWorkbookPath);
    const original = await loadPgnWorkbook(config.pgnSourceWorkbookPath);
    await initializeRecoveryCheckpoint({ projectRoot: root, runId: "SCHEMA-RECOVERY", mode: "full", sourceWorkbookPath: config.pgnSourceWorkbookPath, executedWorkbookPath: config.pgnExecutedWorkbookPath, sourceWorkbookHash: await hashFile(config.pgnSourceWorkbookPath), scenarios: original.parsed.scenarios });
    const sheet = original.workbook.getWorksheet(KB_SHEET_NAME)!;
    if (change === "formatting") sheet.getCell(1, 1).font = { italic: true };
    if (change === "columns") sheet.spliceColumns(8, 0, ["Reviewer"]);
    if (change === "alias") sheet.getCell(1, 8).value = "User Question";
    await original.workbook.xlsx.writeFile(config.pgnSourceWorkbookPath);
    const validation = await validateRecoveryRun(config, "SCHEMA-RECOVERY", { checkDriveAccess: false, profileEntries: async () => ["fixture"] });
    assert.equal(validation.sourceDrift, change === "formatting" ? "formatting-only" : "structural");
    if (change !== "formatting") assert.equal(validation.ready, false);
    assert.equal(opened.parsed.scenarios[0].schemaFingerprint, mainSchemaFingerprint(opened.workbook.getWorksheet(KB_SHEET_NAME)!));
  });
}

test("validation and doctor report schema without workbook or mapping writes", async (context) => {
  const { config, root } = await project(context);
  const before = await hashFile(config.pgnSourceWorkbookPath);
  context.mock.method(console, "log", () => undefined);
  assert.equal(await validatePgnWorkbook(config), true);
  const report = await collectDiagnostics({ projectRoot: root, environment: { GOOGLE_DRIVE_EVIDENCE_ENABLED: "false", DISCORD_NOTIFICATIONS_ENABLED: "false" }, checkDriveAccess: false, npmVersion: async () => "11", packageVersion: async () => "1", chromiumExecutablePath: async () => undefined, hasCommand: async () => false });
  assert.equal(report.checks.find((check) => check.id === "workbook-schema")?.status, "ok");
  assert.equal(await hashFile(config.pgnSourceWorkbookPath), before);
  await assert.rejects(access(path.join(root, WORKBOOK_MAPPING_FILE)));
  await assert.rejects(access(config.pgnExecutedWorkbookPath));
});

test("recovery rejects changed header overrides even when source file bytes are unchanged", async (context) => {
  const workbook = fixtureWorkbook();
  const sheet = workbook.getWorksheet(KB_SHEET_NAME)!;
  sheet.getCell(1, 15).value = "Customer Prompt";
  sheet.getCell(2, 15).value = "Different synthetic input";
  const { config, root } = await project(context, workbook);
  const opened = await openExecutedPgnWorkbook(config.pgnSourceWorkbookPath, config.pgnExecutedWorkbookPath);
  const originalHash = await hashFile(config.pgnSourceWorkbookPath);
  await initializeRecoveryCheckpoint({
    projectRoot: root, runId: "OVERRIDE-RECOVERY", mode: "full",
    sourceWorkbookPath: config.pgnSourceWorkbookPath, executedWorkbookPath: config.pgnExecutedWorkbookPath,
    sourceWorkbookHash: originalHash, scenarios: opened.parsed.scenarios,
  });
  await saveWorkbookMappingStore(root, { version: 1, workbooks: {
    [workbookMappingKey(root, config.pgnSourceWorkbookPath)]: {
      overrides: [{ sheet: KB_SHEET_NAME, field: "userInput", header: "Customer Prompt" }], sheets: {},
    },
  } });
  assert.equal(await hashFile(config.pgnSourceWorkbookPath), originalHash);
  const validation = await validateRecoveryRun(config, "OVERRIDE-RECOVERY", { checkDriveAccess: false });
  assert.equal(validation.sourceDrift, "structural");
  assert.equal(validation.ready, false);
  await assert.rejects(reviewWorkbookMapping(scriptedUi([]).ui, config), /active or recoverable run exists/);
});

for (const change of ["formatting", "columns"] as const) {
  test(`legacy checkpoints retain safe ${change} drift behavior without inventing a schema snapshot`, async (context) => {
    const { config, root } = await project(context);
    await openExecutedPgnWorkbook(config.pgnSourceWorkbookPath, config.pgnExecutedWorkbookPath);
    const source = await loadPgnWorkbook(config.pgnSourceWorkbookPath);
    await initializeRecoveryCheckpoint({
      projectRoot: root, runId: "LEGACY-SCHEMA", mode: "full",
      sourceWorkbookPath: config.pgnSourceWorkbookPath, executedWorkbookPath: config.pgnExecutedWorkbookPath,
      sourceWorkbookHash: await hashFile(config.pgnSourceWorkbookPath),
      scenarios: source.parsed.scenarios.map(({ schemaFingerprint: _, ...scenario }) => scenario),
    });
    const sheet = source.workbook.getWorksheet(KB_SHEET_NAME)!;
    if (change === "formatting") sheet.getCell(1, 1).font = { italic: true };
    else sheet.spliceColumns(8, 0, ["Reviewer"]);
    await source.workbook.xlsx.writeFile(config.pgnSourceWorkbookPath);
    const validation = await validateRecoveryRun({ ...config, target: { kind: "chat", value: "Synthetic fixture" } }, "LEGACY-SCHEMA", { checkDriveAccess: false, profileEntries: async () => ["fixture"] });
    assert.equal(validation.sourceDrift, change === "formatting" ? "formatting-only" : "structural");
    assert.equal(validation.ready, change === "formatting");
  });
}

test("Workbook menu exposes status, review, and redetection without dispatching live actions", async () => {
  const calls: string[] = [];
  const { ui } = scriptedUi(["workbook", "status", "review", "redetect", "back", "exit"]);
  const actions = new Proxy({ workbookSchema: async (review: boolean, redetect?: boolean) => { calls.push(`${review}:${redetect}`); } }, { get(target, key) { return key in target ? target[key as keyof typeof target] : undefined; } }) as unknown as OperatorActions;
  await runControlPanel(ui, actions);
  assert.deepEqual(calls, ["false:false", "true:false", "false:true"]);
});

test("setup offers workbook review only after schema diagnostics", async (context) => {
  const { root } = await project(context);
  let reviews = 0;
  const { ui } = scriptedUi([false, true, "exit"]);
  const result = await runSetupWizard(ui, {
    projectRoot: root, environment: {}, platform: "win32",
    diagnose: async () => ({ checks: [{ id: "workbook-schema", label: "Workbook schema", status: "ok", detail: "valid" }], browserRuntime: { mode: "direct", reason: "Fixture" }, chromiumInstalled: true, profilePresent: true, environmentFilePresent: true, ready: true }),
    reviewWorkbookMapping: async () => { reviews += 1; return true; },
  });
  assert.equal(reviews, 1);
  assert.equal(result.nextAction, "exit");
});
