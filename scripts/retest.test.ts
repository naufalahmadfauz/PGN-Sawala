import assert from "node:assert/strict";
import { execFile as execFileCallback } from "node:child_process";
import { mkdir, mkdtemp, readFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import test from "node:test";
import { promisify } from "node:util";
import ExcelJS from "exceljs";
import { loadConfig, type AppConfig } from "../src/config";
import {
  MAIN_EVIDENCE_COLUMN,
  readEvidenceHyperlink,
} from "../src/excel/evidence-workbook";
import { createFreshPgnWorkbook } from "../src/excel/fresh-workbook";
import {
  isScenarioComplete,
  parsePgnWorkbook,
} from "../src/excel/pgn-workbook-loader";
import {
  PGN_TEST_STATUSES,
  normalizeTestStatus,
} from "../src/excel/pgn-test-status";
import {
  RETEST_HISTORY_SHEET_NAME,
  applyRetestStatusTransition,
  ensureRetestWorkbookSchema,
  getRetestRunMetadata,
  setScenarioStatus,
  snapshotRetestHistory,
  updateRetestHistory,
  upsertRetestRunMetadata,
} from "../src/excel/retest-workbook";
import type {
  ExecutedTurn,
  PgnTestScenario,
} from "../src/excel/pgn-types";
import {
  applyScenarioExecution,
  applyScenarioResults,
  openExecutedPgnWorkbook,
  saveExecutedPgnWorkbook,
} from "../src/excel/pgn-workbook-writer";
import { parseCliOptions } from "../src/pgn-cli";
import { selectScenarios } from "../src/pgn-selection";
import {
  createRetestRunId,
  needsFinalRetestCleanup,
  retestDriveFolderName,
} from "../src/retest/retest-run";
import { selectRetestScenarios } from "../src/retest/retest-selection";

const sourcePath = path.resolve(
  "data/PGN AI Assistant - Knowledge Base Testing Report - User Inputs.xlsx",
);
const execFile = promisify(execFileCallback);
const expectedRetestIds = [
  "PGN-KB-017",
  "PGN-KB-023",
  "PGN-KB-026",
  "PGN-KB-040",
  "PGN-KB-059",
  "PGN-KB-060",
  "PGN-KB-066",
  "PGN-KB-069",
  "PGN-KB-070",
  "PGN-KB-073",
  "PGN-KB-075",
];

function scenario(
  testCaseId: string,
  rawStatus: string,
): PgnTestScenario {
  return {
    testCaseId,
    sheetKind: "kb",
    sheetName: "Test Case Knowledge Base",
    sourceRowNumber: 2,
    category: "Test",
    rawStatus,
    status: normalizeTestStatus(rawStatus),
    turns: [
      {
        sheetName: "Test Case Knowledge Base",
        rowNumber: 2,
        turnNumber: 1,
        userInput: "Test input",
      },
    ],
  };
}

function execution(
  selectedScenario: PgnTestScenario,
  turnIndex: number,
  label: string,
  evidenceUrl?: string,
): ExecutedTurn {
  const turn = selectedScenario.turns[turnIndex];
  return {
    turn,
    technicalStatus: "CAPTURED",
    sentAt: new Date("2026-09-02T05:30:00Z"),
    completedAt: new Date("2026-09-02T05:30:01Z"),
    botMessages: [
      {
        sequence: 1,
        message: `${label} response ${turn.turnNumber}`,
        timestamp: new Date("2026-09-02T05:30:01Z"),
      },
    ],
    combinedResponse: `${label} response ${turn.turnNumber}`,
    firstResponseMs: 500,
    totalResponseMs: 1_000,
    evidenceUrl,
    evidenceStatus: evidenceUrl
      ? "EVIDENCE_SYNCED"
      : "EVIDENCE_UPLOAD_ERROR",
  };
}

test("status normalization selects only explicit Ready for Re-test values", () => {
  const scenarios = [
    scenario("PASSED", "Passed"),
    scenario("FAILED", "Failed"),
    scenario("BLOCKED", "Blocked"),
    scenario("REVIEW", "Review"),
    scenario("READY", " Ready for Re-test "),
    scenario("READY-CASE", "READY FOR RE-TEST"),
    scenario("READY-ALIAS", "Ready for Retest"),
    scenario("PENDING", "Pending Evaluation"),
    scenario("BROAD-RETEST", "retest"),
    scenario("BROAD-READY", "ready"),
  ];
  const selection = selectRetestScenarios(scenarios);
  assert.deepEqual(
    selection.selected.map((item) => item.testCaseId),
    ["READY", "READY-CASE", "READY-ALIAS"],
  );
  const explicit = selectRetestScenarios(scenarios, {
    testIds: new Set(["FAILED"]),
  });
  assert.deepEqual(
    explicit.selected.map((item) => item.testCaseId),
    ["FAILED"],
  );
  assert.equal(explicit.warnings.length, 1);
  const runId = createRetestRunId(new Date("2026-09-02T05:30:00Z"));
  assert.equal(runId, "RETEST-20260902T053000Z");
  assert.equal(
    retestDriveFolderName("PGN-WhatsApp-Retest", runId),
    "PGN-WhatsApp-Retest-20260902T053000Z",
  );
  assert.equal(
    selectRetestScenarios(scenarios, { limit: 2 }).selected.length,
    2,
  );
  assert.equal(
    needsFinalRetestCleanup(
      {
        state: "IN_PROGRESS",
        selectedIds: ["READY"],
        finishedIds: ["READY"],
      },
      0,
    ),
    true,
  );
  assert.equal(
    needsFinalRetestCleanup(
      {
        state: "IN_PROGRESS",
        selectedIds: ["READY"],
        finishedIds: [],
      },
      1,
    ),
    false,
  );
});

test("current workbook fixture selects the expected eleven IDs and whole multi-turn scenarios", async () => {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(sourcePath);
  const worksheet = workbook.getWorksheet("Test Case Knowledge Base")!;
  const variants = [
    "Ready for Re-test",
    " ready for re-test ",
    "READY FOR RE-TEST",
    "Ready for Retest",
  ];
  for (const [index, testCaseId] of expectedRetestIds.entries()) {
    for (let rowNumber = 2; rowNumber <= worksheet.rowCount; rowNumber += 1) {
      if (worksheet.getCell(rowNumber, 3).text === testCaseId) {
        worksheet.getCell(rowNumber, 12).value = variants[index % variants.length];
      }
    }
  }
  let parsed = parsePgnWorkbook(workbook);
  let selection = selectRetestScenarios(parsed.scenarios);
  assert.deepEqual(
    selection.selected.map((item) => item.testCaseId),
    expectedRetestIds,
  );
  assert.equal(
    selection.selected.filter((item) => item.turns.length > 1).length,
    0,
  );

  const multiTurnRow = parsed.scenarios.find(
    (item) => item.testCaseId === "PGN-KB-031",
  )!.sourceRowNumber;
  worksheet.getCell(multiTurnRow, 12).value = "Ready for Re-test";
  parsed = parsePgnWorkbook(workbook);
  selection = selectRetestScenarios(parsed.scenarios, {
    testIds: new Set(["PGN-KB-031"]),
  });
  assert.equal(selection.selected.length, 1);
  assert.equal(selection.selected[0].turns.length, 2);

  const negative = workbook.getWorksheet("Negative Case")!;
  const negativeScenario = parsed.scenarios.find(
    (item) => item.testCaseId === "PGN-NEG-016",
  )!;
  negative.getCell(negativeScenario.sourceRowNumber, 11).value =
    "Ready for Re-test";
  parsed = parsePgnWorkbook(workbook);
  selection = selectRetestScenarios(parsed.scenarios, {
    testIds: new Set(["PGN-NEG-016"]),
  });
  assert.equal(selection.selected.length, 1);
  assert.equal(selection.selected[0].turns.length, 2);

  worksheet.getCell(51, 12).value = "ready retest now";
  parsed = parsePgnWorkbook(workbook);
  assert(
    parsed.issues.some(
      (issue) => issue.code === "UNKNOWN_STATUS" && issue.rowNumber === 51,
    ),
  );
});

test("zero automatic candidates produce an empty selection", () => {
  const nonReadyScenarios = [
    scenario("PASSED", "Passed"),
    scenario("FAILED", "Failed"),
    scenario("BLOCKED", "Blocked"),
    scenario("REVIEW", "Review"),
    scenario("PENDING", "Pending Evaluation"),
    scenario("BLANK", ""),
  ];
  assert.equal(selectRetestScenarios(nonReadyScenarios).selected.length, 0);
});

test("zero-candidate retest command exits before WhatsApp startup", async (context) => {
  await mkdir(path.resolve("reports"), { recursive: true });
  const sourceDirectory = await mkdtemp(
    path.join(path.dirname(sourcePath), ".retest-zero-"),
  );
  const reportDirectory = await mkdtemp(
    path.resolve("reports", ".retest-zero-"),
  );
  context.after(async () => {
    await Promise.all([
      rm(sourceDirectory, { recursive: true, force: true }),
      rm(reportDirectory, { recursive: true, force: true }),
    ]);
  });
  const fixtureSourcePath = path.join(sourceDirectory, "source.xlsx");
  const fixtureOutputPath = path.join(reportDirectory, "executed.xlsx");
  const fixture = new ExcelJS.Workbook();
  await fixture.xlsx.readFile(sourcePath);
  for (const item of parsePgnWorkbook(fixture).scenarios) {
    fixture
      .getWorksheet(item.sheetName)!
      .getCell(item.sourceRowNumber, item.sheetKind === "kb" ? 12 : 11).value =
      PGN_TEST_STATUSES.Passed;
  }
  await fixture.xlsx.writeFile(fixtureSourcePath);

  const { stdout, stderr } = await execFile(
    path.resolve("node_modules/.bin/tsx"),
    [path.resolve("scripts/retest-pgn.ts")],
    {
      cwd: path.resolve("."),
      env: {
        ...process.env,
        GOOGLE_DRIVE_EVIDENCE_ENABLED: "false",
        GOOGLE_DRIVE_EVIDENCE_PARENT_FOLDER: "",
        GOOGLE_SERVICE_ACCOUNT_JSON: "",
        PGN_EXECUTED_WORKBOOK: path.relative(
          path.resolve("."),
          fixtureOutputPath,
        ),
        PGN_SOURCE_WORKBOOK: path.relative(
          path.resolve("."),
          fixtureSourcePath,
        ),
        PGN_WHATSAPP_CHAT: "",
        PGN_WHATSAPP_PHONE: "",
        WHATSAPP_BROWSER_CHANNEL: "invalid-if-opened",
      },
      timeout: 30_000,
    },
  );
  assert.match(stdout, /Ready for Re-test: 0/);
  assert.match(stdout, /Nothing to execute\./);
  assert.equal(stderr, "");
  const output = new ExcelJS.Workbook();
  await output.xlsx.readFile(fixtureOutputPath);
  assert.equal(output.getWorksheet("Retest Metadata"), undefined);
});

test("Retest History preserves old values, deduplicates snapshots, and supports resume", async (context) => {
  const temporaryDirectory = await mkdtemp(path.join(tmpdir(), "pgn-retest-history-"));
  context.after(() => rm(temporaryDirectory, { recursive: true, force: true }));
  const outputPath = path.join(temporaryDirectory, "executed.xlsx");
  const opened = await openExecutedPgnWorkbook(sourcePath, outputPath);
  const selectedScenario = opened.parsed.scenarios.find(
    (item) => item.testCaseId === "PGN-KB-031",
  )!;
  const untouchedScenario = opened.parsed.scenarios.find(
    (item) => item.testCaseId === "PGN-KB-032",
  )!;
  const untouchedWorksheet = opened.workbook.getWorksheet(
    untouchedScenario.sheetName,
  )!;
  const untouchedBefore = [9, 10, 11, 12, 13, 14].map(
    (column) =>
      untouchedWorksheet.getCell(untouchedScenario.sourceRowNumber, column).text,
  );
  const oldExecutions = selectedScenario.turns.map((_, index) =>
    execution(
      selectedScenario,
      index,
      "Old",
      `https://drive.google.com/file/d/old-${index + 1}/view`,
    ),
  );
  applyScenarioExecution(
    opened.workbook,
    "20260901T010000Z",
    selectedScenario,
    oldExecutions,
  );
  setScenarioStatus(
    opened.workbook,
    selectedScenario,
    PGN_TEST_STATUSES.ReadyForRetest,
  );
  ensureRetestWorkbookSchema(opened.workbook);
  const retestRunId = "RETEST-20260902T053000Z";
  snapshotRetestHistory(
    opened.workbook,
    retestRunId,
    selectedScenario,
    new Date("2026-09-02T05:30:00Z"),
  );
  snapshotRetestHistory(
    opened.workbook,
    retestRunId,
    selectedScenario,
    new Date("2026-09-02T05:31:00Z"),
  );
  const history = opened.workbook.getWorksheet(RETEST_HISTORY_SHEET_NAME)!;
  assert.equal(history.rowCount, 3);
  assert.equal(history.getCell(2, 7).text, "Ready for Re-test");
  assert.equal(history.getCell(2, 8).text, "Old response 1");
  assert.equal(
    readEvidenceHyperlink(history.getCell(2, 11)),
    "https://drive.google.com/file/d/old-1/view",
  );

  const newExecutions = selectedScenario.turns.map((_, index) =>
    execution(
      selectedScenario,
      index,
      "New",
      `https://drive.google.com/file/d/new-${index + 1}/view`,
    ),
  );
  applyScenarioResults(opened.workbook, selectedScenario, newExecutions);
  updateRetestHistory(
    opened.workbook,
    retestRunId,
    selectedScenario,
    newExecutions,
  );
  assert.equal(
    applyRetestStatusTransition(
      opened.workbook,
      selectedScenario,
      newExecutions,
    ),
    true,
  );
  upsertRetestRunMetadata(opened.workbook, {
    runId: retestRunId,
    startedAt: new Date("2026-09-02T05:30:00Z"),
    state: "IN_PROGRESS",
    selectedIds: ["PGN-KB-031", "PGN-KB-075"],
    finishedIds: ["PGN-KB-031"],
    updatedAt: new Date("2026-09-02T05:32:00Z"),
  });
  await saveExecutedPgnWorkbook(opened.workbook, outputPath);

  const reloaded = await openExecutedPgnWorkbook(sourcePath, outputPath);
  const metadata = getRetestRunMetadata(reloaded.workbook, retestRunId)!;
  assert.deepEqual(metadata.finishedIds, ["PGN-KB-031"]);
  const resumed = selectRetestScenarios(reloaded.parsed.scenarios, {
    resumeSelectedIds: metadata.selectedIds,
    completedIds: new Set(metadata.finishedIds),
  });
  assert.deepEqual(
    resumed.selected.map((item) => item.testCaseId),
    ["PGN-KB-075"],
  );
  const reloadedHistory = reloaded.workbook.getWorksheet(
    RETEST_HISTORY_SHEET_NAME,
  )!;
  assert.equal(reloadedHistory.getCell(2, 8).text, "Old response 1");
  assert.equal(reloadedHistory.getCell(2, 14).text, "New response 1");
  assert.equal(
    readEvidenceHyperlink(reloadedHistory.getCell(3, 17)),
    "https://drive.google.com/file/d/new-2/view",
  );
  assert.equal(
    reloaded.workbook
      .getWorksheet(selectedScenario.sheetName)!
      .getCell(selectedScenario.sourceRowNumber, 12).text,
    "Pending Evaluation",
  );
  assert.deepEqual(
    [9, 10, 11, 12, 13, 14].map(
      (column) =>
        reloaded.workbook
          .getWorksheet(untouchedScenario.sheetName)!
          .getCell(untouchedScenario.sourceRowNumber, column).text,
    ),
    untouchedBefore,
  );
});

test("a technical retest failure preserves the semantic status", async (context) => {
  const temporaryDirectory = await mkdtemp(
    path.join(tmpdir(), "pgn-retest-failure-"),
  );
  context.after(() => rm(temporaryDirectory, { recursive: true, force: true }));
  const outputPath = path.join(temporaryDirectory, "executed.xlsx");
  const opened = await openExecutedPgnWorkbook(sourcePath, outputPath);
  const selectedScenario = opened.parsed.scenarios.find(
    (item) => item.testCaseId === "PGN-KB-075",
  )!;
  setScenarioStatus(
    opened.workbook,
    selectedScenario,
    PGN_TEST_STATUSES.ReadyForRetest,
  );
  const failedExecution: ExecutedTurn = {
    ...execution(selectedScenario, 0, "Failed"),
    technicalStatus: "TIMEOUT",
    botMessages: [],
    combinedResponse: "",
    error: "Timed out waiting for a bot response",
  };
  applyScenarioResults(opened.workbook, selectedScenario, [failedExecution]);

  assert.equal(
    applyRetestStatusTransition(
      opened.workbook,
      selectedScenario,
      [failedExecution],
    ),
    false,
  );
  assert.equal(
    opened.workbook
      .getWorksheet(selectedScenario.sheetName)!
      .getCell(selectedScenario.sourceRowNumber, 12).text,
    PGN_TEST_STATUSES.ReadyForRetest,
  );
  assert.deepEqual(
    selectRetestScenarios(opened.parsed.scenarios, {
      resumeSelectedIds: [selectedScenario.testCaseId],
      completedIds: new Set(),
    }).selected.map((item) => item.testCaseId),
    [selectedScenario.testCaseId],
  );
});

test("fresh preparation archives the old report and leaves the full suite runnable", async (context) => {
  const temporaryDirectory = await mkdtemp(path.join(tmpdir(), "pgn-fresh-run-"));
  context.after(() => rm(temporaryDirectory, { recursive: true, force: true }));
  const base = loadConfig();
  const config: AppConfig = {
    ...base,
    pgnSourceWorkbookPath: sourcePath,
    pgnExecutedWorkbookPath: path.join(temporaryDirectory, "reports", "executed.xlsx"),
    reportArchiveDir: path.join(temporaryDirectory, "reports", "archive"),
  };
  const previousReport = await openExecutedPgnWorkbook(
    sourcePath,
    config.pgnExecutedWorkbookPath,
  );
  await saveExecutedPgnWorkbook(
    previousReport.workbook,
    config.pgnExecutedWorkbookPath,
  );
  const previous = await readFile(config.pgnExecutedWorkbookPath);
  const sourceBefore = await readFile(sourcePath);
  const result = await createFreshPgnWorkbook(
    config,
    new Date("2026-09-02T05:30:00Z"),
  );
  assert(result.archivePath);
  assert((await readFile(result.archivePath)).equals(previous));
  assert((await readFile(sourcePath)).equals(sourceBefore));
  const fresh = new ExcelJS.Workbook();
  await fresh.xlsx.readFile(config.pgnExecutedWorkbookPath);
  const parsed = parsePgnWorkbook(fresh);
  assert(parsed.scenarios.every((item) => !isScenarioComplete(fresh, item)));
  const normalSelection = selectScenarios(
    parsed.scenarios,
    parseCliOptions([]),
    fresh,
  );
  assert.equal(normalSelection.runnable.length, parsed.scenarios.length);
  assert.equal(fresh.getWorksheet("Execution Transcript"), undefined);
  assert.equal(fresh.getWorksheet("Retest History"), undefined);
});
