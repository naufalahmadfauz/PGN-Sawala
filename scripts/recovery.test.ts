import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { once } from "node:events";
import {
  access,
  copyFile,
  mkdir,
  mkdtemp,
  readFile,
  readdir,
  rm,
  writeFile,
} from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { pathToFileURL } from "node:url";
import test from "node:test";
import ExcelJS from "exceljs";
import { loadConfig, type AppConfig } from "../src/config";
import {
  isScenarioComplete,
  isScenarioPartiallyComplete,
} from "../src/excel/pgn-workbook-loader";
import {
  appendLatestTurnExecution,
  appendRecoveryTranscriptEvent,
  applyScenarioExecution,
  applyScenarioResults,
  openExecutedPgnWorkbook,
  saveExecutedPgnWorkbook,
} from "../src/excel/pgn-workbook-writer";
import { PGN_TEST_STATUSES } from "../src/excel/pgn-test-status";
import type {
  ExecutedTurn,
  PgnTestScenario,
} from "../src/excel/pgn-types";
import {
  applyRetestStatusTransition,
  getRetestRunMetadata,
  setScenarioStatus,
  upsertRetestRunMetadata,
} from "../src/excel/retest-workbook";
import type {
  DriveEvidenceItem,
  EvidenceDrivePublisher,
} from "../src/evidence/google-drive";
import { isProcessAlive } from "../src/process-liveness";
import {
  assertResumeOptionsCompatible,
  parseCliOptions,
} from "../src/pgn-cli";
import {
  abandonRecoveryRun,
  reconcileRecoveryArtifacts,
  repairRecoveryProgress,
  selectRecoveryScenarios,
  skipRecoveryScenario,
  validateRecoveryRun,
} from "../src/recovery/recovery-service";
import {
  acquireRunProcessLock,
  atomicWriteJson,
  discoverRecoveryRun,
  hashFile,
  initializeRecoveryCheckpoint,
  inspectRunProcessLock,
  readRecoveryRun,
  recoveryPaths,
  type RecoveryCheckpoint,
} from "../src/recovery/run-state";
import { prepareFreshPgnWorkbook } from "./fresh-pgn";
import { parseResumeValidationArgs } from "./validate-resume";

const repositoryRoot = path.resolve(".");
const repositorySourcePath = path.resolve(
  "data/PGN AI Assistant - Knowledge Base Testing Report - User Inputs.xlsx",
);

interface RecoveryFixture {
  root: string;
  config: AppConfig;
  checkpoint: RecoveryCheckpoint;
  runId: string;
  scenarios: PgnTestScenario[];
}

async function pathExists(filePath: string): Promise<boolean> {
  return access(filePath)
    .then(() => true)
    .catch(() => false);
}

async function createFixture(
  mode: "full" | "retest" = "full",
  scenarioSelector?: (
    scenarios: readonly PgnTestScenario[],
    allScenarios: readonly PgnTestScenario[],
  ) => PgnTestScenario[],
): Promise<RecoveryFixture> {
  const root = await mkdtemp(path.join(tmpdir(), "pgn-recovery-"));
  const environment: NodeJS.ProcessEnv = {
    PGN_WHATSAPP_CHAT: "Recovery fixture chat",
    WHATSAPP_HEADLESS: "true",
    GOOGLE_DRIVE_EVIDENCE_ENABLED: mode === "retest" ? "true" : "false",
    DISCORD_NOTIFICATIONS_ENABLED: "false",
  };
  const config = loadConfig({ repositoryRoot: root, environment });
  await Promise.all([
    mkdir(path.dirname(config.pgnSourceWorkbookPath), { recursive: true }),
    mkdir(path.dirname(config.pgnExecutedWorkbookPath), { recursive: true }),
    mkdir(config.profileDir, { recursive: true }),
  ]);
  await copyFile(repositorySourcePath, config.pgnSourceWorkbookPath);
  await writeFile(path.join(config.profileDir, "fixture-session"), "fixture");
  const opened = await openExecutedPgnWorkbook(
    config.pgnSourceWorkbookPath,
    config.pgnExecutedWorkbookPath,
  );
  const runnableScenarios = opened.parsed.scenarios.filter(
    (scenario) =>
      !isScenarioComplete(opened.workbook, scenario) &&
      !isScenarioPartiallyComplete(opened.workbook, scenario),
  );
  const scenarios = scenarioSelector
    ? scenarioSelector(runnableScenarios, opened.parsed.scenarios)
    : [
        runnableScenarios.find(
          (scenario) => scenario.sheetKind === "kb" && scenario.turns.length === 1,
        )!,
        runnableScenarios.find(
          (scenario) => scenario.sheetKind === "kb" && scenario.turns.length > 1,
        )!,
      ];
  assert(scenarios.length > 0 && scenarios.every(Boolean));
  const runId = mode === "retest" ? "RETEST-RECOVERY-FIXTURE" : "RECOVERY-FIXTURE";
  const startedAt = new Date("2026-09-05T10:00:00.000Z");
  if (mode === "retest") {
    for (const scenario of scenarios) {
      setScenarioStatus(opened.workbook, scenario, PGN_TEST_STATUSES.ReadyForRetest);
    }
    upsertRetestRunMetadata(opened.workbook, {
      runId,
      startedAt,
      state: "IN_PROGRESS",
      selectedIds: scenarios.map((scenario) => scenario.testCaseId),
      finishedIds: [],
      updatedAt: startedAt,
    });
    await saveExecutedPgnWorkbook(opened.workbook, config.pgnExecutedWorkbookPath);
  }
  const initialized = await initializeRecoveryCheckpoint({
    projectRoot: root,
    runId,
    mode,
    sourceWorkbookPath: config.pgnSourceWorkbookPath,
    executedWorkbookPath: config.pgnExecutedWorkbookPath,
    sourceWorkbookHash: await hashFile(config.pgnSourceWorkbookPath),
    scenarios,
    startedAt,
  });
  return { root, config, checkpoint: initialized.checkpoint, runId, scenarios };
}

function capturedExecution(
  scenario: PgnTestScenario,
  turnIndex: number,
  completedAt = new Date("2026-09-05T10:01:01.000Z"),
): ExecutedTurn {
  const turn = scenario.turns[turnIndex];
  return {
    turn,
    technicalStatus: "CAPTURED",
    sentAt: new Date(completedAt.getTime() - 1_000),
    completedAt,
    botMessages: [
      {
        sequence: 1,
        message: `Fixture response ${turn.turnNumber}`,
        timestamp: completedAt,
      },
    ],
    combinedResponse: `Fixture response ${turn.turnNumber}`,
    firstResponseMs: 500,
    totalResponseMs: 1_000,
    evidencePath: `artifacts/evidence/${scenario.testCaseId}-${turn.turnNumber}.png`,
    evidenceStatus: "EVIDENCE_LOCAL_ONLY",
  };
}

function representativeScenarios(
  scenarios: readonly PgnTestScenario[],
): PgnTestScenario[] {
  return (["kb", "negative"] as const).flatMap((sheetKind) =>
    [false, true].map((multiTurn) => {
      const scenario = scenarios.find(
        (scenario) =>
          scenario.sheetKind === sheetKind &&
          (multiTurn ? scenario.turns.length > 1 : scenario.turns.length === 1),
      );
      assert(
        scenario,
        `Fixture needs a ${sheetKind} ${multiTurn ? "multi" : "single"}-turn scenario`,
      );
      return scenario;
    }),
  );
}

async function recordScenarioAttempt(
  fixture: RecoveryFixture,
  scenario: PgnTestScenario,
  updateCheckpoint: boolean,
  technicalStatus: ExecutedTurn["technicalStatus"] = "CAPTURED",
): Promise<void> {
  const opened = await openExecutedPgnWorkbook(
    fixture.config.pgnSourceWorkbookPath,
    fixture.config.pgnExecutedWorkbookPath,
  );
  const workbookScenario = opened.parsed.scenarios.find(
    (item) => item.testCaseId === scenario.testCaseId,
  )!;
  const startedAt = new Date("2026-09-05T10:01:00.000Z");
  const finishedAt = new Date("2026-09-05T10:02:00.000Z");
  const successfullyCaptured = technicalStatus === "CAPTURED";
  const executions: ExecutedTurn[] = successfullyCaptured
    ? workbookScenario.turns.map((_, index) =>
        capturedExecution(workbookScenario, index),
      )
    : [
        {
          turn: workbookScenario.turns[0],
          technicalStatus,
          completedAt: new Date("2026-09-05T10:01:01.000Z"),
          botMessages: [],
          combinedResponse: "",
          error: `Fixture ${technicalStatus}`,
        },
      ];
  appendRecoveryTranscriptEvent(opened.workbook, {
    runId: fixture.runId,
    event: "SCENARIO_ATTEMPT_STARTED",
    message: "Scenario attempt 1 started from Turn 1.",
    timestamp: startedAt,
    scenario: workbookScenario,
  });
  applyScenarioExecution(
    opened.workbook,
    fixture.runId,
    workbookScenario,
    executions,
  );
  if (fixture.checkpoint.snapshot().mode === "retest") {
    const retest = getRetestRunMetadata(opened.workbook, fixture.runId)!;
    if (applyRetestStatusTransition(opened.workbook, workbookScenario, executions)) {
      retest.finishedIds.push(scenario.testCaseId);
    }
    retest.updatedAt = finishedAt;
    upsertRetestRunMetadata(opened.workbook, retest);
  }
  appendRecoveryTranscriptEvent(opened.workbook, {
    runId: fixture.runId,
    event: successfullyCaptured
      ? "SCENARIO_ATTEMPT_COMPLETED"
      : "SCENARIO_ATTEMPT_FAILED",
    message: `Scenario attempt 1 finished with ${technicalStatus}.`,
    timestamp: finishedAt,
    scenario: workbookScenario,
  });
  await saveExecutedPgnWorkbook(
    opened.workbook,
    fixture.config.pgnExecutedWorkbookPath,
  );
  if (updateCheckpoint) {
    await fixture.checkpoint.update((state) => {
      state.status = "INTERRUPTED";
      if (state.mode === "full" || successfullyCaptured) {
        state.completedScenarioIds.push(scenario.testCaseId);
        state.lastCompletedScenarioId = scenario.testCaseId;
      }
      state.scenarioAttempts.push({
        scenarioId: scenario.testCaseId,
        attempt: 1,
        status: successfullyCaptured ? "COMPLETED" : "FAILED",
        startedAt: startedAt.toISOString(),
        finishedAt: finishedAt.toISOString(),
        reason: successfullyCaptured
          ? undefined
          : "Scenario did not capture every turn",
      });
      state.metrics.executedScenarios += 1;
      if (successfullyCaptured) state.metrics.capturedScenarios += 1;
      else if (technicalStatus === "TIMEOUT") state.metrics.timeouts += 1;
      else state.metrics.technicalErrors += 1;
      state.updatedAt = finishedAt.toISOString();
      state.heartbeatAt = finishedAt.toISOString();
    });
  }
}

test("atomic recovery writes leave the previous checkpoint intact on failure", async (context) => {
  const root = await mkdtemp(path.join(tmpdir(), "pgn-recovery-atomic-"));
  context.after(() => rm(root, { recursive: true, force: true }));
  const target = path.join(root, "state.json");
  await atomicWriteJson(target, { generation: 1 });
  await assert.rejects(
    atomicWriteJson(
      target,
      { generation: 2 },
      {
        beforeRename: async () => {
          throw new Error("simulated power loss before rename");
        },
      },
    ),
    /simulated power loss/,
  );
  assert.deepEqual(JSON.parse(await readFile(target, "utf8")), {
    generation: 1,
  });
  assert.equal(
    (await readdir(root)).some((entry) => entry.endsWith(".tmp")),
    false,
  );
});

test("process liveness treats missing PIDs as dead and permission failures as alive", () => {
  assert.equal(
    isProcessAlive(123, () => {
      throw Object.assign(new Error("missing"), { code: "ESRCH" });
    }),
    false,
  );
  assert.equal(
    isProcessAlive(123, () => {
      throw Object.assign(new Error("denied"), { code: "EPERM" });
    }),
    true,
  );
  assert.equal(isProcessAlive(0, () => undefined), false);
});

test("run lock heartbeats block concurrency and dead owners are recovered", async (context) => {
  const root = await mkdtemp(path.join(tmpdir(), "pgn-recovery-lock-"));
  context.after(() => rm(root, { recursive: true, force: true }));
  const first = await acquireRunProcessLock(root, "first fixture");
  await first.heartbeat({ runId: "LOCK-FIXTURE", mode: "full" });
  const active = await inspectRunProcessLock(root);
  assert.equal(active.status, "active");
  if (active.status === "active") {
    assert.equal(active.record.runId, "LOCK-FIXTURE");
    assert.equal(active.record.pid, process.pid);
  }
  await assert.rejects(
    acquireRunProcessLock(root, "concurrent fixture"),
    /Another PGN process is active/,
  );
  await first.release();
  assert.equal((await inspectRunProcessLock(root)).status, "unlocked");

  const abandonedOwner = await acquireRunProcessLock(root, "dead fixture");
  const stale = await inspectRunProcessLock(root, {
    processAlive: () => false,
  });
  assert.equal(stale.status, "stale");
  const replacement = await acquireRunProcessLock(root, "replacement fixture", {
    processAlive: () => false,
  });
  assert.notEqual(replacement.token, abandonedOwner.token);
  await replacement.release();
  await abandonedOwner.release();
  assert.equal((await inspectRunProcessLock(root)).status, "unlocked");
});

test("a killed scenario process leaves a stale lock and recoverable Turn 1 checkpoint", async (context) => {
  const root = await mkdtemp(path.join(tmpdir(), "pgn-recovery-kill-"));
  context.after(() => rm(root, { recursive: true, force: true }));
  const moduleUrl = pathToFileURL(
    path.join(repositoryRoot, "src", "recovery", "run-state.ts"),
  ).href;
  const scenario: PgnTestScenario = {
    testCaseId: "KILL-001",
    sheetKind: "kb",
    sheetName: "Test Case Knowledge Base",
    sourceRowNumber: 2,
    category: "Fixture",
    rawStatus: "",
    turns: [
      {
        sheetName: "Test Case Knowledge Base",
        rowNumber: 2,
        turnNumber: 1,
        userInput: "First fixture turn",
      },
      {
        sheetName: "Test Case Knowledge Base",
        rowNumber: 3,
        turnNumber: 2,
        userInput: "Second fixture turn",
      },
    ],
  };
  const childCode = `
    import { acquireRunProcessLock, initializeRecoveryCheckpoint } from ${JSON.stringify(moduleUrl)};
    const root = ${JSON.stringify(root)};
    const scenario = ${JSON.stringify(scenario)};
    const lock = await acquireRunProcessLock(root, "killed fixture");
    const { checkpoint } = await initializeRecoveryCheckpoint({
      projectRoot: root,
      runId: "KILLED-RUN",
      mode: "full",
      sourceWorkbookPath: root + "/data/source.xlsx",
      executedWorkbookPath: root + "/reports/executed.xlsx",
      sourceWorkbookHash: "${"a".repeat(64)}",
      scenarios: [scenario],
      startedAt: new Date("2026-09-05T10:00:00.000Z"),
    });
    await lock.heartbeat({ runId: "KILLED-RUN", mode: "full" });
    await checkpoint.update((state) => {
      state.status = "RUNNING";
      state.activeScenarioId = "KILL-001";
      state.activeScenarioAttempt = 1;
      state.activeScenarioStartedAt = "2026-09-05T10:01:00.000Z";
      state.scenarioAttempts.push({
        scenarioId: "KILL-001",
        attempt: 1,
        status: "RUNNING",
        startedAt: "2026-09-05T10:01:00.000Z",
      });
    });
    console.log("READY");
    setInterval(() => {}, 1000);
  `;
  const child = spawn(
    process.execPath,
    ["--import", "tsx", "--input-type=module", "--eval", childCode],
    { cwd: repositoryRoot, stdio: ["ignore", "pipe", "pipe"] },
  );
  let stderr = "";
  child.stderr.setEncoding("utf8");
  child.stderr.on("data", (chunk: string) => {
    stderr += chunk;
  });
  await new Promise<void>((resolve, reject) => {
    let stdout = "";
    const timeout = setTimeout(() => {
      child.kill("SIGKILL");
      reject(new Error(`child did not become ready: ${stderr}`));
    }, 15_000);
    child.stdout.setEncoding("utf8");
    child.stdout.on("data", (chunk: string) => {
      stdout += chunk;
      if (stdout.includes("READY")) {
        clearTimeout(timeout);
        resolve();
      }
    });
    child.once("error", (error) => {
      clearTimeout(timeout);
      reject(error);
    });
    child.once("exit", () => {
      if (!stdout.includes("READY")) {
        clearTimeout(timeout);
        reject(new Error(`child exited before ready: ${stderr}`));
      }
    });
  });
  assert.equal(child.kill("SIGKILL"), true);
  await once(child, "exit");

  const lock = await inspectRunProcessLock(root);
  assert.equal(lock.status, "stale");
  const recovery = await discoverRecoveryRun(root);
  assert.equal(recovery.kind, "recoverable");
  if (recovery.kind === "recoverable") {
    assert.equal(recovery.state.runId, "KILLED-RUN");
    assert.equal(recovery.state.activeScenarioId, "KILL-001");
    assert.equal(recovery.state.completedScenarioIds.length, 0);
    assert.equal(
      selectRecoveryScenarios([scenario], recovery.state)[0]?.turns[0]
        .turnNumber,
      1,
    );
  }
  const replacement = await acquireRunProcessLock(root, "post-kill recovery");
  await replacement.release();
});

test("checkpoint discovery preserves terminal history without offering auto-resume", async (context) => {
  const fixture = await createFixture();
  context.after(() => rm(fixture.root, { recursive: true, force: true }));
  const discovered = await discoverRecoveryRun(fixture.root);
  assert.equal(discovered.kind, "recoverable");
  assert.deepEqual(
    discovered.kind === "recoverable"
      ? discovered.state.selectedScenarioIds
      : [],
    fixture.scenarios.map((scenario) => scenario.testCaseId),
  );
  await fixture.checkpoint.update((state) => {
    state.status = "ABANDONED";
    state.updatedAt = "2026-09-05T10:05:00.000Z";
    state.heartbeatAt = "2026-09-05T10:05:00.000Z";
  });
  assert.equal((await discoverRecoveryRun(fixture.root)).kind, "none");
  assert.equal(
    (await readRecoveryRun(fixture.root, fixture.runId)).state.status,
    "ABANDONED",
  );
  assert.equal(await pathExists(recoveryPaths(fixture.root).active), false);
});

test("malformed active recovery state is reported and never deleted", async (context) => {
  const fixture = await createFixture();
  context.after(() => rm(fixture.root, { recursive: true, force: true }));
  const statePath = recoveryPaths(fixture.root, fixture.runId).state!;
  await writeFile(statePath, "{ malformed", "utf8");
  const discovery = await discoverRecoveryRun(fixture.root);
  assert.equal(discovery.kind, "unreadable");
  assert.equal(await pathExists(statePath), true);
});

test("reconciliation preserves completed scenarios and restarts a partial multi-turn scenario", async (context) => {
  const fixture = await createFixture();
  context.after(() => rm(fixture.root, { recursive: true, force: true }));
  const completed = fixture.scenarios.find(
    (scenario) => scenario.turns.length === 1,
  )!;
  const partial = fixture.scenarios.find(
    (scenario) => scenario.turns.length > 1,
  )!;
  await recordScenarioAttempt(fixture, completed, true);

  const opened = await openExecutedPgnWorkbook(
    fixture.config.pgnSourceWorkbookPath,
    fixture.config.pgnExecutedWorkbookPath,
  );
  const workbookPartial = opened.parsed.scenarios.find(
    (scenario) => scenario.testCaseId === partial.testCaseId,
  )!;
  appendRecoveryTranscriptEvent(opened.workbook, {
    runId: fixture.runId,
    event: "SCENARIO_ATTEMPT_STARTED",
    message: "Scenario attempt 1 started from Turn 1.",
    timestamp: new Date("2026-09-05T10:03:00.000Z"),
    scenario: workbookPartial,
  });
  applyScenarioExecution(opened.workbook, fixture.runId, workbookPartial, [
    capturedExecution(workbookPartial, 0),
  ]);
  await saveExecutedPgnWorkbook(
    opened.workbook,
    fixture.config.pgnExecutedWorkbookPath,
  );
  await fixture.checkpoint.update((state) => {
    state.status = "INTERRUPTED";
    state.activeScenarioId = partial.testCaseId;
    state.activeScenarioAttempt = 1;
    state.activeScenarioStartedAt = "2026-09-05T10:03:00.000Z";
    state.scenarioAttempts.push({
      scenarioId: partial.testCaseId,
      attempt: 1,
      status: "RUNNING",
      startedAt: "2026-09-05T10:03:00.000Z",
    });
    state.updatedAt = "2026-09-05T10:04:00.000Z";
    state.heartbeatAt = "2026-09-05T10:04:00.000Z";
  });

  const reloaded = await openExecutedPgnWorkbook(
    fixture.config.pgnSourceWorkbookPath,
    fixture.config.pgnExecutedWorkbookPath,
  );
  const recovered = await readRecoveryRun(fixture.root, fixture.runId);
  const selected = fixture.scenarios.map(
    (scenario) =>
      reloaded.parsed.scenarios.find(
        (item) => item.testCaseId === scenario.testCaseId,
      )!,
  );
  const reconciliation = reconcileRecoveryArtifacts(
    reloaded.workbook,
    selected,
    recovered.state,
    recovered.manifest,
  );
  assert.deepEqual(reconciliation.mismatchedScenarioIds, []);
  assert.deepEqual(reconciliation.safeCompletedIds, [completed.testCaseId]);
  assert.equal(reconciliation.interruptedScenarioId, partial.testCaseId);
  assert.equal(reconciliation.restartInterruptedScenarioFromTurnOne, true);
  const remaining = selectRecoveryScenarios(selected, recovered.state);
  assert.deepEqual(
    remaining.map((scenario) => scenario.testCaseId),
    [partial.testCaseId],
  );
  assert.equal(remaining[0].turns[0].turnNumber, 1);
});

for (const mode of ["full", "retest"] as const) {
  for (const technicalStatus of ["TIMEOUT", "SEND_ERROR", "CHAT_ERROR"] as const) {
    test(`${mode} recovery reloads ${technicalStatus} outcomes across KB and negative layouts`, async (context) => {
      const fixture = await createFixture(mode, (scenarios, allScenarios) => {
        const failures = representativeScenarios(allScenarios);
        return [
          ...failures,
          scenarios.find(
            (scenario) =>
              scenario.sheetKind === "kb" &&
              scenario.turns.length === 1 &&
              !failures.includes(scenario),
          )!,
        ];
      });
      context.after(() => rm(fixture.root, { recursive: true, force: true }));
      const failures = fixture.scenarios.slice(0, -1);
      const captured = fixture.scenarios.at(-1)!;
      for (const scenario of failures) {
        await recordScenarioAttempt(fixture, scenario, true, technicalStatus);
      }
      await recordScenarioAttempt(fixture, captured, true);

      const reloaded = await openExecutedPgnWorkbook(
        fixture.config.pgnSourceWorkbookPath,
        fixture.config.pgnExecutedWorkbookPath,
      );
      const recovered = await readRecoveryRun(fixture.root, fixture.runId);
      const selectedIds = fixture.scenarios.map((scenario) => scenario.testCaseId);
      const failedIds = failures.map((scenario) => scenario.testCaseId);
      const completedIds = mode === "full" ? selectedIds : [captured.testCaseId];
      const reconciliation = reconcileRecoveryArtifacts(
        reloaded.workbook,
        fixture.scenarios,
        recovered.state,
        recovered.manifest,
      );
      assert.equal(recovered.state.status, "INTERRUPTED");
      assert.equal(recovered.state.runId, fixture.runId);
      assert.deepEqual(recovered.state.selectedScenarioIds, selectedIds);
      assert.deepEqual(
        recovered.manifest.scenarios.map((scenario) => scenario.testCaseId),
        selectedIds,
      );
      assert.deepEqual(
        recovered.state.scenarioAttempts.map(
          ({ scenarioId, attempt, status }) => ({ scenarioId, attempt, status }),
        ),
        [
          ...failedIds.map((scenarioId) => ({
            scenarioId, attempt: 1, status: "FAILED",
          })),
          { scenarioId: captured.testCaseId, attempt: 1, status: "COMPLETED" },
        ],
      );
      assert.deepEqual(reconciliation.checkpointCompletedIds, completedIds);
      assert.deepEqual(reconciliation.workbookCompletedIds, completedIds);
      assert.deepEqual(reconciliation.transcriptCompletedIds, completedIds);
      assert.deepEqual(reconciliation.artifactConfirmedIds, completedIds);
      assert.deepEqual(reconciliation.safeCompletedIds, completedIds);
      assert.deepEqual(reconciliation.mismatchedScenarioIds, []);
      assert.equal(
        reconciliation.nextScenarioId,
        mode === "full" ? undefined : failedIds[0],
      );
      const remaining = selectRecoveryScenarios(
        reloaded.parsed.scenarios,
        recovered.state,
      );
      assert.deepEqual(
        remaining.map((scenario) => scenario.testCaseId),
        mode === "full" ? [] : failedIds,
      );
      assert(remaining.every((scenario) => scenario.turns[0].turnNumber === 1));

      const transcript = reloaded.workbook.getWorksheet("Execution Transcript")!;
      for (const scenario of failures) {
        assert.equal(isScenarioComplete(reloaded.workbook, scenario), false);
        const row = reloaded.workbook
          .getWorksheet(scenario.sheetName)!
          .getRow(scenario.sourceRowNumber);
        const date = row.getCell(scenario.sheetKind === "kb" ? 11 : 10).value;
        assert(date instanceof Date);
        assert(
          row.getCell(scenario.sheetKind === "kb" ? 13 : 12).text.includes(
            `[Technical execution ${date.toISOString()}] Turn 1: ${technicalStatus}`,
          ),
        );
        assert.deepEqual(
          transcript
            .getRows(2, transcript.rowCount - 1)!
            .filter(
              (row) =>
                row.getCell(1).text === fixture.runId &&
                row.getCell(2).text === scenario.testCaseId &&
                row.getCell(6).text === "USER",
            )
            .map((row) => [row.getCell(5).value, row.getCell(11).text]),
          [[1, technicalStatus]],
        );
        if (mode === "retest") {
          assert.equal(
            reloaded.parsed.scenarios.find(
              (item) => item.testCaseId === scenario.testCaseId,
            )!.status,
            PGN_TEST_STATUSES.ReadyForRetest,
          );
        }
      }
      if (mode === "retest") {
        assert.equal(fixture.config.googleDriveEvidenceEnabled, true);
        const retest = getRetestRunMetadata(reloaded.workbook, fixture.runId)!;
        assert.deepEqual(retest.selectedIds, selectedIds);
        assert.deepEqual(retest.finishedIds, [captured.testCaseId]);
        assert.equal(retest.state, "IN_PROGRESS");
        assert.equal(
          reloaded.parsed.scenarios.find(
            (item) => item.testCaseId === captured.testCaseId,
          )!.status,
          PGN_TEST_STATUSES.PendingEvaluation,
        );
      }
      const validation = await validateRecoveryRun(fixture.config, fixture.runId, {
        checkDriveAccess: false,
      });
      assert.equal(validation.ready, true, JSON.stringify(validation.checks));
      if (mode === "retest") {
        const drive = validation.checks.find((check) => check.id === "drive")!;
        assert.equal(drive.status, "info");
        assert.match(drive.detail, /access not checked/);
      }
    });
  }
}

test("an older technical note cannot complete an interrupted same-run KB rerun", async (context) => {
  const fixture = await createFixture("full", (scenarios) => [
    scenarios.find(
      (scenario) => scenario.sheetKind === "kb" && scenario.turns.length > 1,
    )!,
  ]);
  context.after(() => rm(fixture.root, { recursive: true, force: true }));
  const scenario = fixture.scenarios[0];
  await recordScenarioAttempt(fixture, scenario, true, "TIMEOUT");
  const opened = await openExecutedPgnWorkbook(
    fixture.config.pgnSourceWorkbookPath,
    fixture.config.pgnExecutedWorkbookPath,
  );
  const oldNote = opened.workbook
    .getWorksheet(scenario.sheetName)!
    .getCell(scenario.sourceRowNumber, 13).text;
  assert(oldNote.includes("[Technical execution 2026-09-05T10:01:01.000Z]"));

  const startedAt = new Date("2026-09-05T10:03:00.000Z");
  const capturedAt = new Date("2026-09-05T10:03:01.000Z");
  appendRecoveryTranscriptEvent(opened.workbook, {
    runId: fixture.runId,
    event: "SCENARIO_ATTEMPT_STARTED",
    message: "Scenario attempt 2 started from Turn 1.",
    timestamp: startedAt,
    scenario,
  });
  applyScenarioExecution(opened.workbook, fixture.runId, scenario, [
    capturedExecution(scenario, 0, capturedAt),
  ]);
  // Archived turns must not fill the new attempt's missing turns.
  scenario.turns.forEach((_, index) => {
    appendLatestTurnExecution(opened.workbook, "OLDER-RECOVERY-RUN", scenario, [
      capturedExecution(scenario, index, new Date("2026-09-05T09:59:01.000Z")),
    ]);
  });
  await saveExecutedPgnWorkbook(
    opened.workbook,
    fixture.config.pgnExecutedWorkbookPath,
  );
  await fixture.checkpoint.update((state) => {
    state.status = "INTERRUPTED";
    state.completedScenarioIds = [];
    state.lastCompletedScenarioId = undefined;
    state.activeScenarioId = scenario.testCaseId;
    state.activeScenarioAttempt = 2;
    state.activeScenarioStartedAt = startedAt.toISOString();
    state.scenarioAttempts.push({
      scenarioId: scenario.testCaseId,
      attempt: 2,
      status: "RUNNING",
      startedAt: startedAt.toISOString(),
    });
    state.updatedAt = capturedAt.toISOString();
    state.heartbeatAt = capturedAt.toISOString();
  });

  const reloaded = await openExecutedPgnWorkbook(
    fixture.config.pgnSourceWorkbookPath,
    fixture.config.pgnExecutedWorkbookPath,
  );
  const recovered = await readRecoveryRun(fixture.root, fixture.runId);
  const row = reloaded.workbook
    .getWorksheet(scenario.sheetName)!
    .getRow(scenario.sourceRowNumber);
  assert.equal(row.getCell(13).text, oldNote);
  assert.deepEqual(row.getCell(11).value, capturedAt);
  assert.equal(isScenarioPartiallyComplete(reloaded.workbook, scenario), true);
  assert.deepEqual(
    recovered.state.scenarioAttempts.map((attempt) => attempt.status),
    ["FAILED", "RUNNING"],
  );
  const reconciliation = reconcileRecoveryArtifacts(
    reloaded.workbook,
    fixture.scenarios,
    recovered.state,
    recovered.manifest,
  );
  assert.deepEqual(reconciliation.transcriptCompletedIds, []);
  assert.deepEqual(reconciliation.workbookCompletedIds, []);
  assert.deepEqual(reconciliation.artifactConfirmedIds, []);
  assert.deepEqual(reconciliation.safeCompletedIds, []);
  assert.deepEqual(reconciliation.mismatchedScenarioIds, []);
  assert.equal(reconciliation.interruptedScenarioId, scenario.testCaseId);
  assert.equal(reconciliation.restartInterruptedScenarioFromTurnOne, true);
  const remaining = selectRecoveryScenarios(
    reloaded.parsed.scenarios,
    recovered.state,
  );
  assert.deepEqual(remaining.map((item) => item.testCaseId), [scenario.testCaseId]);
  assert.deepEqual(remaining[0].turns, scenario.turns);
  assert.equal(remaining[0].turns[0].turnNumber, 1);
  const validation = await validateRecoveryRun(fixture.config, fixture.runId, {
    checkDriveAccess: false,
  });
  assert.equal(validation.ready, true, JSON.stringify(validation.checks));
});

for (const missingArtifact of [
  "checkpoint", "execution date", "transcript",
] as const) {
  test(`technical outcomes require reconciliation when the ${missingArtifact} is missing`, async (context) => {
    const fixture = await createFixture("full", (_, allScenarios) =>
      representativeScenarios(allScenarios),
    );
    context.after(() => rm(fixture.root, { recursive: true, force: true }));
    for (const scenario of fixture.scenarios) {
      await recordScenarioAttempt(
        fixture,
        scenario,
        missingArtifact !== "checkpoint",
        "TIMEOUT",
      );
    }
    await fixture.checkpoint.update((state) => {
      state.status = "INTERRUPTED";
    });
    if (missingArtifact !== "checkpoint") {
      const opened = await openExecutedPgnWorkbook(
        fixture.config.pgnSourceWorkbookPath,
        fixture.config.pgnExecutedWorkbookPath,
      );
      if (missingArtifact === "execution date") {
        for (const scenario of fixture.scenarios) {
          applyScenarioResults(opened.workbook, scenario, []);
        }
      } else {
        opened.workbook.removeWorksheet("Execution Transcript");
      }
      await saveExecutedPgnWorkbook(
        opened.workbook,
        fixture.config.pgnExecutedWorkbookPath,
      );
    }
    const reloaded = await openExecutedPgnWorkbook(
      fixture.config.pgnSourceWorkbookPath,
      fixture.config.pgnExecutedWorkbookPath,
    );
    const recovered = await readRecoveryRun(fixture.root, fixture.runId);
    const selectedIds = fixture.scenarios.map((scenario) => scenario.testCaseId);
    if (missingArtifact === "execution date") {
      for (const scenario of fixture.scenarios) {
        const row = reloaded.workbook
          .getWorksheet(scenario.sheetName)!
          .getRow(scenario.sourceRowNumber);
        assert.equal(row.getCell(scenario.sheetKind === "kb" ? 11 : 10).value, null);
        assert(
          row.getCell(scenario.sheetKind === "kb" ? 13 : 12).text.includes(
            "[Technical execution 2026-09-05T10:01:01.000Z]",
          ),
        );
      }
    }
    const reconciliation = reconcileRecoveryArtifacts(
      reloaded.workbook,
      fixture.scenarios,
      recovered.state,
      recovered.manifest,
    );
    assert.deepEqual(
      reconciliation.checkpointCompletedIds,
      missingArtifact === "checkpoint" ? [] : selectedIds,
    );
    assert.deepEqual(
      reconciliation.workbookCompletedIds,
      missingArtifact === "execution date" ? [] : selectedIds,
    );
    assert.deepEqual(
      reconciliation.transcriptCompletedIds,
      missingArtifact === "transcript" ? [] : selectedIds,
    );
    assert.deepEqual(
      reconciliation.artifactConfirmedIds,
      missingArtifact === "checkpoint" ? selectedIds : [],
    );
    assert.deepEqual(reconciliation.safeCompletedIds, []);
    assert.deepEqual(reconciliation.mismatchedScenarioIds, selectedIds);
    const validation = await validateRecoveryRun(fixture.config, fixture.runId, {
      checkDriveAccess: false,
    });
    assert.equal(validation.ready, false);
    assert.equal(
      validation.checks.find((check) => check.id === "reconciliation")?.status,
      "error",
    );
    assert.deepEqual(validation.reconciliation?.mismatchedScenarioIds, selectedIds);
  });
}

test("operator reconciliation records a durable rerun decision for crash-window disagreement", async (context) => {
  const fixture = await createFixture("full", (scenarios) => [
    scenarios.find(
      (scenario) => scenario.sheetKind === "kb" && scenario.turns.length === 1,
    )!,
  ]);
  context.after(() => rm(fixture.root, { recursive: true, force: true }));
  await recordScenarioAttempt(fixture, fixture.scenarios[0], false);
  await fixture.checkpoint.update((state) => {
    state.status = "INTERRUPTED";
    state.updatedAt = "2026-09-05T10:03:00.000Z";
    state.heartbeatAt = "2026-09-05T10:03:00.000Z";
  });

  const blocked = await validateRecoveryRun(fixture.config, fixture.runId, {
    checkDriveAccess: false,
  });
  assert.equal(blocked.ready, false);
  assert.deepEqual(blocked.reconciliation?.mismatchedScenarioIds, [
    fixture.scenarios[0].testCaseId,
  ]);

  await repairRecoveryProgress(fixture.config, fixture.runId, "rerun");
  const repaired = await validateRecoveryRun(fixture.config, fixture.runId, {
    checkDriveAccess: false,
  });
  assert.equal(repaired.ready, true);
  assert.deepEqual(repaired.reconciliation?.mismatchedScenarioIds, []);
  assert.deepEqual(repaired.reconciliation?.reconciledScenarioIds, [
    fixture.scenarios[0].testCaseId,
  ]);
  assert.deepEqual(
    selectRecoveryScenarios(fixture.scenarios, repaired.state).map(
      (scenario) => scenario.testCaseId,
    ),
    [fixture.scenarios[0].testCaseId],
  );
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(fixture.config.pgnExecutedWorkbookPath);
  const transcript = workbook.getWorksheet("Execution Transcript")!;
  assert(
    transcript
      .getColumn(11)
      .values.some((value) => value === "RECOVERY_RECONCILED"),
  );
});

test("validation warns for formatting drift and blocks selected-input drift", async (context) => {
  const fixture = await createFixture("full", (scenarios) => [
    scenarios.find(
      (scenario) => scenario.sheetKind === "kb" && scenario.turns.length === 1,
    )!,
  ]);
  context.after(() => rm(fixture.root, { recursive: true, force: true }));
  await fixture.checkpoint.update((state) => {
    state.status = "INTERRUPTED";
  });

  const source = new ExcelJS.Workbook();
  await source.xlsx.readFile(fixture.config.pgnSourceWorkbookPath);
  source.getWorksheet(fixture.scenarios[0].sheetName)!.getCell(1, 1).font = {
    bold: false,
    italic: true,
  };
  await source.xlsx.writeFile(fixture.config.pgnSourceWorkbookPath);
  const formattingOnly = await validateRecoveryRun(
    fixture.config,
    fixture.runId,
    { checkDriveAccess: false },
  );
  assert.equal(formattingOnly.sourceDrift, "formatting-only");
  assert.equal(formattingOnly.ready, true);

  const changed = new ExcelJS.Workbook();
  await changed.xlsx.readFile(fixture.config.pgnSourceWorkbookPath);
  changed
    .getWorksheet(fixture.scenarios[0].sheetName)!
    .getCell(fixture.scenarios[0].turns[0].rowNumber, 8).value =
    "Structurally changed fixture input";
  await changed.xlsx.writeFile(fixture.config.pgnSourceWorkbookPath);
  const structural = await validateRecoveryRun(
    fixture.config,
    fixture.runId,
    { checkDriveAccess: false },
  );
  assert.equal(structural.sourceDrift, "structural");
  assert.equal(structural.ready, false);
});

test("recovery validation only reads Drive and reuses the stored folder identity", async (context) => {
  const fixture = await createFixture("full", (scenarios) => [
    scenarios.find(
      (scenario) => scenario.sheetKind === "kb" && scenario.turns.length === 1,
    )!,
  ]);
  context.after(() => rm(fixture.root, { recursive: true, force: true }));
  const config: AppConfig = {
    ...fixture.config,
    googleDriveEvidenceEnabled: true,
    googleDriveEvidenceParentFolderId: "fixture-parent-folder",
  };
  await fixture.checkpoint.update((state) => {
    state.status = "INTERRUPTED";
    state.driveRunFolderId = "fixture-run-folder";
    state.driveRunFolderUrl =
      "https://drive.google.com/drive/folders/fixture-run-folder";
  });
  const stateBefore = await readFile(
    recoveryPaths(fixture.root, fixture.runId).state!,
  );
  const workbookBefore = await readFile(
    fixture.config.pgnExecutedWorkbookPath,
  );
  const calls: string[] = [];
  const item = (id: string, name: string): DriveEvidenceItem => ({
    id,
    name,
    webViewLink: `https://drive.google.com/drive/folders/${id}`,
    reused: true,
  });
  const publisher: EvidenceDrivePublisher = {
    parentFolderId: "fixture-parent-folder",
    folderPrefix: config.googleDriveEvidenceFolderPrefix,
    retestFolderPrefix: config.googleDriveRetestFolderPrefix,
    validateParentFolder: async () => {
      calls.push("validate-parent");
      return item("fixture-parent-folder", "Fixture parent");
    },
    validateRunFolder: async (folderId, expectedName) => {
      calls.push(`validate-run:${folderId}:${expectedName}`);
      return item(folderId, expectedName ?? "Fixture run");
    },
    ensureRunFolder: async () => {
      throw new Error("validation must not create a Drive folder");
    },
    uploadPng: async () => {
      throw new Error("validation must not upload evidence");
    },
  };
  const validation = await validateRecoveryRun(config, fixture.runId, {
    drivePublisher: publisher,
  });
  assert.equal(validation.ready, true);
  assert.deepEqual(calls, [
    "validate-parent",
    `validate-run:fixture-run-folder:${config.googleDriveEvidenceFolderPrefix}-${fixture.runId}`,
  ]);
  assert(
    (await readFile(recoveryPaths(fixture.root, fixture.runId).state!)).equals(
      stateBefore,
    ),
  );
  assert(
    (await readFile(fixture.config.pgnExecutedWorkbookPath)).equals(
      workbookBefore,
    ),
  );
});

test("skip and abandon are explicit, audited, and preserve recovery history", async (context) => {
  const fixture = await createFixture();
  context.after(() => rm(fixture.root, { recursive: true, force: true }));
  const active = fixture.scenarios[0];
  await fixture.checkpoint.update((state) => {
    state.status = "INTERRUPTED";
    state.activeScenarioId = active.testCaseId;
    state.activeScenarioAttempt = 1;
    state.activeScenarioStartedAt = "2026-09-05T10:01:00.000Z";
    state.scenarioAttempts.push({
      scenarioId: active.testCaseId,
      attempt: 1,
      status: "RUNNING",
      startedAt: "2026-09-05T10:01:00.000Z",
    });
  });
  const skipped = await skipRecoveryScenario(
    fixture.config,
    fixture.runId,
  );
  assert.equal(skipped.scenarioId, active.testCaseId);
  const afterSkip = (await readRecoveryRun(fixture.root, fixture.runId)).state;
  assert.deepEqual(afterSkip.skippedScenarioIds, [active.testCaseId]);
  assert.equal(afterSkip.scenarioAttempts[0].status, "SKIPPED_BY_OPERATOR");
  assert.equal(afterSkip.activeScenarioId, undefined);

  await abandonRecoveryRun(fixture.config, fixture.runId);
  const afterAbandon = await readRecoveryRun(fixture.root, fixture.runId);
  assert.equal(afterAbandon.state.status, "ABANDONED");
  assert.equal((await discoverRecoveryRun(fixture.root)).kind, "none");
  assert.equal(
    await pathExists(recoveryPaths(fixture.root, fixture.runId).manifest!),
    true,
  );
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(fixture.config.pgnExecutedWorkbookPath);
  const events = workbook
    .getWorksheet("Execution Transcript")!
    .getColumn(11).values;
  assert(events.includes("SCENARIO_SKIPPED_BY_OPERATOR"));
  assert(events.includes("RUN_ABANDONED"));
});

test("fresh preparation refuses recoverable state without changing the workbook", async (context) => {
  const fixture = await createFixture();
  context.after(() => rm(fixture.root, { recursive: true, force: true }));
  const before = await readFile(fixture.config.pgnExecutedWorkbookPath);
  await assert.rejects(
    prepareFreshPgnWorkbook(fixture.config),
    new RegExp(`Recoverable PGN run ${fixture.runId}`),
  );
  assert((await readFile(fixture.config.pgnExecutedWorkbookPath)).equals(before));
  assert.equal(await pathExists(fixture.config.reportArchiveDir), false);
});

test("retest recovery keeps the original selection, Run ID, and folder snapshot", async (context) => {
  const fixture = await createFixture("retest");
  context.after(() => rm(fixture.root, { recursive: true, force: true }));
  const [finished, remaining] = fixture.scenarios;
  await fixture.checkpoint.update((state) => {
    state.status = "INTERRUPTED";
    state.completedScenarioIds = [finished.testCaseId];
    state.lastCompletedScenarioId = finished.testCaseId;
    state.driveRunFolderId = "original-retest-folder";
    state.driveRunFolderUrl =
      "https://drive.google.com/drive/folders/original-retest-folder";
  });
  const recovered = await readRecoveryRun(fixture.root, fixture.runId);
  assert.equal(recovered.state.runId, fixture.runId);
  assert.deepEqual(recovered.state.selectedScenarioIds, [
    finished.testCaseId,
    remaining.testCaseId,
  ]);
  assert.equal(recovered.state.driveRunFolderId, "original-retest-folder");
  assert.deepEqual(
    selectRecoveryScenarios(fixture.scenarios, recovered.state).map(
      (scenario) => scenario.testCaseId,
    ),
    [remaining.testCaseId],
  );
});

test("resume validation CLI accepts only an optional explicit Run ID", () => {
  assert.deepEqual(parseResumeValidationArgs([]), {});
  assert.deepEqual(parseResumeValidationArgs(["--run", "RUN-123"]), {
    runId: "RUN-123",
  });
  assert.throws(
    () => parseResumeValidationArgs(["--unknown"]),
    /Usage:/,
  );
  assert.throws(
    () => parseResumeValidationArgs(["--run", ""]),
    /Usage:/,
  );
  assert.throws(
    () =>
      assertResumeOptionsCompatible(
        parseCliOptions(["--resume", "RUN-123", "--limit", "1"]),
      ),
    /original selection snapshot/,
  );
});
