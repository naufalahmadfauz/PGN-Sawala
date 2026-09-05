import { randomUUID } from "node:crypto";
import { lstat, mkdir, readdir, rm } from "node:fs/promises";
import path from "node:path";
import ExcelJS from "exceljs";
import { REPOSITORY_ROOT } from "../environment";
import {
  EVIDENCE_MIGRATION_VERSION,
  upsertEvidenceFileMetadata,
  upsertEvidenceRunMetadata,
} from "../excel/evidence-workbook";
import { KB_HEADERS, NEGATIVE_HEADERS, parsePgnWorkbook } from "../excel/pgn-workbook-loader";
import { assertPgnWorkbookValid } from "../excel/pgn-workbook-validator";
import {
  appendRecoveryTranscriptEvent,
  applyScenarioExecution,
  applyScenarioResults,
  openExecutedPgnWorkbook,
  saveExecutedPgnWorkbook,
} from "../excel/pgn-workbook-writer";
import { KB_SHEET_NAME, NEGATIVE_SHEET_NAME, type ExecutedTurn } from "../excel/pgn-types";
import {
  applyRetestStatusTransition,
  upsertRetestRunMetadata,
} from "../excel/retest-workbook";
import { assertSafeDemoStorage, recoveryDemoPaths } from "./demo-safety";
import {
  acquireRunProcessLock,
  discoverRecoveryRun,
  hashFile,
  initializeRecoveryCheckpoint,
  openRecoveryCheckpoint,
  readRecoveryRun,
  recoveryPaths,
} from "./run-state";

export interface RecoveryDemoOptions {
  mode?: "full" | "retest";
  sourceDrift?: boolean;
  mismatch?: boolean;
}

async function demoIds(projectRoot: string): Promise<string[]> {
  const directory = path.join(recoveryPaths(projectRoot).root, "demos");
  const entries = await readdir(directory, { withFileTypes: true }).catch((error) => {
    if ((error as NodeJS.ErrnoException).code === "ENOENT") return [];
    throw error;
  });
  if (entries.some((entry) => !entry.isDirectory() || !entry.name.startsWith("DEMO-RECOVERY-"))) {
    throw new Error("Unexpected demo storage entries; refusing automatic creation or cleanup");
  }
  return entries.map((entry) => entry.name);
}

export async function createRecoveryDemo(
  projectRoot = REPOSITORY_ROOT,
  options: RecoveryDemoOptions = {},
) {
  await assertSafeDemoStorage(projectRoot);
  const existing = await demoIds(projectRoot);
  if (existing.length) {
    throw new Error(`A recovery demo already exists: ${existing.join(", ")}. Inspect it with npm run pgn, or run npm run recovery:demo:reset before creating another demo.`);
  }
  const discovery = await discoverRecoveryRun(projectRoot);
  if (discovery.kind !== "none" || discovery.lock.status !== "unlocked") {
    throw new Error("Existing recovery state or process lock found. Demo creation will not overwrite real or demo run data.");
  }
  const lock = await acquireRunProcessLock(projectRoot, "Creating local recovery demo", { recoverStale: false });
  try {
    // Recheck under the same process lock used by real runners.
    const paths = recoveryPaths(projectRoot);
    const active = await lstat(paths.active).catch((error) => {
      if ((error as NodeJS.ErrnoException).code === "ENOENT") return undefined;
      throw error;
    });
    if (active || (await demoIds(projectRoot)).length) {
      throw new Error("Recovery state appeared during demo creation; nothing was overwritten");
    }
    const now = new Date();
    const startedAt = new Date(now.getTime() - 10 * 60_000);
    const interruptedAt = new Date(now.getTime() - 5 * 60_000);
    const runId = `DEMO-RECOVERY-${now.toISOString().replace(/[-:.]/g, "")}-${randomUUID().slice(0, 8)}`;
    const demo = recoveryDemoPaths(projectRoot, runId);
    await assertSafeDemoStorage(projectRoot, runId);
    await mkdir(path.dirname(demo.directory), { recursive: true });
    await mkdir(demo.directory);
    const mode = options.mode ?? "full";
    const source = new ExcelJS.Workbook();
    const sheet = source.addWorksheet(KB_SHEET_NAME);
    sheet.addRow(KB_HEADERS);
    source.addWorksheet(NEGATIVE_SHEET_NAME).addRow(NEGATIVE_HEADERS);
    for (let index = 1; index <= 10; index += 1) {
      const id = `DEMO-KB-${String(index).padStart(3, "0")}`;
      sheet.addRow([
        index, "Synthetic demo", id, "Demo operator", "Local recovery preview",
        "Synthetic response only", 1, `Synthetic input for ${id}, Turn 1`,
        null, null, null, mode === "retest" ? "Ready for Re-test" : "",
      ]);
      if (index === 4 || (mode === "retest" && index === 6)) {
        sheet.addRow([null, null, null, null, null, null, 2, `Synthetic input for ${id}, Turn 2`]);
      }
    }
    const parsed = parsePgnWorkbook(source);
    assertPgnWorkbookValid(parsed);
    const scenarios = mode === "retest"
      ? parsed.scenarios.filter((scenario) => ["DEMO-KB-002", "DEMO-KB-004", "DEMO-KB-006", "DEMO-KB-008"].includes(scenario.testCaseId))
      : parsed.scenarios;
    const completed = scenarios.slice(0, mode === "retest" ? 2 : 3);
    const activeScenario = scenarios[completed.length];
    await saveExecutedPgnWorkbook(source, demo.sourceWorkbookPath);
    const { checkpoint } = await initializeRecoveryCheckpoint({
      projectRoot, runId, isDemo: true, mode,
      sourceWorkbookPath: demo.sourceWorkbookPath,
      executedWorkbookPath: demo.executedWorkbookPath,
      sourceWorkbookHash: await hashFile(demo.sourceWorkbookPath),
      scenarios, startedAt,
    });
    await lock.heartbeat({ runId, mode });
    const { workbook } = await openExecutedPgnWorkbook(demo.sourceWorkbookPath, demo.executedWorkbookPath);
    appendRecoveryTranscriptEvent(workbook, {
      runId, event: "RUN_PREPARED", timestamp: startedAt,
      message: "DEMO: Synthetic local recovery fixture. No live services were used.",
    });
    let capturedTurns = 0;
    for (const scenario of [...completed, activeScenario]) {
      const isComplete = completed.includes(scenario);
      appendRecoveryTranscriptEvent(workbook, {
        runId, scenario, event: "SCENARIO_ATTEMPT_STARTED", timestamp: startedAt,
        message: "Synthetic scenario attempt 1 started from Turn 1.",
      });
      const executions: ExecutedTurn[] = (isComplete ? scenario.turns : scenario.turns.slice(0, 1)).map((turn) => ({
        turn, technicalStatus: "CAPTURED", sentAt: startedAt, completedAt: interruptedAt,
        combinedResponse: `Synthetic response for ${scenario.testCaseId}, Turn ${turn.turnNumber}`,
        botMessages: [{ sequence: 1, message: "Synthetic bot response; no message was sent.", timestamp: interruptedAt }],
        firstResponseMs: 500, totalResponseMs: 1_000,
        evidenceStatus: "EVIDENCE_MISSING",
      }));
      capturedTurns += executions.length;
      applyScenarioExecution(workbook, runId, scenario, executions);
      for (const execution of executions) {
        upsertEvidenceFileMetadata(workbook, {
          evidenceKey: `${runId}|${scenario.testCaseId}|${execution.turn.turnNumber}`,
          runId, testCaseId: scenario.testCaseId, turnNumber: execution.turn.turnNumber,
          driveFileName: `DEMO-${scenario.testCaseId}-${execution.turn.turnNumber}.png`,
          status: "EVIDENCE_MISSING",
        });
      }
      if (isComplete) {
        if (mode === "retest") applyRetestStatusTransition(workbook, scenario, executions);
        appendRecoveryTranscriptEvent(workbook, {
          runId, scenario, event: "SCENARIO_ATTEMPT_COMPLETED", timestamp: interruptedAt,
          message: "Synthetic scenario attempt 1 completed and checkpointed.",
        });
      }
    }
    const interruptionReason = "Interrupted during Turn 2 of 2; demo only.";
    appendRecoveryTranscriptEvent(workbook, {
      runId, scenario: activeScenario, event: "RUN_INTERRUPTED", timestamp: interruptedAt,
      message: interruptionReason,
    });
    upsertEvidenceRunMetadata(workbook, {
      runId, folderId: "", folderUrl: "", migrationVersion: EVIDENCE_MIGRATION_VERSION,
      timestamp: startedAt, mode: mode === "retest" ? "RETEST" : "FUTURE",
    });
    if (mode === "retest") {
      upsertRetestRunMetadata(workbook, {
        runId, state: "IN_PROGRESS", startedAt, updatedAt: interruptedAt,
        selectedIds: scenarios.map((scenario) => scenario.testCaseId),
        finishedIds: (options.mismatch ? completed.slice(0, -1) : completed).map((scenario) => scenario.testCaseId),
      });
    }
    if (options.mismatch) applyScenarioResults(workbook, completed.at(-1)!, []);
    await saveExecutedPgnWorkbook(workbook, demo.executedWorkbookPath);
    await checkpoint.update((state) => {
      state.status = "INTERRUPTED";
      state.completedScenarioIds = completed.map((scenario) => scenario.testCaseId);
      state.lastCompletedScenarioId = completed.at(-1)!.testCaseId;
      state.activeScenarioId = activeScenario.testCaseId;
      state.activeScenarioAttempt = 1;
      state.activeScenarioStartedAt = startedAt.toISOString();
      state.scenarioAttempts = [...completed, activeScenario].map((scenario) => ({
        scenarioId: scenario.testCaseId, attempt: 1,
        status: completed.includes(scenario) ? "COMPLETED" : "RUNNING",
        startedAt: startedAt.toISOString(),
        ...(completed.includes(scenario) ? { finishedAt: interruptedAt.toISOString() } : {}),
      }));
      state.metrics.executedScenarios = completed.length;
      state.metrics.capturedScenarios = completed.length;
      state.workbookProgress = `Synthetic transcript saved: ${capturedTurns} captured turns; no screenshots or remote evidence`;
      state.interruptionReason = interruptionReason;
      state.interruptedAt = interruptedAt.toISOString();
      state.updatedAt = interruptedAt.toISOString();
      state.heartbeatAt = interruptedAt.toISOString();
    });
    if (options.sourceDrift) {
      sheet.getCell(1, 1).font = { bold: true, italic: true };
      await saveExecutedPgnWorkbook(source, demo.sourceWorkbookPath);
    }
    return readRecoveryRun(projectRoot, runId);
  } finally {
    // A released lock is sufficient for real discovery. No fake PID or signals.
    await lock.release();
  }
}

export async function resetRecoveryDemos(projectRoot = REPOSITORY_ROOT): Promise<string[]> {
  await assertSafeDemoStorage(projectRoot);
  const ids = await demoIds(projectRoot);
  if (!ids.length) return [];
  const lock = await acquireRunProcessLock(projectRoot, "Resetting local recovery demos", { recoverStale: false });
  try {
    // Validate every target before removing anything; never trust stored deletion paths.
    for (const id of ids) {
      await assertSafeDemoStorage(projectRoot, id);
      const { state, manifest } = await readRecoveryRun(projectRoot, id);
      const demo = recoveryDemoPaths(projectRoot, id);
      if (
        !state.isDemo || !manifest.isDemo ||
        path.resolve(projectRoot, state.sourceWorkbookPath) !== demo.sourceWorkbookPath ||
        path.resolve(projectRoot, state.executedWorkbookPath) !== demo.executedWorkbookPath
      ) {
        throw new Error("Refusing to remove data that is not a verified recovery demo");
      }
    }
    for (const id of ids) {
      const { checkpoint } = await openRecoveryCheckpoint(projectRoot, id);
      await checkpoint.update((state) => { state.status = "ABANDONED"; });
      await rm(recoveryDemoPaths(projectRoot, id).directory, { recursive: true });
      await rm(recoveryPaths(projectRoot, id).runDirectory!, { recursive: true });
    }
    return ids;
  } finally {
    await lock.release();
  }
}
