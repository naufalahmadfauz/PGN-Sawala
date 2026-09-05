import assert from "node:assert/strict";
import {
  link,
  lstat,
  mkdir,
  mkdtemp,
  readFile,
  readdir,
  readlink,
  rename,
  rm,
  symlink,
  writeFile,
} from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import test, { type TestContext } from "node:test";
import { format } from "node:util";
import { chromium } from "playwright";
import { loadConfig, type AppConfig } from "../src/config";
import {
  GoogleDriveEvidencePublisher,
  type EvidenceDrivePublisher,
} from "../src/evidence/google-drive";
import {
  isScenarioComplete,
  isScenarioPartiallyComplete,
  loadPgnWorkbook,
} from "../src/excel/pgn-workbook-loader";
import { TRANSCRIPT_SHEET_NAME } from "../src/excel/pgn-types";
import { getRetestRunMetadata } from "../src/excel/retest-workbook";
import { acquireWorkbookLock } from "../src/excel/workbook-lock";
import { runControlPanel, type OperatorActions } from "../src/operator/control-panel";
import { inspectPgnExecution } from "../src/operator/pgn-preflight";
import type { OperatorUi } from "../src/operator/ui";
import { runPgnWorkbook } from "../src/pgn-runner";
import { recoveryDemoPaths } from "../src/recovery/demo-safety";
import {
  createRecoveryDemo,
  resetRecoveryDemos,
  type RecoveryDemoOptions,
} from "../src/recovery/recovery-demo";
import {
  abandonRecoveryRun,
  formatRecoveryDiscovery,
  formatRecoveryValidation,
  reconcileRecoveryArtifacts,
  repairRecoveryProgress,
  selectRecoveryScenarios,
  skipRecoveryScenario,
  validateRecoveryRun,
} from "../src/recovery/recovery-service";
import {
  acquireRunProcessLock,
  discoverRecoveryRun,
  hashFile,
  initializeRecoveryCheckpoint,
  inspectRunProcessLock,
  openRecoveryCheckpoint,
  readRecoveryRun,
  recoveryPaths,
  type RecoveryRunState,
} from "../src/recovery/run-state";
import { WhatsAppClient } from "../src/whatsapp/client";
import { parseRecoveryDemoArgs } from "./recovery-demo";

const FAKE_SECRET = "OFFLINE-RECOVERY-SECRET-DO-NOT-PRINT";
const FAKE_WEBHOOK = `https://discord.com/api/webhooks/123456789012345678/${FAKE_SECRET}`;
const DEMO_IDS = Array.from({ length: 10 }, (_, index) =>
  `DEMO-KB-${String(index + 1).padStart(3, "0")}`,
);
const EXECUTION_BLOCKED = /Recovery demo detected.*Real WhatsApp execution is disabled/s;

function assertNoSecrets(output: string): void {
  assert.equal(output.includes(FAKE_SECRET), false, "Output exposed a sentinel secret");
  assert.doesNotMatch(output, /https?:\/\/[^\s]*\/api\/webhooks\//);
}

async function snapshotTree(root: string): Promise<Map<string, string>> {
  const snapshot = new Map<string, string>();
  async function visit(relative: string): Promise<void> {
    const target = path.join(root, relative);
    const info = await lstat(target);
    if (info.isSymbolicLink()) {
      snapshot.set(relative, `symlink:${await readlink(target)}`);
    } else if (info.isDirectory()) {
      snapshot.set(relative, "directory");
      for (const name of (await readdir(target)).sort()) {
        await visit(path.join(relative, name));
      }
    } else {
      assert(info.isFile(), `Unexpected fixture entry: ${relative}`);
      snapshot.set(relative, `${info.ino}:${info.nlink}:${(await readFile(target)).toString("base64")}`);
    }
  }
  await visit("");
  return snapshot;
}

function excluding(
  snapshot: Map<string, string>,
  root: string,
  targets: string[],
): Map<string, string> {
  const relative = targets.map((target) => path.relative(root, target));
  return new Map([...snapshot].filter(([entry]) =>
    !relative.some((target) => entry === target || entry.startsWith(`${target}${path.sep}`)),
  ));
}

async function createProject(context: TestContext) {
  const root = await mkdtemp(path.join(tmpdir(), "pgn-recovery-demo-"));
  context.after(() => rm(root, { recursive: true, force: true }));
  const calls: string[] = [];
  const output: string[] = [];
  const forbid = (boundary: string) => async (): Promise<never> => {
    calls.push(boundary);
    throw new Error(`Offline test blocked ${boundary}`);
  };
  context.mock.method(WhatsAppClient.prototype, "open", forbid("WhatsAppClient.open"));
  context.mock.method(chromium, "launchPersistentContext", forbid("chromium.launchPersistentContext"));
  context.mock.method(globalThis, "fetch", forbid("fetch"));
  for (const method of ["validateParentFolder", "validateRunFolder", "ensureRunFolder", "uploadPng"] as const) {
    context.mock.method(GoogleDriveEvidencePublisher.prototype, method, forbid(`Drive.${method}`));
  }
  for (const method of ["log", "warn", "error"] as const) {
    context.mock.method(console, method, (...args: unknown[]) => {
      output.push(format(...args));
    });
  }
  context.after(() => {
    assert.deepEqual(calls, [], "No live boundary or external prerequisite may be called");
    assertNoSecrets(output.join("\n"));
  });
  const config = loadConfig({
    repositoryRoot: root,
    environment: {
      WHATSAPP_HEADLESS: "true",
      GOOGLE_DRIVE_EVIDENCE_ENABLED: "true",
      GOOGLE_DRIVE_EVIDENCE_PARENT_FOLDER: "offline-parent-folder",
      // Invalid JSON also prevents credential resolution from contacting Google on regression.
      GOOGLE_SERVICE_ACCOUNT_JSON: FAKE_SECRET,
      DISCORD_NOTIFICATIONS_ENABLED: "true",
      DISCORD_WEBHOOK_URL: FAKE_WEBHOOK,
    },
  });
  for (const filePath of [
    config.environmentFilePath,
    path.join(root, ".secrets", "service-account.json"),
    path.join(config.profileDir, "Default", "Cookies"),
    config.pgnSourceWorkbookPath,
    config.pgnExecutedWorkbookPath,
    path.join(config.reportArchiveDir, "real-report.xlsx"),
    path.join(config.evidenceDir, "real-evidence.png"),
    path.join(config.debugDir, "real-debug.json"),
  ]) {
    await mkdir(path.dirname(filePath), { recursive: true });
    await writeFile(filePath, `# Real file sentinel: ${path.relative(root, filePath)}\n${FAKE_SECRET}\n`);
  }
  return { root, config, output, forbid };
}

async function createFixture(context: TestContext, options?: RecoveryDemoOptions) {
  const project = await createProject(context);
  const protectedBefore = await snapshotTree(project.root);
  const recovered = await createRecoveryDemo(project.root, options);
  const after = await snapshotTree(project.root);
  for (const [entry, bytes] of protectedBefore) {
    assert.equal(after.get(entry), bytes, `Demo creation changed ${entry}`);
  }
  return {
    ...project,
    ...recovered,
    demo: recoveryDemoPaths(project.root, recovered.state.runId),
    metadata: recoveryPaths(project.root, recovered.state.runId),
  };
}

async function createRealCheckpoint(
  config: AppConfig,
  runId: string,
  status: "INTERRUPTED" | "COMPLETED" | "ABANDONED",
) {
  const { checkpoint } = await initializeRecoveryCheckpoint({
    projectRoot: config.projectRoot,
    runId,
    mode: "full",
    sourceWorkbookPath: config.pgnSourceWorkbookPath,
    executedWorkbookPath: config.pgnExecutedWorkbookPath,
    sourceWorkbookHash: await hashFile(config.pgnSourceWorkbookPath),
    scenarios: [{
      testCaseId: "REAL-KB-001",
      sheetKind: "kb",
      sheetName: "Test Case Knowledge Base",
      sourceRowNumber: 2,
      category: "Real checkpoint sentinel",
      rawStatus: "",
      turns: [{
        sheetName: "Test Case Knowledge Base",
        rowNumber: 2,
        turnNumber: 1,
        userInput: "Preserve this real selection",
      }],
    }],
  });
  await checkpoint.update((state) => {
    state.status = status;
    state.workbookProgress = "Real checkpoint bytes must be preserved";
  });
  return checkpoint;
}

test("default demo is discoverable with three completed scenarios and a genuine partial Turn 1 transcript", async (context) => {
  const fixture = await createFixture(context);
  const { root, state, manifest, demo, metadata } = fixture;
  const discovery = await discoverRecoveryRun(root);
  assert.equal(discovery.kind, "recoverable");
  assert(discovery.kind === "recoverable");
  assert.deepEqual(discovery.state, state);
  assert.deepEqual(discovery.manifest, manifest);
  assert.deepEqual(discovery.lock, { status: "unlocked" });
  assert.equal(state.isDemo, true);
  assert.equal(manifest.isDemo, true);
  assert.match(state.runId, /^DEMO-RECOVERY-/);
  assert.equal(manifest.runId, state.runId);
  assert.equal(state.mode, "full");
  assert.equal(state.status, "INTERRUPTED");
  assert.deepEqual(state.selectedScenarioIds, DEMO_IDS);
  assert.equal(state.totalScenarios, 10);
  assert.deepEqual(state.completedScenarioIds, DEMO_IDS.slice(0, 3));
  assert.deepEqual(state.skippedScenarioIds, []);
  assert.equal(state.lastCompletedScenarioId, "DEMO-KB-003");
  assert.equal(state.activeScenarioId, "DEMO-KB-004");
  assert.equal(state.activeScenarioAttempt, 1);
  assert.equal(state.finalCleanupComplete, false);
  assert.equal(state.driveRunFolderId, undefined);
  assert.deepEqual(manifest.scenarios.map((scenario) => scenario.testCaseId), DEMO_IDS);
  assert.equal(manifest.scenarios[3].turnCount, 2);
  assert.equal(manifest.sourceWorkbookHash, await hashFile(demo.sourceWorkbookPath));
  assert.equal(metadata.runDirectory, path.join(root, ".runtime", "pgn", "runs", state.runId));
  assert.equal(demo.directory, path.join(root, ".runtime", "pgn", "demos", state.runId));
  assert.equal(path.resolve(root, state.sourceWorkbookPath), demo.sourceWorkbookPath);
  assert.equal(path.resolve(root, state.executedWorkbookPath), demo.executedWorkbookPath);
  assert.deepEqual((await readdir(demo.directory)).sort(), ["executed.xlsx", "source.xlsx"]);
  assert.deepEqual((await readdir(metadata.runDirectory!)).sort(), ["manifest.json", "state.json"]);
  await assert.rejects(lstat(metadata.lock), { code: "ENOENT" });

  const { workbook, parsed } = await loadPgnWorkbook(demo.executedWorkbookPath);
  const active = parsed.scenarios.find((scenario) => scenario.testCaseId === state.activeScenarioId)!;
  assert.deepEqual(active.turns.map((turn) => turn.turnNumber), [1, 2]);
  assert.equal(isScenarioComplete(workbook, active), false);
  assert.equal(isScenarioPartiallyComplete(workbook, active), true);
  const transcript = workbook.getWorksheet(TRANSCRIPT_SHEET_NAME)!;
  const activeRows = transcript.getRows(2, transcript.rowCount - 1)!
    .filter((row) => row.getCell(2).text === active.testCaseId);
  assert.deepEqual(activeRows.filter((row) => ["USER", "BOT"].includes(row.getCell(6).text))
    .map((row) => [row.getCell(5).value, row.getCell(6).text, row.getCell(11).text]), [
    [1, "USER", "CAPTURED"],
    [1, "BOT", "CAPTURED"],
  ]);
  assert(!activeRows.some((row) => row.getCell(11).text === "SCENARIO_ATTEMPT_COMPLETED"));
  const reconciliation = reconcileRecoveryArtifacts(workbook, parsed.scenarios, state, manifest);
  assert.deepEqual(reconciliation.checkpointCompletedIds, DEMO_IDS.slice(0, 3));
  assert.deepEqual(reconciliation.workbookCompletedIds, DEMO_IDS.slice(0, 3));
  assert.deepEqual(reconciliation.transcriptCompletedIds, DEMO_IDS.slice(0, 3));
  assert.deepEqual(reconciliation.safeCompletedIds, DEMO_IDS.slice(0, 3));
  assert.deepEqual(reconciliation.mismatchedScenarioIds, []);
  assert.deepEqual(reconciliation.evidenceCapturedScenarioIds, []);
  assert.deepEqual(reconciliation.evidenceUploadedScenarioIds, []);
  assert.equal(reconciliation.nextScenarioId, "DEMO-KB-004");
  assert.equal(reconciliation.restartInterruptedScenarioFromTurnOne, true);
  const remaining = selectRecoveryScenarios(parsed.scenarios, state);
  assert.deepEqual(remaining.map((scenario) => scenario.testCaseId), DEMO_IDS.slice(3));
  assert.equal(remaining[0].turns[0].turnNumber, 1);
  const output = formatRecoveryDiscovery(discovery);
  assert.match(output, /Recovery type: DEMO/);
  assert.match(output, /Interrupted during Turn 2 of 2/);
  assert.match(output, /Recommended: Restart DEMO-KB-004 from Turn 1/);
  assert.match(output, /Process lock: no active process/);
  assertNoSecrets(output);
});

for (const mode of ["full", "retest"] as const) {
  test(`${mode} demo validation is ready, byte-read-only, and skips injected and default external checks`, async (context) => {
    const fixture = await createFixture(context, { mode });
    const { root, config, state, demo, forbid } = fixture;
    await rm(config.profileDir, { recursive: true });
    assert.equal(config.target, undefined);
    assert.equal(config.googleDriveEvidenceEnabled, true);
    assert.equal(config.discordNotificationsEnabled, true);
    const configBefore = structuredClone(config);
    const before = await snapshotTree(root);
    const publisher: EvidenceDrivePublisher = {
      parentFolderId: "offline-parent-folder",
      folderPrefix: config.googleDriveEvidenceFolderPrefix,
      retestFolderPrefix: config.googleDriveRetestFolderPrefix,
      validateParentFolder: forbid("injected Drive parent check"),
      validateRunFolder: forbid("injected Drive run check"),
      ensureRunFolder: forbid("injected Drive folder creation"),
      uploadPng: forbid("injected Drive upload"),
    };
    const inspectedPaths: string[] = [];
    const injected = await validateRecoveryRun(config, state.runId, {
      checkDriveAccess: true,
      drivePublisher: publisher,
      profileEntries: forbid("profile inspection"),
      fileExists: async (filePath) => {
        inspectedPaths.push(filePath);
        assert([demo.sourceWorkbookPath, demo.executedWorkbookPath].includes(filePath));
        return true;
      },
    });
    assert.deepEqual(inspectedPaths.sort(), [demo.sourceWorkbookPath, demo.executedWorkbookPath].sort());
    const defaults = await validateRecoveryRun(config);
    for (const validation of [injected, defaults]) {
      assert.equal(validation.ready, true, JSON.stringify(validation.checks));
      assert.equal(validation.runId, state.runId);
      assert.equal(validation.state.isDemo, true);
      assert.equal(validation.manifest.isDemo, true);
      assert.equal(validation.sourceDrift, "unchanged");
      assert.equal(validation.reconciliation?.restartInterruptedScenarioFromTurnOne, true);
      assert.equal(validation.reconciliation?.nextScenarioId, state.activeScenarioId);
      for (const id of ["drive", "discord", "whatsapp-profile"]) {
        const check = validation.checks.find((item) => item.id === id)!;
        assert.equal(check.status, "info");
        assert.match(check.detail, /demo mode/);
      }
      const output = formatRecoveryValidation(validation);
      assert.match(output, /Resume readiness: READY \(preview only\)/);
      assert.match(output, /Real WhatsApp execution is disabled/);
      assert.match(output, /Google Drive: skipped in demo mode/);
      assert.match(output, /Discord notifications: suppressed in demo mode/);
      assertNoSecrets(output);
    }
    assert.deepEqual(config, configBefore);
    assert.deepEqual(await snapshotTree(root), before);
  });

  for (const entrypoint of ["runner", "preflight"] as const) {
    test(`${mode} demo ${entrypoint} rejects before workbook reads, locks, notifications, or browser startup`, async (context) => {
      const fixture = await createFixture(context, { mode });
      const { root, config, state } = fixture;
      const guardedConfig: AppConfig = { ...config, googleDriveEvidenceEnabled: false };
      const reads: string[] = [];
      for (const field of ["pgnSourceWorkbookPath", "pgnExecutedWorkbookPath", "profileDir"] as const) {
        Object.defineProperty(guardedConfig, field, {
          enumerable: true,
          get() {
            reads.push(field);
            throw new Error(`Demo execution guard must precede ${field} access`);
          },
        });
      }
      // Test both an unlocked workspace and ownership belonging to a real process.
      for (const locked of [false, true]) {
        const runLock = locked ? await acquireRunProcessLock(root, "Real owner sentinel") : undefined;
        const releaseWorkbook = locked
          ? await acquireWorkbookLock(config.pgnExecutedWorkbookPath, "Real workbook owner sentinel")
          : undefined;
        try {
          const before = await snapshotTree(root);
          for (const requestedMode of ["full", "retest"] as const) {
            const args = ["--resume", state.runId, "--accept-source-drift"];
            await assert.rejects(
              entrypoint === "runner"
                ? runPgnWorkbook(args, requestedMode, guardedConfig)
                : inspectPgnExecution(args, requestedMode, guardedConfig),
              (error: unknown) => {
                assert(error instanceof Error);
                assert.match(error.message, EXECUTION_BLOCKED);
                assertNoSecrets(error.message);
                return true;
              },
            );
            assert.deepEqual(await snapshotTree(root), before);
          }
        } finally {
          await releaseWorkbook?.();
          await runLock?.release();
        }
      }
      assert.deepEqual(reads, []);
    });
  }
}

test("a reset demo cannot fall back to legacy retest execution", async (context) => {
  const { root, config, state } = await createFixture(context, { mode: "retest" });
  await resetRecoveryDemos(root);
  const before = await snapshotTree(root);
  const args = ["--resume", state.runId];
  await assert.rejects(runPgnWorkbook(args, "retest", config), EXECUTION_BLOCKED);
  await assert.rejects(inspectPgnExecution(args, "retest", config), EXECUTION_BLOCKED);
  assert.deepEqual(await snapshotTree(root), before);
});

test("real skip and abandon services change only demo progress and append audit events", async (context) => {
  const { root, config, state, manifest, demo, metadata } = await createFixture(context);
  const before = await snapshotTree(root);
  const original = await loadPgnWorkbook(demo.executedWorkbookPath);
  const originalSheets = original.workbook.worksheets
    .filter((sheet) => sheet.name !== TRANSCRIPT_SHEET_NAME)
    .map((sheet) => [sheet.name, sheet.getSheetValues()]);
  const originalTranscript = original.workbook.getWorksheet(TRANSCRIPT_SHEET_NAME)!.getSheetValues();
  const skipped = await skipRecoveryScenario(config, state.runId);
  assert.equal(skipped.scenarioId, "DEMO-KB-004");
  assert.equal(skipped.warning, undefined);
  const afterSkip = await readRecoveryRun(root, state.runId);
  assert.equal(afterSkip.state.isDemo, true);
  assert.deepEqual(afterSkip.manifest, manifest);
  assert.equal(afterSkip.state.status, "RECOVERABLE");
  assert.deepEqual(afterSkip.state.completedScenarioIds, DEMO_IDS.slice(0, 3));
  assert.equal(afterSkip.state.lastCompletedScenarioId, "DEMO-KB-003");
  assert.deepEqual(afterSkip.state.skippedScenarioIds, ["DEMO-KB-004"]);
  assert.equal(afterSkip.state.activeScenarioId, undefined);
  assert.deepEqual(afterSkip.state.scenarioAttempts.slice(0, 3), state.scenarioAttempts.slice(0, 3));
  assert.equal(afterSkip.state.scenarioAttempts[3].status, "SKIPPED_BY_OPERATOR");
  const validation = await validateRecoveryRun(config, state.runId);
  assert.equal(validation.ready, true);
  assert.equal(validation.reconciliation?.nextScenarioId, "DEMO-KB-005");
  assert.deepEqual(validation.reconciliation?.safeCompletedIds, DEMO_IDS.slice(0, 3));
  assert.equal((await discoverRecoveryRun(root)).kind, "recoverable");

  const abandoned = await abandonRecoveryRun(config, state.runId);
  assert.equal(abandoned.warning, undefined);
  const afterAbandon = await readRecoveryRun(root, state.runId);
  assert.equal(afterAbandon.state.status, "ABANDONED");
  assert.equal(afterAbandon.state.isDemo, true);
  assert.equal(afterAbandon.manifest.isDemo, true);
  assert.deepEqual(afterAbandon.state.completedScenarioIds, DEMO_IDS.slice(0, 3));
  assert.deepEqual(afterAbandon.state.skippedScenarioIds, ["DEMO-KB-004"]);
  assert.equal((await discoverRecoveryRun(root)).kind, "none");
  assert.equal((await inspectRunProcessLock(root)).status, "unlocked");
  await assert.rejects(lstat(metadata.active), { code: "ENOENT" });
  const executed = await loadPgnWorkbook(demo.executedWorkbookPath);
  assert.deepEqual(executed.workbook.worksheets
    .filter((sheet) => sheet.name !== TRANSCRIPT_SHEET_NAME)
    .map((sheet) => [sheet.name, sheet.getSheetValues()]), originalSheets);
  const transcript = executed.workbook.getWorksheet(TRANSCRIPT_SHEET_NAME)!;
  assert.deepEqual(transcript.getSheetValues().slice(0, originalTranscript.length), originalTranscript);
  assert.deepEqual(transcript.getRows(transcript.rowCount - 1, 2)!
    .map((row) => [row.getCell(1).text, row.getCell(11).text]), [
    [state.runId, "SCENARIO_SKIPPED_BY_OPERATOR"],
    [state.runId, "RUN_ABANDONED"],
  ]);
  const allowed = [metadata.state!, metadata.active, demo.executedWorkbookPath];
  assert.deepEqual(excluding(await snapshotTree(root), root, allowed), excluding(before, root, allowed));
});

test("retest demo retains its original four-case selection and partial sixth case", async (context) => {
  const { config, state, manifest, demo } = await createFixture(context, { mode: "retest" });
  const selected = ["DEMO-KB-002", "DEMO-KB-004", "DEMO-KB-006", "DEMO-KB-008"];
  assert.equal(state.mode, "retest");
  assert.deepEqual(state.selectedScenarioIds, selected);
  assert.deepEqual(state.completedScenarioIds, selected.slice(0, 2));
  assert.equal(state.lastCompletedScenarioId, "DEMO-KB-004");
  assert.equal(state.activeScenarioId, "DEMO-KB-006");
  assert.equal(state.totalScenarios, 4);
  assert.deepEqual(manifest.scenarios.map((scenario) => scenario.testCaseId), selected);
  assert.equal(manifest.scenarios[2].turnCount, 2);
  const { workbook, parsed } = await loadPgnWorkbook(demo.executedWorkbookPath);
  const retest = getRetestRunMetadata(workbook, state.runId)!;
  assert.deepEqual(retest.selectedIds, selected);
  assert.deepEqual(retest.finishedIds, selected.slice(0, 2));
  assert.equal(retest.state, "IN_PROGRESS");
  assert.deepEqual(selectRecoveryScenarios(parsed.scenarios, state)
    .map((scenario) => scenario.testCaseId), selected.slice(2));
  const transcript = workbook.getWorksheet(TRANSCRIPT_SHEET_NAME)!;
  assert.deepEqual(transcript.getRows(2, transcript.rowCount - 1)!
    .filter((row) => row.getCell(2).text === "DEMO-KB-006" && row.getCell(6).text === "USER")
    .map((row) => row.getCell(5).value), [1]);
  const validation = await validateRecoveryRun(config);
  assert.equal(validation.ready, true);
  assert.deepEqual(validation.reconciliation?.safeCompletedIds, selected.slice(0, 2));
  assert.equal(validation.reconciliation?.nextScenarioId, "DEMO-KB-006");
});

test("source-drift demo produces a formatting-only warning without changing selected inputs", async (context) => {
  const { root, config, state, demo } = await createFixture(context, { sourceDrift: true });
  assert.notEqual(await hashFile(demo.sourceWorkbookPath), state.sourceWorkbookHash);
  const before = await snapshotTree(root);
  const validation = await validateRecoveryRun(config);
  assert.equal(validation.sourceDrift, "formatting-only");
  assert.equal(validation.ready, true);
  assert.equal(validation.checks.find((check) => check.id === "source")?.status, "warn");
  assert.match(formatRecoveryValidation(validation), /selected scenario inputs are unchanged/);
  assert.deepEqual(validation.reconciliation?.mismatchedScenarioIds, []);
  assert.deepEqual(await snapshotTree(root), before);
});

test("mismatch demo blocks readiness on actual 3/3/2 progress and supports conservative shared repair", async (context) => {
  const { root, config, state, demo, metadata } = await createFixture(context, { mismatch: true });
  const before = await snapshotTree(root);
  const validation = await validateRecoveryRun(config);
  assert.equal(validation.ready, false);
  assert.deepEqual(validation.reconciliation?.checkpointCompletedIds, DEMO_IDS.slice(0, 3));
  assert.deepEqual(validation.reconciliation?.transcriptCompletedIds, DEMO_IDS.slice(0, 3));
  assert.deepEqual(validation.reconciliation?.workbookCompletedIds, DEMO_IDS.slice(0, 2));
  assert.deepEqual(validation.reconciliation?.safeCompletedIds, DEMO_IDS.slice(0, 2));
  assert.deepEqual(validation.reconciliation?.mismatchedScenarioIds, ["DEMO-KB-003"]);
  assert.equal(validation.checks.find((check) => check.id === "reconciliation")?.status, "error");
  assert.match(formatRecoveryValidation(validation), /choose a repair strategy before resume/);
  assert.match(formatRecoveryValidation(validation), /Resume readiness: BLOCKED/);
  assert.deepEqual(await snapshotTree(root), before);

  await repairRecoveryProgress(config, state.runId, "rerun");
  const repaired = await validateRecoveryRun(config);
  assert.equal(repaired.ready, true);
  assert.deepEqual(repaired.state.completedScenarioIds, DEMO_IDS.slice(0, 2));
  assert.equal(repaired.state.lastCompletedScenarioId, "DEMO-KB-002");
  assert.equal(repaired.reconciliation?.nextScenarioId, "DEMO-KB-003");
  assert.deepEqual(repaired.reconciliation?.mismatchedScenarioIds, []);
  assert.deepEqual(repaired.reconciliation?.reconciledScenarioIds, ["DEMO-KB-003"]);
  assert.deepEqual(repaired.state.reconciliationDecisions.map(({ scenarioId, strategy }) =>
    ({ scenarioId, strategy })), [{ scenarioId: "DEMO-KB-003", strategy: "rerun" }]);
  const { workbook } = await loadPgnWorkbook(demo.executedWorkbookPath);
  const transcript = workbook.getWorksheet(TRANSCRIPT_SHEET_NAME)!;
  const audit = transcript.getRow(transcript.rowCount);
  assert.equal(audit.getCell(11).text, "RECOVERY_RECONCILED");
  assert.match(audit.getCell(7).text, /re-run every mismatched scenario \(safest\)/);
  const allowed = [metadata.state!, metadata.active, demo.executedWorkbookPath];
  assert.deepEqual(excluding(await snapshotTree(root), root, allowed), excluding(before, root, allowed));
});

test("reset removes only verified demo fixtures and history, preserving real terminal and active checkpoints", async (context) => {
  const { root, config } = await createProject(context);
  await createRealCheckpoint(config, "REAL-COMPLETED", "COMPLETED");
  await createRealCheckpoint(config, "REAL-ABANDONED", "ABANDONED");
  const { state } = await createRecoveryDemo(root);
  const demo = recoveryDemoPaths(root, state.runId);
  const metadata = recoveryPaths(root, state.runId);
  await createRealCheckpoint(config, "REAL-ACTIVE", "INTERRUPTED");
  const before = await snapshotTree(root);
  assert.deepEqual(await resetRecoveryDemos(root), [state.runId]);
  await assert.rejects(lstat(demo.directory), { code: "ENOENT" });
  await assert.rejects(lstat(metadata.runDirectory!), { code: "ENOENT" });
  assert.deepEqual(await snapshotTree(root), excluding(before, root, [demo.directory, metadata.runDirectory!]));
  const discovery = await discoverRecoveryRun(root);
  assert.equal(discovery.kind, "recoverable");
  assert(discovery.kind === "recoverable");
  assert.equal(discovery.state.runId, "REAL-ACTIVE");
  assert.equal(discovery.state.isDemo, undefined);
  assert.deepEqual(discovery.lock, { status: "unlocked" });
  const after = await snapshotTree(root);
  assert.deepEqual(await resetRecoveryDemos(root), []);
  assert.deepEqual(await snapshotTree(root), after);
});

test("reset without a demo is a no-op even with real recovery and a real process lock", async (context) => {
  const { root, config } = await createProject(context);
  const pristine = await snapshotTree(root);
  assert.deepEqual(await resetRecoveryDemos(root), []);
  assert.deepEqual(await snapshotTree(root), pristine);
  await createRealCheckpoint(config, "REAL-ACTIVE", "INTERRUPTED");
  const lock = await acquireRunProcessLock(root, "Real run must not be touched");
  try {
    const before = await snapshotTree(root);
    assert.deepEqual(await resetRecoveryDemos(root), []);
    assert.deepEqual(await snapshotTree(root), before);
  } finally {
    await lock.release();
  }
});

test("creation refuses a real recoverable checkpoint without changing any bytes", async (context) => {
  const { root, config } = await createProject(context);
  await createRealCheckpoint(config, "REAL-RECOVERABLE", "INTERRUPTED");
  const before = await snapshotTree(root);
  await assert.rejects(createRecoveryDemo(root), /Existing recovery state or process lock/);
  assert.deepEqual(await snapshotTree(root), before);
});

for (const kind of ["active", "unreadable"] as const) {
  test(`creation refuses an existing ${kind} process lock and leaves its owner bytes intact`, async (context) => {
    const { root } = await createProject(context);
    const lock = kind === "active" ? await acquireRunProcessLock(root, "Real owner sentinel") : undefined;
    if (!lock) {
      await mkdir(recoveryPaths(root).lock, { recursive: true });
      await writeFile(path.join(recoveryPaths(root).lock, "owner-invalid.json"), "Real unreadable owner sentinel");
    }
    try {
      const before = await snapshotTree(root);
      await assert.rejects(createRecoveryDemo(root), /Existing recovery state or process lock/);
      assert.deepEqual(await snapshotTree(root), before);
    } finally {
      await lock?.release();
    }
  });
}

test("reset and demo mutations refuse an existing real process lock", async (context) => {
  const { root, config, state } = await createFixture(context);
  const lock = await acquireRunProcessLock(root, "Real owner sentinel");
  try {
    const before = await snapshotTree(root);
    for (const operation of [
      () => resetRecoveryDemos(root),
      () => skipRecoveryScenario(config, state.runId),
      () => abandonRecoveryRun(config, state.runId),
      () => repairRecoveryProgress(config, state.runId, "rerun"),
    ]) {
      await assert.rejects(operation(), /process lock already exists/);
      assert.deepEqual(await snapshotTree(root), before);
    }
  } finally {
    await lock.release();
  }
});

for (const abandoned of [false, true]) {
  test(`${abandoned ? "abandoned" : "existing"} demo cannot be overwritten by another creation`, async (context) => {
    const { root, config, state } = await createFixture(context);
    if (abandoned) await abandonRecoveryRun(config, state.runId);
    const before = await snapshotTree(root);
    await assert.rejects(createRecoveryDemo(root, { mode: "retest", sourceDrift: true, mismatch: true }), /recovery demo already exists/);
    assert.deepEqual(await snapshotTree(root), before);
    assert.deepEqual(await resetRecoveryDemos(root), [state.runId]);
    assert.equal((await discoverRecoveryRun(root)).kind, "none");
    const replacement = await createRecoveryDemo(root);
    assert.notEqual(replacement.state.runId, state.runId);
  });
}

test("checkpoint updates cannot strip the demo marker or move into the real namespace", async (context) => {
  const { root, state } = await createFixture(context);
  const { checkpoint } = await openRecoveryCheckpoint(root, state.runId);
  const before = await snapshotTree(root);
  const changes: Array<(draft: RecoveryRunState) => void> = [
    (draft) => { delete draft.isDemo; },
    (draft) => { Reflect.set(draft, "isDemo", false); },
    (draft) => { draft.runId = "REAL-RECOVERY"; },
    (draft) => { draft.runId = "REAL-RECOVERY"; delete draft.isDemo; },
    (draft) => { draft.runId = "DEMO-RECOVERY-../REAL-RECOVERY"; },
  ];
  for (const change of changes) {
    await assert.rejects(checkpoint.update(change), /demo identity|cannot become real|Unsafe recovery Run ID/);
    assert.deepEqual(checkpoint.snapshot(), state);
    assert.deepEqual(await snapshotTree(root), before);
  }
});

for (const target of ["state", "manifest", "both"] as const) {
  for (const marker of [undefined, false] as const) {
    test(`tampered ${target} with isDemo ${marker === undefined ? "removed" : "false"} is rejected without cleanup or mutation`, async (context) => {
      const { root, config, state, metadata } = await createFixture(context);
      for (const filePath of target === "both" ? [metadata.state!, metadata.manifest!] : [metadata[target]!]) {
        const value = JSON.parse(await readFile(filePath, "utf8")) as Record<string, unknown>;
        if (marker === undefined) delete value.isDemo;
        else value.isDemo = marker;
        await writeFile(filePath, JSON.stringify(value));
      }
      const before = await snapshotTree(root);
      assert.equal((await discoverRecoveryRun(root)).kind, "unreadable");
      for (const operation of [
        () => readRecoveryRun(root, state.runId),
        () => validateRecoveryRun(config, state.runId),
        () => skipRecoveryScenario(config, state.runId),
        () => abandonRecoveryRun(config, state.runId),
        () => repairRecoveryProgress(config, state.runId, "rerun"),
        () => resetRecoveryDemos(root),
        () => runPgnWorkbook(["--resume", state.runId], "full", { ...config, googleDriveEvidenceEnabled: false }),
        () => inspectPgnExecution(["--resume", state.runId], "retest", config),
      ]) {
        await assert.rejects(operation(), /demo identity/);
        assert.deepEqual(await snapshotTree(root), before);
      }
    });
  }
}

for (const field of ["sourceWorkbookPath", "executedWorkbookPath"] as const) {
  for (const form of ["relative", "absolute", "traversal"] as const) {
    test(`tampered demo ${field} using a ${form} real-file path is rejected and preserved`, async (context) => {
      const { root, config, state, metadata } = await createFixture(context);
      const realPath = field === "sourceWorkbookPath" ? config.pgnSourceWorkbookPath : config.pgnExecutedWorkbookPath;
      const relative = path.relative(root, realPath);
      const tampered = {
        ...state,
        [field]: form === "absolute" ? realPath : form === "relative" ? relative
          : `.runtime/pgn/demos/${state.runId}/../../../../${relative}`,
      };
      await writeFile(metadata.state!, JSON.stringify(tampered));
      const before = await snapshotTree(root);
      for (const operation of [
        () => validateRecoveryRun(config, state.runId),
        () => skipRecoveryScenario(config, state.runId),
        () => abandonRecoveryRun(config, state.runId),
        () => repairRecoveryProgress(config, state.runId, "rerun"),
        () => resetRecoveryDemos(root),
      ]) {
        await assert.rejects(operation(), /dedicated demo directory|not a verified recovery demo/);
        assert.deepEqual(await snapshotTree(root), before);
      }
    });
  }
}

test("demo storage rejects real namespaces, traversal IDs, and unverified cleanup entries", async (context) => {
  const { root } = await createProject(context);
  const before = await snapshotTree(root);
  for (const id of ["REAL-RUN", "../REAL-RUN", "DEMO-RECOVERY-../REAL-RUN", "DEMO-RECOVERY-..\\REAL-RUN", path.join(root, "REAL-RUN")]) {
    assert.throws(() => recoveryDemoPaths(root, id), /explicitly identified recovery demos|Unsafe recovery Run ID/);
  }
  assert.deepEqual(await snapshotTree(root), before);
  const unverified = path.join(recoveryPaths(root).root, "demos", "REAL-RUN");
  await mkdir(unverified, { recursive: true });
  await writeFile(path.join(unverified, "state.json"), "Preserve unverified real data");
  const withUnverified = await snapshotTree(root);
  await assert.rejects(createRecoveryDemo(root), /Unexpected demo storage entries/);
  await assert.rejects(resetRecoveryDemos(root), /Unexpected demo storage entries/);
  assert.deepEqual(await snapshotTree(root), withUnverified);
});

for (const relative of [".runtime", ".runtime/pgn", ".runtime/pgn/runs", ".runtime/pgn/demos"]) {
  test(`creation and reset reject a symlinked ${relative} storage ancestor`, async (context) => {
    const { root } = await createProject(context);
    const target = path.join(root, relative);
    await mkdir(path.dirname(target), { recursive: true });
    await symlink(path.join(root, "data"), target, "dir");
    const before = await snapshotTree(root);
    await assert.rejects(createRecoveryDemo(root), /symbolic or hard links/);
    await assert.rejects(resetRecoveryDemos(root), /symbolic or hard links/);
    assert.deepEqual(await snapshotTree(root), before);
  });
}

for (const kind of ["symlink", "hardlink"] as const) {
  const targets = kind === "hardlink"
    ? ["state", "manifest", "active", "source", "executed"] as const
    : ["state", "manifest", "active", "source", "executed", "demo-directory", "run-directory", "workbook-lock", "run-lock"] as const;
  for (const target of targets) {
    test(`${kind} at demo ${target} rejects validation, mutations, and reset without touching real targets`, async (context) => {
      const { root, config, state, demo, metadata } = await createFixture(context);
      const filePath = {
        state: metadata.state!,
        manifest: metadata.manifest!,
        active: metadata.active,
        source: demo.sourceWorkbookPath,
        executed: demo.executedWorkbookPath,
        "demo-directory": demo.directory,
        "run-directory": metadata.runDirectory!,
        "workbook-lock": `${demo.executedWorkbookPath}.lock`,
        "run-lock": metadata.lock,
      }[target];
      if (target === "workbook-lock" || target === "run-lock") {
        await mkdir(filePath);
        await writeFile(path.join(filePath, "owner-sentinel.json"), "Preserve real lock contents");
      }
      // The linked target remains inside this temporary project, never the repository.
      const realTarget = path.join(root, "data", `preserved-${target}`);
      await rename(filePath, realTarget);
      if (kind === "symlink") await symlink(realTarget, filePath);
      else await link(realTarget, filePath);
      const before = await snapshotTree(root);
      for (const operation of [
        () => validateRecoveryRun(config, state.runId),
        () => skipRecoveryScenario(config, state.runId),
        () => abandonRecoveryRun(config, state.runId),
        () => repairRecoveryProgress(config, state.runId, "rerun"),
        () => resetRecoveryDemos(root),
      ]) {
        await assert.rejects(operation(), /symbolic or hard links|Unexpected demo storage entries/);
        assert.deepEqual(await snapshotTree(root), before);
      }
    });
  }
}

test("actual generated fixtures drive menu preview, audited skip, and abandonment without live actions", async (context) => {
  const { root, config, state, demo, output } = await createFixture(context);
  const choices = ["resume", "skip", "abandon", "exit"];
  const unexpected = async (): Promise<never> => { throw new Error("Unexpected operator action"); };
  const record = (message: string) => { output.push(message); };
  const ui: OperatorUi = {
    intro: record, outro: record, cancel: record, info: record, success: record, warn: record,
    error: (message) => { assert.fail(message); },
    note: (message, title) => { output.push(`${title}: ${message}`); },
    select: async <Value extends string>(prompt: { options: Array<{ value: Value }> }) => {
      const choice = choices.shift();
      assert(prompt.options.some((option) => option.value === choice), `Unexpected menu choice ${choice}`);
      return choice as Value;
    },
    confirm: async (prompt) => {
      assert.equal(prompt.initialValue, false);
      assert.match(prompt.message, /Explicitly skip DEMO-KB-004|Abandon recovery/);
      return true;
    },
    text: unexpected,
    secret: unexpected,
    task: async (_message, operation) => operation(),
  };
  const actions = new Proxy({
    inspectRecovery: () => discoverRecoveryRun(root),
    validateRecovery: (runId: string) => validateRecoveryRun(config, runId),
    skipRecoveryScenario: (runId: string) => skipRecoveryScenario(config, runId),
    abandonRecovery: (runId: string) => abandonRecoveryRun(config, runId),
  } as OperatorActions, {
    get: (target, property, receiver) => Reflect.get(target, property, receiver) ?? unexpected,
  });
  await runControlPanel(ui, actions);
  assert.deepEqual(choices, []);
  assert.match(output.join("\n"), /Recoverable run found \[DEMO\]/);
  assert.match(output.join("\n"), /Restart interrupted scenario: DEMO-KB-004 from Turn 1 of 2/);
  assert.match(output.join("\n"), /Next scenario: DEMO-KB-005 from Turn 1 \(preview only\)/);
  assert.equal((await readRecoveryRun(root, state.runId)).state.status, "ABANDONED");
  assert.equal((await discoverRecoveryRun(root)).kind, "none");
  assert((await lstat(demo.executedWorkbookPath)).isFile());
});

test("demo CLI parses supported flags without accepting duplicates, reset mixtures, or secret-bearing values", () => {
  assert.deepEqual(parseRecoveryDemoArgs([]), {});
  assert.deepEqual(parseRecoveryDemoArgs(["--reset"]), { reset: true });
  assert.deepEqual(parseRecoveryDemoArgs(["--mode=full"]), { mode: "full" });
  assert.deepEqual(parseRecoveryDemoArgs(["--mismatch", "--mode=retest", "--source-drift"]), {
    mode: "retest", sourceDrift: true, mismatch: true,
  });
  for (const args of [
    ["--mode=full", "--mode=retest"],
    ["--mode=full", "--mode=full"],
    ["--source-drift", "--source-drift"],
    ["--mismatch", "--mismatch"],
    ["--reset", "--reset"],
    ["--reset", "--mismatch"],
    ["--mismatch", "--reset"],
    ["--mode", "retest"],
    ["--mode=unknown"],
    ["--source-drift=true"],
    ["--mismatch=false"],
    ["--unknown"],
    [FAKE_WEBHOOK],
    [`--mode=${FAKE_SECRET}`],
    ["--mode=full", `--mode=${FAKE_WEBHOOK}`],
    ["--mismatch", `--mismatch=${FAKE_SECRET}`],
  ]) {
    assert.throws(() => parseRecoveryDemoArgs(args), (error: unknown) => {
      assert(error instanceof Error);
      assert.match(error.message, /Usage:|Duplicate recovery demo option/);
      assertNoSecrets(error.message);
      return true;
    });
  }
});
