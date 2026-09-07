import path from "node:path";
import { pathToFileURL } from "node:url";
import type { Workbook } from "exceljs";
import { loadConfig, requireTarget, type AppConfig } from "./config";
import { loadPgnWorkbook } from "./excel/pgn-workbook-loader";
import {
  appendLatestTurnExecution,
  applyScenarioResults,
  appendPostResetDrainTranscript,
  appendRecoveryTranscriptEvent,
  appendSessionResetTranscript,
  openExecutedPgnWorkbook,
  saveExecutedPgnWorkbook,
} from "./excel/pgn-workbook-writer";
import {
  EVIDENCE_MIGRATION_VERSION,
  getEvidenceRunMetadata,
  upsertEvidenceFileMetadata,
  upsertEvidenceRunMetadata,
} from "./excel/evidence-workbook";
import {
  applyRetestStatusTransition,
  ensureRetestWorkbookSchema,
  getRetestRunMetadata,
  snapshotRetestHistory,
  updateRetestHistory,
  upsertRetestRunMetadata,
  type RetestRunMetadata,
} from "./excel/retest-workbook";
import { assertPgnWorkbookValid } from "./excel/pgn-workbook-validator";
import { acquireWorkbookLock } from "./excel/workbook-lock";
import type {
  ExecutedTurn,
  PgnTestScenario,
  PgnTestTurn,
  TechnicalStatus,
} from "./excel/pgn-types";
import type { BotSessionResetAttempt, SentMessage } from "./types";
import {
  createGoogleDriveEvidencePublisher,
  type EvidenceDrivePublisher,
} from "./evidence/google-drive";
import { safeGoogleCredentialError } from "./evidence/google-service-account";
import { evidenceFileName } from "./evidence/evidence-migration";
import {
  createDiscordNotifier,
  registerDiscordInterruptionHandlers,
  type DiscordRunProgressEvent,
} from "./notifications/discord";
import {
  assertResumeOptionsCompatible,
  parseCliOptions,
  type CliOptions,
} from "./pgn-cli";
import { selectScenarios } from "./pgn-selection";
import {
  CONTINUOUS_RECOVERY_WARNING, CONTINUOUS_SESSION_WARNING, readSessionMode,
  sessionModeLabel, shouldResetBeforeScenario,
} from "./session-mode";
import { getRunConfiguration, upsertRunConfiguration } from "./excel/run-configuration";
import {
  createRetestRunId,
  needsFinalRetestCleanup,
  retestDriveFolderName,
} from "./retest/retest-run";
import { selectRetestScenarios } from "./retest/retest-selection";
import {
  formatRecoveryValidation,
  selectRecoveryScenarios,
  validateRecoveryRun,
} from "./recovery/recovery-service";
import {
  acquireRunProcessLock,
  assertRecoveryRunExecutable,
  createRecoveryManifest,
  discoverRecoveryRun,
  hashFile,
  initializeRecoveryCheckpoint,
  openRecoveryCheckpoint,
  readRecoveryRun,
  recoveryManifestMatches,
  type RecoveryCheckpoint,
  type RecoveryRunManifest,
  type RecoveryRunState,
  type RunProcessLock,
} from "./recovery/run-state";
import { WhatsAppClient } from "./whatsapp/client";
import {
  BotSessionResetError,
  resetBotSession,
  waitForPostResetQuiet,
} from "./whatsapp/session-reset";

export type PgnExecutionMode = "full" | "retest";

interface RunEvidenceContext {
  publisher: EvidenceDrivePublisher;
  folderId: string;
}

function createRunId(): string {
  return new Date().toISOString().replace(/[-:.]/g, "");
}

function safeFileName(value: string): string {
  return value.replace(/[^a-zA-Z0-9_-]+/g, "-").slice(0, 80) || "test";
}

function relativeToProject(config: AppConfig, absolutePath: string): string {
  return path.relative(config.projectRoot, absolutePath).replaceAll(path.sep, "/");
}

async function saveFailureEvidence(
  client: WhatsAppClient,
  name: string,
  config: AppConfig,
): Promise<string | undefined> {
  try {
    const debug = await client.saveDebugArtifacts(name);
    return relativeToProject(config, debug.screenshotPath);
  } catch (error) {
    console.error(
      `[Debug] Could not save failure evidence: ${error instanceof Error ? error.message : String(error)}`,
    );
    return undefined;
  }
}

async function executeTurn(
  client: WhatsAppClient,
  config: AppConfig,
  runId: string,
  scenario: PgnTestScenario,
  turn: PgnTestTurn,
  driveEvidence?: RunEvidenceContext,
): Promise<ExecutedTurn> {
  const artifactKey = `${runId}-${safeFileName(scenario.testCaseId)}-turn-${turn.turnNumber}`;
  let sentMessage: SentMessage | undefined;
  try {
    const baseline = await client.captureMessageState();
    sentMessage = await client.sendMessage(turn.userInput, baseline);
    let response;
    try {
      response = await client.waitForBotResponse(
        baseline,
        sentMessage,
        `${scenario.testCaseId} Turn ${turn.turnNumber}`,
      );
    } catch (error) {
      const completedAt = new Date();
      const detail = error instanceof Error ? error.message : String(error);
      return {
        turn,
        technicalStatus: "CHAT_ERROR",
        sentAt: sentMessage.sentAt,
        completedAt,
        botMessages: [],
        combinedResponse: "",
        error: detail,
        evidencePath: await saveFailureEvidence(
          client,
          `chat-error-${artifactKey}`,
          config,
        ),
        evidenceStatus: "EVIDENCE_CAPTURE_ERROR",
      };
    }

    const evidenceAbsolutePath = path.join(
      config.evidenceDir,
      `${artifactKey}.png`,
    );
    let evidencePath: string | undefined;
    let evidenceUrl: string | undefined;
    let evidenceStatus: ExecutedTurn["evidenceStatus"];
    let evidenceDriveFileId: string | undefined;
    let evidenceDriveFileName: string | undefined;
    try {
      await client.captureScreenshot(evidenceAbsolutePath);
      evidencePath = relativeToProject(config, evidenceAbsolutePath);
      evidenceStatus = driveEvidence
        ? "EVIDENCE_PENDING"
        : "EVIDENCE_LOCAL_ONLY";
    } catch (error) {
      evidenceStatus = "EVIDENCE_CAPTURE_ERROR";
      console.error(
        `[Evidence] Screenshot failed: ${error instanceof Error ? error.message : String(error)}`,
      );
    }
    if (driveEvidence && evidencePath) {
      try {
        const uploaded = await driveEvidence.publisher.uploadPng({
          folderId: driveEvidence.folderId,
          localPath: evidenceAbsolutePath,
          fileName: evidenceFileName(
            scenario.testCaseId,
            turn.turnNumber,
          ),
        });
        evidenceUrl = uploaded.webViewLink;
        evidenceDriveFileId = uploaded.id;
        evidenceDriveFileName = uploaded.name;
        evidenceStatus = "EVIDENCE_SYNCED";
        console.log(`[Evidence] Uploaded ${uploaded.name}`);
      } catch (error) {
        evidenceStatus = "EVIDENCE_UPLOAD_ERROR";
        console.error(
          `[Evidence] EVIDENCE_UPLOAD_ERROR: ${safeGoogleCredentialError(error, config.googleServiceAccount?.value)}`,
        );
      }
    }

    const technicalStatus: TechnicalStatus = response.timedOut
      ? "TIMEOUT"
      : "CAPTURED";
    const error = response.timedOut
      ? `TIMEOUT after ${config.responseTimeoutMs} ms`
      : undefined;
    if (response.timedOut) {
      evidencePath ??= await saveFailureEvidence(
        client,
        `response-timeout-${artifactKey}`,
        config,
      );
    }
    return {
      turn,
      technicalStatus,
      sentAt: sentMessage.sentAt,
      completedAt: response.completedAt,
      botMessages: response.messages.map((message, index) => ({
        sequence: index + 1,
        message: message.text,
        timestamp: message.observedAt,
      })),
      combinedResponse: response.combinedResponse,
      firstResponseMs: response.firstResponseMs,
      totalResponseMs: response.totalResponseMs,
      error,
      evidencePath,
      evidenceUrl,
      evidenceStatus,
      evidenceDriveFileId,
      evidenceDriveFileName,
    };
  } catch (error) {
    const completedAt = new Date();
    const detail = error instanceof Error ? error.message : String(error);
    return {
      turn,
      technicalStatus: "SEND_ERROR",
      sentAt: sentMessage?.sentAt,
      completedAt,
      botMessages: [],
      combinedResponse: "",
      error: detail,
      evidencePath: await saveFailureEvidence(
        client,
        `send-error-${artifactKey}`,
        config,
      ),
      evidenceStatus: "EVIDENCE_CAPTURE_ERROR",
    };
  }
}

function makeResetArtifactPathsRelative(
  config: AppConfig,
  attempt: BotSessionResetAttempt,
): void {
  if (attempt.evidencePath) {
    attempt.evidencePath = relativeToProject(config, attempt.evidencePath);
  }
  if (attempt.diagnosticsPath) {
    attempt.diagnosticsPath = relativeToProject(
      config,
      attempt.diagnosticsPath,
    );
  }
}

async function resetAndDrainSession(
  client: WhatsAppClient,
  config: AppConfig,
  runId: string,
  scenario: PgnTestScenario,
  workbook: Workbook,
  finalCleanup = false,
): Promise<void> {
  try {
    const attempt = await resetBotSession(client, {
      command: config.resetCommand,
      confirmation: config.resetConfirmation,
      timeoutMs: config.resetTimeoutMs,
      failureArtifactName: `reset-failure-${safeFileName(scenario.testCaseId)}-${runId}`,
    });
    const drain = await waitForPostResetQuiet(client, {
      baselineMessages:
        attempt.messageStateAtCompletion ?? attempt.responseMessages,
      quietMs: config.postResetQuietMs,
    });
    makeResetArtifactPathsRelative(config, attempt);
    appendSessionResetTranscript(workbook, runId, scenario, attempt);
    appendPostResetDrainTranscript(workbook, runId, scenario, drain);
    await saveExecutedPgnWorkbook(workbook, config.pgnExecutedWorkbookPath);
  } catch (error) {
    if (!(error instanceof BotSessionResetError)) {
      throw error;
    }

    makeResetArtifactPathsRelative(config, error.attempt);
    appendSessionResetTranscript(workbook, runId, scenario, error.attempt);
    await saveExecutedPgnWorkbook(workbook, config.pgnExecutedWorkbookPath);
    if (error.attempt.evidencePath) {
      console.error(
        `[Session] Debug screenshot: ${error.attempt.evidencePath}`,
      );
    }
    if (error.attempt.diagnosticsPath) {
      console.error(
        `[Session] Diagnostics: ${error.attempt.diagnosticsPath}`,
      );
    }
    console.error("[Session] ABORTING TEST RUN");
    console.error(
      finalCleanup
        ? "[Session] Reason: Unable to confirm final PGN bot session cleanup."
        : "[Session] Reason: Unable to confirm clean PGN bot session before next scenario.",
    );
    console.error("[Session] Completed test results have been saved.");
    if (!finalCleanup) {
      console.error("[Session] Remaining scenarios were NOT executed.");
    }
    throw new Error(
      finalCleanup
        ? "Unable to confirm final PGN bot session cleanup."
        : "Unable to confirm clean PGN bot session before next scenario.",
      { cause: error },
    );
  }
}

function uniqueRetestRunId(workbook: Workbook, now: Date): string {
  const base = createRetestRunId(now);
  let runId = base;
  let suffix = 2;
  while (getRetestRunMetadata(workbook, runId)) {
    runId = `${base}-${suffix}`;
    suffix += 1;
  }
  return runId;
}

export async function runPgnWorkbook(
  args = process.argv.slice(2),
  mode: PgnExecutionMode = "full",
  config = loadConfig(),
): Promise<void> {
  const options = parseCliOptions(args);
  assertResumeOptionsCompatible(options);
  let recoveryResume:
    | { state: RecoveryRunState; manifest: RecoveryRunManifest }
    | undefined;
  let recoveryRestart: typeof recoveryResume;
  try {
    const discovered = await discoverRecoveryRun(config.projectRoot);
    const recoveryRunId = options.resumeRunId ?? options.restartRunId;
    if (recoveryRunId) {
      try {
        recoveryResume = await readRecoveryRun(
          config.projectRoot,
          recoveryRunId,
        );
      } catch (error) {
        const missing = (error as NodeJS.ErrnoException).code === "ENOENT";
        if (missing) assertRecoveryRunExecutable({ runId: recoveryRunId });
        const legacyRetest = mode === "retest" && missing && !options.restartRunId;
        if (!legacyRetest || discovered.kind !== "none") throw error;
      }
      if (recoveryResume) {
        assertRecoveryRunExecutable(recoveryResume.state);
        const storedSessionMode = readSessionMode(recoveryResume.state.sessionMode);
        if (options.sessionModeExplicit && options.sessionMode !== storedSessionMode) {
          throw new Error("Session mode conflicts with the stored run; recovery cannot change isolation semantics");
        }
        options.sessionMode = storedSessionMode;
        if (options.resumeRunId && storedSessionMode === "continuous") throw new Error(CONTINUOUS_RECOVERY_WARNING);
        if (options.restartRunId && storedSessionMode !== "continuous") {
          throw new Error("--restart-run is only for interrupted continuous runs; use normal isolated recovery");
        }
        if (recoveryResume.state.mode !== mode) {
          throw new Error(
            `Run ${recoveryResume.state.runId} is a ${recoveryResume.state.mode} run, not a ${mode} run`,
          );
        }
        if (
          recoveryResume.state.status === "COMPLETED" ||
          recoveryResume.state.status === "ABANDONED"
        ) {
          throw new Error(
            `Run ${recoveryResume.state.runId} is ${recoveryResume.state.status} and cannot be resumed`,
          );
        }
        const validation = await validateRecoveryRun(
          config,
          recoveryResume.state.runId,
        );
        if (
          validation.sourceDrift === "formatting-only" &&
          !options.acceptSourceDrift
        ) {
          throw new Error(
            "Source workbook formatting changed. Review it in npm run pgn, or pass --accept-source-drift after explicit review.",
          );
        }
        if (options.restartRunId ? !validation.restartReady : !validation.ready) {
          throw new Error(formatRecoveryValidation(validation));
        }
        if (options.restartRunId) {
          recoveryRestart = recoveryResume;
          recoveryResume = undefined;
        }
      }
    } else if (discovered.kind === "running") {
      throw new Error(
        `Another PGN process is active for Run ${discovered.state.runId}`,
      );
    } else if (discovered.kind === "recoverable") {
      throw new Error(
        `Recoverable PGN run ${discovered.state.runId} exists. Use npm run pgn to resume or abandon it explicitly.`,
      );
    } else if (discovered.kind === "unreadable") {
      throw new Error(
        `Recovery state is unreadable${discovered.runId ? ` for ${discovered.runId}` : ""}: ${discovered.reason}`,
      );
    }

    const purpose = mode === "retest" ? "PGN retest runner" : "PGN test runner";
    const runProcessLock = await acquireRunProcessLock(
      config.projectRoot,
      purpose,
    );
    let releaseWorkbookLock: (() => Promise<void>) | undefined;
    let lockReleaseOperation: Promise<void> | undefined;
    const releaseLocks = (): Promise<void> => {
      lockReleaseOperation ??= (async () => {
        try {
          await releaseWorkbookLock?.();
        } finally {
          await runProcessLock.release();
        }
      })();
      return lockReleaseOperation;
    };
    try {
      if (recoveryRestart && JSON.stringify(await readRecoveryRun(config.projectRoot, recoveryRestart.state.runId)) !== JSON.stringify(recoveryRestart)) {
        throw new Error("Recovery state changed before restart; inspect it again before execution");
      }
      releaseWorkbookLock = await acquireWorkbookLock(
        config.pgnExecutedWorkbookPath,
        purpose,
      );
      await runPgnWorkbookLocked(
        options,
        config,
        mode,
        runProcessLock,
        releaseLocks,
        recoveryResume,
        recoveryRestart,
      );
    } finally {
      await releaseLocks();
    }
  } catch (error) {
    throw new Error(
      safeGoogleCredentialError(error, config.googleServiceAccount?.value),
      { cause: error },
    );
  }
}

async function runPgnWorkbookLocked(
  options: CliOptions,
  config: AppConfig,
  mode: PgnExecutionMode,
  runProcessLock: RunProcessLock,
  releaseLocks: () => Promise<void>,
  recoveryResume?: {
    state: RecoveryRunState;
    manifest: RecoveryRunManifest;
  },
  recoveryRestart?: { state: RecoveryRunState; manifest: RecoveryRunManifest },
): Promise<void> {
  const sessionMode = options.sessionMode;
  const executionContext = { transport: "whatsapp" as const, sessionMode };
  const source = await loadPgnWorkbook(config.pgnSourceWorkbookPath);
  assertPgnWorkbookValid(source.parsed);
  const executed = await openExecutedPgnWorkbook(
    config.pgnSourceWorkbookPath,
    config.pgnExecutedWorkbookPath,
  );
  assertPgnWorkbookValid(executed.parsed);
  for (const issue of executed.parsed.issues.filter((issue) => issue.severity === "WARNING")) {
    console.log(`[Workbook warning] ${issue.sheetName}: ${issue.message}`);
  }

  console.log(
    `[Workbook] ${executed.resumed ? "Resuming" : "Created"}: ${relativeToProject(config, config.pgnExecutedWorkbookPath)}`,
  );
  let selectedScenarios: PgnTestScenario[];
  let runId: string;
  let retestRun: RetestRunMetadata | undefined;
  let skippedByStatusCount = 0;
  let finalCleanupOnly = false;
  let allRunScenarios: PgnTestScenario[];
  if (mode === "retest") {
    if (options.rerunAll || options.rerunIds.size) {
      throw new Error("--rerun is not used in retest mode; use --test instead");
    }
    const resumedRun = options.resumeRunId
      ? getRetestRunMetadata(executed.workbook, options.resumeRunId)
      : undefined;
    if (options.resumeRunId && !resumedRun) {
      throw new Error(`Retest Run was not found: ${options.resumeRunId}`);
    }
    if (resumedRun && options.sessionModeExplicit && options.sessionMode !== readSessionMode(resumedRun.sessionMode)) {
      throw new Error("Session mode conflicts with the stored retest; recovery cannot change isolation semantics");
    }
    if (resumedRun && readSessionMode(resumedRun.sessionMode) === "continuous") throw new Error(CONTINUOUS_RECOVERY_WARNING);
    const recoveryFinished = recoveryResume
      ? new Set([
          ...recoveryResume.state.completedScenarioIds,
          ...recoveryResume.state.skippedScenarioIds,
        ])
      : undefined;
    const retestSelection = selectRetestScenarios(executed.parsed.scenarios, {
      testIds: options.testIds,
      sheet: options.sheet,
      limit: options.limit,
      resumeSelectedIds:
        recoveryRestart?.state.selectedScenarioIds ?? recoveryResume?.state.selectedScenarioIds ?? resumedRun?.selectedIds,
      completedIds: recoveryFinished ?? new Set(resumedRun?.finishedIds ?? []),
    });
    const scopedScenarioCount = options.sheet
      ? executed.parsed.scenarios.filter(
          (scenario) => scenario.sheetKind === options.sheet,
        ).length
      : executed.parsed.scenarios.length;
    const readyInScope = options.sheet
      ? retestSelection.readyBySheet[options.sheet].length
      : retestSelection.readyBySheet.kb.length +
        retestSelection.readyBySheet.negative.length;
    skippedByStatusCount = scopedScenarioCount - readyInScope;
    console.log("PGN Retest Selection");
    console.log(
      `Ready for Re-test: ${retestSelection.readyBySheet.kb.length + retestSelection.readyBySheet.negative.length}`,
    );
    retestSelection.warnings.forEach((warning) =>
      console.log(`[Retest Warning] ${warning}`),
    );
    selectedScenarios = retestSelection.selected;
    if (selectedScenarios.length === 0) {
      if (
        (recoveryResume && !recoveryResume.state.finalCleanupComplete) ||
        (!recoveryResume &&
          needsFinalRetestCleanup(resumedRun, selectedScenarios.length))
      ) {
        finalCleanupOnly = true;
        console.log(
          "All selected scenarios are complete; retrying final session cleanup.",
        );
      } else if (!recoveryResume) {
        console.log("Nothing to execute.");
        return;
      }
    }
    if (!config.googleDriveEvidenceEnabled) {
      throw new Error(
        "Retest mode requires Google Drive evidence. Configure Drive before launching the selected retests.",
      );
    }
    const startedAt = recoveryResume
      ? new Date(recoveryResume.state.startedAt)
      : resumedRun?.startedAt ?? new Date();
    runId =
      recoveryResume?.state.runId ??
      resumedRun?.runId ??
      uniqueRetestRunId(executed.workbook, startedAt);
    retestRun = resumedRun ?? {
      runId,
      startedAt,
      state: "IN_PROGRESS",
      selectedIds: selectedScenarios.map((scenario) => scenario.testCaseId),
      finishedIds: [],
      updatedAt: startedAt,
    };
    ensureRetestWorkbookSchema(executed.workbook);
    upsertRetestRunMetadata(executed.workbook, retestRun);
    const runIds =
      recoveryResume?.state.selectedScenarioIds ?? retestRun.selectedIds;
    const byId = new Map(
      executed.parsed.scenarios.map((scenario) => [scenario.testCaseId, scenario]),
    );
    allRunScenarios = runIds.map((id) => {
      const scenario = byId.get(id);
      if (!scenario) throw new Error(`Retest scenario was not found: ${id}`);
      return scenario;
    });
    console.log(`Retest Run: ${runId}`);
  } else {
    if (recoveryRestart) {
      const byId = new Map(executed.parsed.scenarios.map((scenario) => [scenario.testCaseId, scenario]));
      selectedScenarios = recoveryRestart.state.selectedScenarioIds.map((id) => {
        const scenario = byId.get(id);
        if (!scenario) throw new Error(`Restart scenario was not found: ${id}`);
        return scenario;
      });
      allRunScenarios = selectedScenarios;
      runId = createRunId();
    } else if (recoveryResume) {
      selectedScenarios = selectRecoveryScenarios(
        executed.parsed.scenarios,
        recoveryResume.state,
      );
      const byId = new Map(
        executed.parsed.scenarios.map((scenario) => [scenario.testCaseId, scenario]),
      );
      allRunScenarios = recoveryResume.state.selectedScenarioIds.map((id) => {
        const scenario = byId.get(id);
        if (!scenario) throw new Error(`Recovery scenario was not found: ${id}`);
        return scenario;
      });
      finalCleanupOnly =
        selectedScenarios.length === 0 &&
        !recoveryResume.state.finalCleanupComplete;
      runId = recoveryResume.state.runId;
    } else {
      const selection = selectScenarios(
        executed.parsed.scenarios,
        options,
        executed.workbook,
      );
      selection.skipped.forEach((message) => console.log(`[Skip] ${message}`));
      selectedScenarios = selection.runnable;
      if (selectedScenarios.length === 0) {
        console.log("[Test] No scenarios require execution");
        return;
      }
      allRunScenarios = selectedScenarios;
      runId = createRunId();
    }
  }

  const sourceById = new Map(
    source.parsed.scenarios.map((scenario) => [scenario.testCaseId, scenario]),
  );
  const manifestScenarios = allRunScenarios.map((scenario) => {
    const sourceScenario = sourceById.get(scenario.testCaseId);
    if (!sourceScenario) {
      throw new Error(`Scenario is missing from the source workbook: ${scenario.testCaseId}`);
    }
    return sourceScenario;
  });
  const recovering = recoveryResume ?? recoveryRestart;
  if (recovering && !recoveryManifestMatches(recovering.manifest, createRecoveryManifest(
    recovering.state.runId, recovering.state.sourceWorkbookHash, manifestScenarios,
    new Date(recovering.manifest.createdAt), executionContext,
  ))) throw new Error("Selected source scenarios changed before recovery execution; no testcase was sent");
  let recoveryCheckpoint: RecoveryCheckpoint;
  if (recoveryResume) {
    const opened = await openRecoveryCheckpoint(config.projectRoot, runId);
    recoveryCheckpoint = opened.checkpoint;
    const resumedAt = new Date();
    await recoveryCheckpoint.update((state) => {
      const activeAttempt = [...state.scenarioAttempts]
        .reverse()
        .find(
          (attempt) =>
            attempt.scenarioId === state.activeScenarioId &&
            attempt.status === "RUNNING",
        );
      if (activeAttempt) {
        activeAttempt.status = "INTERRUPTED";
        activeAttempt.finishedAt = resumedAt.toISOString();
        activeAttempt.reason =
          state.interruptionReason ?? "Previous process ended before scenario completion";
      }
      state.status = "RECOVERABLE";
      state.sessionMode = sessionMode;
      state.transport = executionContext.transport;
      if (selectedScenarios.length > 0) state.finalCleanupComplete = false;
      state.resumedAt = resumedAt.toISOString();
      state.resumeCount += 1;
      state.updatedAt = resumedAt.toISOString();
      state.heartbeatAt = resumedAt.toISOString();
    });
  } else {
    const initialized = await initializeRecoveryCheckpoint({
      projectRoot: config.projectRoot,
      runId,
      mode,
      ...executionContext,
      restartedFromRunId: recoveryRestart?.state.runId,
      sourceWorkbookPath: config.pgnSourceWorkbookPath,
      executedWorkbookPath: config.pgnExecutedWorkbookPath,
      sourceWorkbookHash: await hashFile(config.pgnSourceWorkbookPath),
      scenarios: manifestScenarios,
      startedAt: retestRun?.startedAt,
    });
    recoveryCheckpoint = initialized.checkpoint;
    if (retestRun?.finishedIds.length) {
      await recoveryCheckpoint.update((state) => {
        state.completedScenarioIds = state.selectedScenarioIds.filter((id) =>
          retestRun!.finishedIds.includes(id),
        );
        state.lastCompletedScenarioId = state.completedScenarioIds.at(-1);
        state.metrics.executedScenarios = state.completedScenarioIds.length;
        state.metrics.capturedScenarios = state.completedScenarioIds.length;
      });
    }
  }
  await runProcessLock.heartbeat({ runId, mode });
  const recoveryEventAt = new Date();
  const recoverySnapshot = recoveryCheckpoint.snapshot();
  upsertRunConfiguration(executed.workbook, {
    runId, ...executionContext, sessionResetAttempts: recoverySnapshot.sessionResetAttempts ?? 0,
    restartedFromRunId: recoveryRestart?.state.runId,
  });
  if (recoveryRestart) {
    upsertRunConfiguration(executed.workbook, {
      runId: recoveryRestart.state.runId, ...executionContext,
      sessionResetAttempts: recoveryRestart.state.sessionResetAttempts ?? 0,
      restartedFromRunId: recoveryRestart.state.restartedFromRunId,
    });
    appendRecoveryTranscriptEvent(executed.workbook, {
      runId: recoveryRestart.state.runId, event: "RUN_RESTARTED", timestamp: recoveryEventAt,
      message: `Continuous run restarted from the beginning as ${runId}; all original scenarios will run again in a new clean session.`,
    });
  }
  appendRecoveryTranscriptEvent(executed.workbook, {
    runId,
    event: recoveryResume ? "RUN_RESUMED" : "RUN_PREPARED",
    message: recoveryResume
      ? `Recovery resumed with ${selectedScenarios.length} incomplete scenario(s).${recoverySnapshot.activeScenarioId ? ` Interrupted scenario ${recoverySnapshot.activeScenarioId} will restart from Turn 1.` : ""}${recoverySnapshot.driveRunFolderId ? " Existing Google Drive run folder will be reused." : ""}`
      : `Recovery checkpoint created for ${allRunScenarios.length} selected scenario(s) before external actions.`,
    timestamp: recoveryEventAt,
    scenario: recoverySnapshot.activeScenarioId
      ? allRunScenarios.find(
          (scenario) => scenario.testCaseId === recoverySnapshot.activeScenarioId,
        )
      : undefined,
  });
  await saveExecutedPgnWorkbook(
    executed.workbook,
    config.pgnExecutedWorkbookPath,
  );
  await recoveryCheckpoint.update((state) => {
    state.workbookProgress = recoveryResume
      ? "Recovery resume event saved"
      : "Run preparation event saved";
    state.updatedAt = recoveryEventAt.toISOString();
    state.heartbeatAt = recoveryEventAt.toISOString();
  });
  if (recoveryRestart) {
    const previous = await openRecoveryCheckpoint(config.projectRoot, recoveryRestart.state.runId);
    await previous.checkpoint.update((state) => {
      for (const attempt of state.scenarioAttempts.filter((attempt) => attempt.status === "RUNNING")) {
        attempt.status = "INTERRUPTED";
        attempt.finishedAt = recoveryEventAt.toISOString();
        attempt.reason = `Superseded by full continuous restart ${runId}`;
      }
      state.status = "ABANDONED";
      state.activeScenarioId = undefined;
      state.activeScenarioAttempt = undefined;
      state.activeScenarioStartedAt = undefined;
      state.interruptionReason = `Restarted from the beginning as ${runId}`;
      state.updatedAt = recoveryEventAt.toISOString();
      state.heartbeatAt = recoveryEventAt.toISOString();
    });
  }

  console.log(
    `[Test] ${recoveryResume ? "Remaining" : "Selected"} ${selectedScenarios.length} scenario(s), ${selectedScenarios.reduce((count, scenario) => count + scenario.turns.length, 0)} turn(s)`,
  );
  console.log(`[Test] Transport: WhatsApp; Session Mode: ${sessionModeLabel(sessionMode)}`);
  if (sessionMode === "continuous") console.warn(CONTINUOUS_SESSION_WARNING);
  else console.log(`[Session] Isolation enabled: send "${config.resetCommand}" and require "${config.resetConfirmation}" before every scenario`);

  const initialRecoveryState = recoveryCheckpoint.snapshot();
  const notificationStartedAt = new Date(initialRecoveryState.startedAt);
  let executedCount = initialRecoveryState.metrics.executedScenarios;
  let capturedCount = initialRecoveryState.metrics.capturedScenarios;
  let timeoutCount = initialRecoveryState.metrics.timeouts;
  let errorCount = initialRecoveryState.metrics.technicalErrors;
  let evidenceCapturedCount = initialRecoveryState.metrics.evidenceCaptured;
  let evidenceUploadedCount = initialRecoveryState.metrics.evidenceUploaded;
  let evidenceUploadErrorCount = initialRecoveryState.metrics.evidenceUploadErrors;
  let sessionResetAttempts = initialRecoveryState.sessionResetAttempts ?? 0;
  let awaitingEvaluationCount = 0;
  let currentScenarioId = initialRecoveryState.activeScenarioId;
  let workbookProgress = initialRecoveryState.workbookProgress;
  let failureStage = "initializing the active test run";
  const notifier = createDiscordNotifier(config);
  const throwIfInterrupted = (): void => {
    if (interruptionSignal) {
      throw new Error(`Run interrupted by ${interruptionSignal}`);
    }
  };
  const performSessionReset = async (scenario: PgnTestScenario, finalCleanup = false): Promise<void> => {
    throwIfInterrupted();
    sessionResetAttempts += 1;
    await recoveryCheckpoint.update((state) => {
      state.sessionResetAttempts = sessionResetAttempts;
      state.updatedAt = new Date().toISOString();
    });
    upsertRunConfiguration(executed.workbook, {
      ...getRunConfiguration(executed.workbook, runId)!, sessionResetAttempts,
    });
    throwIfInterrupted();
    if (!activeClient) throw new Error("No active WhatsApp client is available for reset");
    await resetAndDrainSession(activeClient, config, runId, scenario, executed.workbook, finalCleanup);
  };
  const notificationProgress = (updatedAt = new Date()): DiscordRunProgressEvent => ({
    sessionResetAttempts,
    completedScenarios: executedCount,
    totalScenarios: initialRecoveryState.totalScenarios,
    currentScenarioId,
    capturedScenarios: capturedCount,
    timeouts: timeoutCount,
    technicalErrors: errorCount,
    evidenceUploaded: evidenceUploadedCount,
    evidenceUploadErrors: evidenceUploadErrorCount,
    updatedAt,
  });
  let progressTimer: NodeJS.Timeout | undefined;
  let heartbeatTimer: NodeJS.Timeout | undefined;
  let heartbeatOperation: Promise<void> | undefined;
  let periodicProgress: Promise<void> | undefined;
  let activeClient: WhatsAppClient | undefined;
  let interruptionSignal: "SIGINT" | "SIGTERM" | undefined;
  let resolveShutdownSettled!: () => void;
  const shutdownSettled = new Promise<void>((resolve) => {
    resolveShutdownSettled = resolve;
  });
  const notificationOperations = new Set<Promise<void>>();
  const trackNotification = (operation: Promise<void>): Promise<void> => {
    const tracked = operation
      .catch(() => undefined)
      .finally(() => notificationOperations.delete(tracked));
    notificationOperations.add(tracked);
    return tracked;
  };
  const requestProgressNotification = (): void => {
    if (periodicProgress) return;
    periodicProgress = trackNotification(
      notifier.runProgress(notificationProgress()),
    ).finally(() => {
      periodicProgress = undefined;
    });
  };
  const stopProgressTimer = (): void => {
    if (progressTimer) clearInterval(progressTimer);
    progressTimer = undefined;
  };
  const stopHeartbeatTimer = (): void => {
    if (heartbeatTimer) clearInterval(heartbeatTimer);
    heartbeatTimer = undefined;
  };
  const requestHeartbeat = (): void => {
    if (heartbeatOperation) return;
    heartbeatOperation = Promise.all([
      runProcessLock.heartbeat({ runId, mode }),
      recoveryCheckpoint.heartbeat(),
    ])
      .then(() => undefined)
      .catch((error) => {
        console.error(
          `[Recovery] Heartbeat failed: ${error instanceof Error ? error.message : String(error)}`,
        );
      })
      .finally(() => {
        heartbeatOperation = undefined;
      });
  };
  heartbeatTimer = setInterval(requestHeartbeat, 15_000);
  heartbeatTimer.unref();
  const unregisterInterruptionHandlers = registerDiscordInterruptionHandlers({
    notifier,
    progress: () => notificationProgress(),
    onSignal: (signal) => {
      interruptionSignal ??= signal;
      stopProgressTimer();
      stopHeartbeatTimer();
    },
    details: (signal) => ({
      reason: `Process received ${signal}; recovery checkpoint preserved`,
      workbookProgress,
      evidenceProgress: `${evidenceCapturedCount} captured, ${evidenceUploadedCount} uploaded, ${evidenceUploadErrorCount} upload errors`,
    }),
    settle: async () => {
      await shutdownSettled;
      await Promise.all([...notificationOperations]);
      await releaseLocks();
    },
    cleanup: async (signal) => {
      const interruptedAt = new Date();
      await Promise.all([
        recoveryCheckpoint.update((state) => {
          state.status = "INTERRUPTED";
          state.interruptedAt = interruptedAt.toISOString();
          state.interruptionReason = `Process received ${signal}`;
          state.workbookProgress = workbookProgress;
          state.updatedAt = interruptedAt.toISOString();
          state.heartbeatAt = interruptedAt.toISOString();
        }),
        activeClient?.close(),
      ]);
    },
    notificationTimeoutMs: 5_000,
  });

  try {
    await trackNotification(
      recoveryResume
        ? notifier.runResumed({
            ...executionContext,
            runId,
            mode,
            selectedScenarios: initialRecoveryState.totalScenarios,
            startedAt: notificationStartedAt,
            resumedAt: new Date(initialRecoveryState.resumedAt ?? Date.now()),
            completedScenarios: initialRecoveryState.completedScenarioIds.length,
            remainingScenarios: selectedScenarios.length,
            interruptedScenarioId: initialRecoveryState.activeScenarioId,
            reusedDriveFolder: Boolean(initialRecoveryState.driveRunFolderId),
            googleDriveEvidenceEnabled: config.googleDriveEvidenceEnabled,
            workbookPath: config.pgnExecutedWorkbookPath,
          })
        : notifier.runStarted({
            ...executionContext,
            runId,
            mode,
            selectedScenarios: initialRecoveryState.totalScenarios,
            startedAt: notificationStartedAt,
            googleDriveEvidenceEnabled: config.googleDriveEvidenceEnabled,
            workbookPath: config.pgnExecutedWorkbookPath,
          }),
    );
    throwIfInterrupted();
    if (
      config.discordNotificationsEnabled &&
      config.discordNotifyProgress &&
      selectedScenarios.length > 0
    ) {
      const progressPollingIntervalMs = Math.min(
        config.discordProgressMinutes * 60_000,
        10_000,
      );
      progressTimer = setInterval(() => {
        requestProgressNotification();
      }, progressPollingIntervalMs);
      progressTimer.unref();
    }
    failureStage = "preparing Google Drive evidence";
    let driveEvidence: RunEvidenceContext | undefined;
    const storedEvidenceRun = getEvidenceRunMetadata(executed.workbook, runId);
    const checkpointBeforeDrive = recoveryCheckpoint.snapshot();
    let driveFolderId =
      storedEvidenceRun?.folderId ?? checkpointBeforeDrive.driveRunFolderId ?? "";
    let driveFolderUrl =
      storedEvidenceRun?.folderUrl ?? checkpointBeforeDrive.driveRunFolderUrl ?? "";
    if (config.googleDriveEvidenceEnabled) {
      const publisher = createGoogleDriveEvidencePublisher(config);
      const parent = await publisher.validateParentFolder();
      throwIfInterrupted();
      console.log(`[Evidence] Drive parent ready: ${parent.name} (${parent.id})`);
      const folder = await publisher.ensureRunFolder(
        runId,
        driveFolderId || undefined,
        mode === "retest"
          ? retestDriveFolderName(config.googleDriveRetestFolderPrefix, runId)
          : undefined,
      );
      throwIfInterrupted();
      driveFolderId = folder.id;
      driveFolderUrl = folder.webViewLink;
      driveEvidence = { publisher, folderId: folder.id };
      console.log(`[Evidence] Drive run folder ready: ${folder.name}`);
      const folderCheckpointAt = new Date();
      await recoveryCheckpoint.update((state) => {
        state.driveRunFolderId = folder.id;
        state.driveRunFolderUrl = folder.webViewLink;
        state.workbookProgress = "Drive run folder recorded in recovery state";
        state.updatedAt = folderCheckpointAt.toISOString();
        state.heartbeatAt = folderCheckpointAt.toISOString();
      });
    }
    upsertEvidenceRunMetadata(executed.workbook, {
      runId,
      folderId: driveFolderId,
      folderUrl: driveFolderUrl,
      migrationVersion: EVIDENCE_MIGRATION_VERSION,
      timestamp:
        storedEvidenceRun?.timestamp ?? retestRun?.startedAt ?? new Date(),
      mode: mode === "retest" ? "RETEST" : "FUTURE",
    });
    if (retestRun) {
      retestRun.folderId = driveFolderId || undefined;
      retestRun.folderUrl = driveFolderUrl || undefined;
      retestRun.updatedAt = new Date();
      upsertRetestRunMetadata(executed.workbook, retestRun);
    }
    await saveExecutedPgnWorkbook(
      executed.workbook,
      config.pgnExecutedWorkbookPath,
    );
    workbookProgress = "Run metadata saved";
    const metadataSavedAt = new Date();
    await recoveryCheckpoint.update((state) => {
      state.status = "RUNNING";
      state.workbookProgress = workbookProgress;
      state.updatedAt = metadataSavedAt.toISOString();
      state.heartbeatAt = metadataSavedAt.toISOString();
    });
    throwIfInterrupted();
    const browserRequired =
      selectedScenarios.length > 0 ||
      !recoveryCheckpoint.snapshot().finalCleanupComplete;
    if (browserRequired) {
      const client = new WhatsAppClient(config, { handleProcessSignals: false });
      activeClient = client;
      try {
        failureStage = "opening WhatsApp Web";
        await client.open();
        throwIfInterrupted();
        failureStage = "authenticating WhatsApp Web";
        await client.ensureAuthenticated({ allowQrLogin: false });
        throwIfInterrupted();
        failureStage = "opening the configured WhatsApp chat";
        await client.openChat(requireTarget(config));
        throwIfInterrupted();

        for (
          let scenarioIndex = 0;
          scenarioIndex < selectedScenarios.length;
          scenarioIndex += 1
        ) {
          throwIfInterrupted();
          const scenario = selectedScenarios[scenarioIndex];
          currentScenarioId = scenario.testCaseId;
          failureStage = `preparing scenario ${scenario.testCaseId}`;
          const executions: ExecutedTurn[] = [];
          const attemptStartedAt = new Date();
          const attemptNumber =
            recoveryCheckpoint
              .snapshot()
              .scenarioAttempts.filter(
                (attempt) => attempt.scenarioId === scenario.testCaseId,
              ).length + 1;
          await recoveryCheckpoint.update((state) => {
            state.status = "RUNNING";
            state.activeScenarioId = scenario.testCaseId;
            state.activeScenarioAttempt = attemptNumber;
            state.activeScenarioStartedAt = attemptStartedAt.toISOString();
            state.scenarioAttempts.push({
              scenarioId: scenario.testCaseId,
              attempt: attemptNumber,
              status: "RUNNING",
              startedAt: attemptStartedAt.toISOString(),
            });
            state.updatedAt = attemptStartedAt.toISOString();
            state.heartbeatAt = attemptStartedAt.toISOString();
          });
          appendRecoveryTranscriptEvent(executed.workbook, {
            runId,
            event: "SCENARIO_ATTEMPT_STARTED",
            message: `Scenario attempt ${attemptNumber} started from Turn 1.`,
            timestamp: attemptStartedAt,
            scenario,
          });
          await saveExecutedPgnWorkbook(
            executed.workbook,
            config.pgnExecutedWorkbookPath,
          );
          workbookProgress = `Scenario ${scenario.testCaseId} attempt ${attemptNumber} started`;
          if (retestRun) {
            snapshotRetestHistory(
              executed.workbook,
              runId,
              scenario,
              attemptStartedAt,
            );
            await saveExecutedPgnWorkbook(
              executed.workbook,
              config.pgnExecutedWorkbookPath,
            );
          }
          if (shouldResetBeforeScenario(sessionMode, scenarioIndex)) {
            if (sessionMode === "continuous") console.log("[Session] Preparing one clean initial session for the continuous run");
            await performSessionReset(scenario);
          }
          throwIfInterrupted();
          console.log(
            `[Scenario] ${scenario.testCaseId} (${scenario.sheetName}, ${scenario.turns.length} turn(s))`,
          );
          failureStage = `executing scenario ${scenario.testCaseId}`;
          for (const turn of scenario.turns) {
            throwIfInterrupted();
            console.log(`[Turn ${turn.turnNumber}] Sending: ${turn.userInput}`);
            const execution = await executeTurn(
              client,
              config,
              runId,
              scenario,
              turn,
              driveEvidence,
            );
            throwIfInterrupted();
            executions.push(execution);
            if (
              execution.evidenceStatus &&
              execution.evidenceStatus !== "EVIDENCE_CAPTURE_ERROR" &&
              execution.evidenceStatus !== "EVIDENCE_MISSING" &&
              execution.evidenceStatus !== "EVIDENCE_REQUIRES_RERUN"
            ) {
              evidenceCapturedCount += 1;
            }
            if (execution.evidenceStatus === "EVIDENCE_SYNCED") {
              evidenceUploadedCount += 1;
            }
            if (execution.evidenceStatus === "EVIDENCE_UPLOAD_ERROR") {
              evidenceUploadErrorCount += 1;
            }
            upsertEvidenceFileMetadata(executed.workbook, {
              evidenceKey: `${runId}|${scenario.testCaseId}|${turn.turnNumber}`,
              runId,
              testCaseId: scenario.testCaseId,
              turnNumber: turn.turnNumber,
              driveFileId: execution.evidenceDriveFileId,
              driveFileName:
                execution.evidenceDriveFileName ??
                evidenceFileName(scenario.testCaseId, turn.turnNumber),
              evidenceUrl: execution.evidenceUrl,
              localCleanPath:
                execution.evidenceStatus === "EVIDENCE_CAPTURE_ERROR"
                  ? undefined
                  : execution.evidencePath,
              status: execution.evidenceStatus ?? "EVIDENCE_CAPTURE_ERROR",
            });
            appendLatestTurnExecution(
              executed.workbook,
              runId,
              scenario,
              executions,
            );
            applyScenarioResults(executed.workbook, scenario, executions);
            const scenarioFinished =
              execution.technicalStatus !== "CAPTURED" ||
              executions.length === scenario.turns.length;
            if (retestRun) {
              updateRetestHistory(
                executed.workbook,
                runId,
                scenario,
                executions,
              );
              if (scenarioFinished) {
                const successfullyCaptured = applyRetestStatusTransition(
                  executed.workbook,
                  scenario,
                  executions,
                );
                if (
                  successfullyCaptured &&
                  !retestRun.finishedIds.includes(scenario.testCaseId)
                ) {
                  retestRun.finishedIds.push(scenario.testCaseId);
                }
                retestRun.updatedAt = new Date();
                upsertRetestRunMetadata(executed.workbook, retestRun);
              }
            }
            await saveExecutedPgnWorkbook(
              executed.workbook,
              config.pgnExecutedWorkbookPath,
            );
            workbookProgress = `Saved through ${scenario.testCaseId} turn ${turn.turnNumber}`;
            const turnSavedAt = new Date();
            await recoveryCheckpoint.update((state) => {
              state.workbookProgress = workbookProgress;
              state.metrics.evidenceCaptured = evidenceCapturedCount;
              state.metrics.evidenceUploaded = evidenceUploadedCount;
              state.metrics.evidenceUploadErrors = evidenceUploadErrorCount;
              state.updatedAt = turnSavedAt.toISOString();
              state.heartbeatAt = turnSavedAt.toISOString();
            });
            console.log(
              `[Workbook] Saved after ${scenario.testCaseId} turn ${turn.turnNumber}`,
            );
            if (execution.combinedResponse) {
              console.log(`[Bot] ${execution.combinedResponse}`);
            }
            console.log(
              `[Turn ${turn.turnNumber}] ${execution.technicalStatus}; first=${execution.firstResponseMs ?? "n/a"} ms total=${execution.totalResponseMs ?? "n/a"} ms`,
            );
            if (execution.technicalStatus !== "CAPTURED") {
              console.log(
                `[Scenario] Stopping remaining turns after ${execution.technicalStatus}`,
              );
              break;
            }
          }
          executedCount += 1;
          const successfullyCaptured =
            executions.length === scenario.turns.length &&
            executions.every(
              (execution) => execution.technicalStatus === "CAPTURED",
            );
          if (successfullyCaptured) {
            capturedCount += 1;
            if (retestRun) awaitingEvaluationCount += 1;
          } else if (
            executions.some((execution) => execution.technicalStatus === "TIMEOUT")
          ) {
            timeoutCount += 1;
          } else {
            errorCount += 1;
          }
          const attemptFinishedAt = new Date();
          const recoveryScenarioCompleted =
            mode === "full" || successfullyCaptured;
          appendRecoveryTranscriptEvent(executed.workbook, {
            runId,
            event: successfullyCaptured
              ? "SCENARIO_ATTEMPT_COMPLETED"
              : "SCENARIO_ATTEMPT_FAILED",
            message: successfullyCaptured
              ? `Scenario attempt ${attemptNumber} completed and was checkpointed.`
              : mode === "full"
                ? `Scenario attempt ${attemptNumber} ended with a technical outcome and was checkpointed; remaining turns were not sent.`
                : `Scenario attempt ${attemptNumber} did not complete successfully and remains recoverable.`,
            timestamp: attemptFinishedAt,
            scenario,
          });
          await saveExecutedPgnWorkbook(
            executed.workbook,
            config.pgnExecutedWorkbookPath,
          );
          workbookProgress = `Scenario ${scenario.testCaseId} attempt ${attemptNumber} ${successfullyCaptured ? "completed" : "failed"}`;
          await recoveryCheckpoint.update((state) => {
            const attempt = [...state.scenarioAttempts]
              .reverse()
              .find(
                (item) =>
                  item.scenarioId === scenario.testCaseId &&
                  item.attempt === attemptNumber,
              );
            if (attempt) {
              attempt.status = successfullyCaptured ? "COMPLETED" : "FAILED";
              attempt.finishedAt = attemptFinishedAt.toISOString();
              if (!successfullyCaptured) {
                attempt.reason = "Scenario did not capture every turn";
              }
            }
            if (
              recoveryScenarioCompleted &&
              !state.completedScenarioIds.includes(scenario.testCaseId)
            ) {
              state.completedScenarioIds.push(scenario.testCaseId);
              state.lastCompletedScenarioId = scenario.testCaseId;
            }
            if (recoveryScenarioCompleted) {
              state.reconciliationDecisions =
                state.reconciliationDecisions.filter(
                  (decision) => decision.scenarioId !== scenario.testCaseId,
                );
            }
            state.activeScenarioId = undefined;
            state.activeScenarioAttempt = undefined;
            state.activeScenarioStartedAt = undefined;
            state.workbookProgress = workbookProgress;
            state.metrics = {
              executedScenarios: executedCount,
              capturedScenarios: capturedCount,
              timeouts: timeoutCount,
              technicalErrors: errorCount,
              evidenceCaptured: evidenceCapturedCount,
              evidenceUploaded: evidenceUploadedCount,
              evidenceUploadErrors: evidenceUploadErrorCount,
            };
            state.updatedAt = attemptFinishedAt.toISOString();
            state.heartbeatAt = attemptFinishedAt.toISOString();
          });
          currentScenarioId = undefined;
          requestProgressNotification();
        }

        if (sessionMode === "continuous") {
          await recoveryCheckpoint.update((state) => {
            state.finalCleanupComplete = true;
            state.workbookProgress = "Continuous run ended without a final reset";
          });
          workbookProgress = "Continuous run ended without a final reset";
        } else if (!recoveryCheckpoint.snapshot().finalCleanupComplete) {
          throwIfInterrupted();
          failureStage = "performing final bot session cleanup";
          const finalCleanupScenario = allRunScenarios.at(-1);
          if (!finalCleanupScenario) {
            throw new Error(
              "Could not identify a scenario for final session cleanup",
            );
          }
          await performSessionReset(finalCleanupScenario, true);
          const cleanupAt = new Date();
          await recoveryCheckpoint.update((state) => {
            state.finalCleanupComplete = true;
            state.workbookProgress = "Final bot session cleanup saved";
            state.updatedAt = cleanupAt.toISOString();
            state.heartbeatAt = cleanupAt.toISOString();
          });
          workbookProgress = "Final bot session cleanup saved";
        }
      } finally {
        await client.close();
        if (activeClient === client) activeClient = undefined;
      }
    }

    if (retestRun) {
      failureStage = "saving final retest metadata";
      const skippedIds = new Set(
        recoveryCheckpoint.snapshot().skippedScenarioIds,
      );
      retestRun.state = retestRun.selectedIds.every((testCaseId) =>
        retestRun!.finishedIds.includes(testCaseId) || skippedIds.has(testCaseId),
      )
        ? "COMPLETE"
        : "IN_PROGRESS";
      retestRun.updatedAt = new Date();
      upsertRetestRunMetadata(executed.workbook, retestRun);
      await saveExecutedPgnWorkbook(
        executed.workbook,
        config.pgnExecutedWorkbookPath,
      );
      workbookProgress = "Final retest metadata saved";
      console.log(
        retestRun.state === "COMPLETE"
          ? "PGN RETEST COMPLETE"
          : "PGN RETEST CHECKPOINT",
      );
      console.log(`Retest Run: ${runId}`);
      console.log(`Selected: ${selectedScenarios.length}`);
      if (finalCleanupOnly) {
        console.log("Execution: Final cleanup retry only");
      }
      console.log(`Executed: ${executedCount}`);
      console.log(`Captured: ${capturedCount}`);
      console.log(`Timeout: ${timeoutCount}`);
      console.log(`Errors: ${errorCount}`);
      console.log(`Skipped by Status: ${skippedByStatusCount}`);
      console.log(`Evidence Uploaded: ${evidenceUploadedCount}`);
      console.log(
        `Workbook: ${relativeToProject(config, config.pgnExecutedWorkbookPath)}`,
      );
      console.log(
        `Evidence Folder: ${driveFolderUrl || driveFolderId || "LOCAL ONLY"}`,
      );
      console.log(`Awaiting Evaluation: ${awaitingEvaluationCount}`);
      if (retestRun.state === "IN_PROGRESS") {
        console.log(
          `Remaining in Retest Run: ${retestRun.selectedIds.length - retestRun.finishedIds.length}`,
        );
      }
    }
    stopHeartbeatTimer();
    if (heartbeatOperation) await heartbeatOperation;
    const beforeCompletion = recoveryCheckpoint.snapshot();
    const resolvedIds = new Set([
      ...beforeCompletion.completedScenarioIds,
      ...beforeCompletion.skippedScenarioIds,
    ]);
    const runCompleted =
      beforeCompletion.selectedScenarioIds.every((id) => resolvedIds.has(id)) &&
      beforeCompletion.finalCleanupComplete;
    const completedAt = new Date();
    console.log(`[Summary] Transport: WhatsApp; Session Mode: ${sessionModeLabel(sessionMode)}`);
    console.log(`[Summary] Session reset attempts: ${sessionResetAttempts}${sessionMode === "isolated" ? " (scenario resets plus final cleanup)" : " (initial only)"}`);
    console.log(`[Summary] Duration: ${Math.round((completedAt.getTime() - notificationStartedAt.getTime()) / 1000)} s`);
    appendRecoveryTranscriptEvent(executed.workbook, {
      runId,
      event: runCompleted ? "RUN_COMPLETED" : "RUN_FAILED",
      message: runCompleted
        ? `Run completed with ${capturedCount} captured, ${errorCount + timeoutCount} technical-outcome, and ${beforeCompletion.skippedScenarioIds.length} operator-skipped scenario(s).`
        : `Run ended with ${beforeCompletion.totalScenarios - resolvedIds.size} incomplete scenario(s); ${sessionMode === "continuous" ? "only a full restart or abandonment is available" : "recovery remains available"}.`,
      timestamp: completedAt,
    });
    await saveExecutedPgnWorkbook(
      executed.workbook,
      config.pgnExecutedWorkbookPath,
    );
    workbookProgress = runCompleted
      ? "Final workbook and recovery audit saved"
      : "Recoverable checkpoint and audit saved";
    await recoveryCheckpoint.update((state) => {
      state.status = runCompleted ? "COMPLETED" : "RECOVERABLE";
      state.activeScenarioId = undefined;
      state.activeScenarioAttempt = undefined;
      state.activeScenarioStartedAt = undefined;
      state.workbookProgress = workbookProgress;
      state.metrics = {
        executedScenarios: executedCount,
        capturedScenarios: capturedCount,
        timeouts: timeoutCount,
        technicalErrors: errorCount,
        evidenceCaptured: evidenceCapturedCount,
        evidenceUploaded: evidenceUploadedCount,
        evidenceUploadErrors: evidenceUploadErrorCount,
      };
      state.updatedAt = completedAt.toISOString();
      state.heartbeatAt = completedAt.toISOString();
    });
    if (!retestRun) {
      console.log(
        `[Workbook] ${runCompleted ? "COMPLETE" : "RECOVERABLE CHECKPOINT"}: ${relativeToProject(config, config.pgnExecutedWorkbookPath)}`,
      );
    }
    stopProgressTimer();
    await trackNotification(
      notifier.runCompleted({
        ...notificationProgress(completedAt),
        completedAt,
        checkpoint: !runCompleted,
      }),
    );
    if (periodicProgress) await periodicProgress;
  } catch (error) {
    stopProgressTimer();
    stopHeartbeatTimer();
    if (heartbeatOperation) await heartbeatOperation;
    const failedAt = new Date();
    const reason = interruptionSignal
      ? `Process received ${interruptionSignal}`
      : `Technical failure while ${failureStage}`;
    appendRecoveryTranscriptEvent(executed.workbook, {
      runId,
      event: interruptionSignal ? "RUN_INTERRUPTED" : "RUN_FAILED",
      message: `${reason}. ${sessionMode === "continuous" ? CONTINUOUS_RECOVERY_WARNING : "The current scenario will restart from Turn 1 after explicit recovery confirmation."}`,
      timestamp: failedAt,
      scenario: currentScenarioId
        ? allRunScenarios.find(
            (scenario) => scenario.testCaseId === currentScenarioId,
          )
        : undefined,
    });
    await saveExecutedPgnWorkbook(
      executed.workbook,
      config.pgnExecutedWorkbookPath,
    ).catch((saveError) => {
      console.error(
        `[Recovery] Could not append interruption audit to workbook: ${saveError instanceof Error ? saveError.message : String(saveError)}`,
      );
    });
    await recoveryCheckpoint
      .update((state) => {
        if (state.activeScenarioId) {
          const attempt = [...state.scenarioAttempts]
            .reverse()
            .find(
              (item) =>
                item.scenarioId === state.activeScenarioId &&
                item.status === "RUNNING",
            );
          if (attempt) {
            attempt.status = interruptionSignal ? "INTERRUPTED" : "FAILED";
            attempt.finishedAt = failedAt.toISOString();
            attempt.reason = reason;
          }
        }
        state.status = interruptionSignal ? "INTERRUPTED" : "FAILED";
        state.interruptedAt = failedAt.toISOString();
        state.interruptionReason = reason;
        state.workbookProgress = workbookProgress;
        state.metrics = {
          executedScenarios: executedCount,
          capturedScenarios: capturedCount,
          timeouts: timeoutCount,
          technicalErrors: errorCount,
          evidenceCaptured: evidenceCapturedCount,
          evidenceUploaded: evidenceUploadedCount,
          evidenceUploadErrors: evidenceUploadErrorCount,
        };
        state.updatedAt = failedAt.toISOString();
        state.heartbeatAt = failedAt.toISOString();
      })
      .catch((checkpointError) => {
        console.error(
          `[Recovery] Could not save failure checkpoint: ${checkpointError instanceof Error ? checkpointError.message : String(checkpointError)}`,
        );
      });
    await trackNotification(
      notifier.runFailed({
        ...notificationProgress(failedAt),
        failedAt,
        reason: `${reason}.`,
        workbookProgress,
        evidenceProgress: `${evidenceCapturedCount} captured, ${evidenceUploadedCount} uploaded, ${evidenceUploadErrorCount} upload errors`,
      }),
    );
    if (periodicProgress) await periodicProgress;
    throw error;
  } finally {
    stopProgressTimer();
    stopHeartbeatTimer();
    if (heartbeatOperation) await heartbeatOperation;
    await recoveryCheckpoint.flush();
    unregisterInterruptionHandlers();
    resolveShutdownSettled();
  }
}

const entrypoint = process.argv[1];
if (entrypoint && import.meta.url === pathToFileURL(entrypoint).href) {
  runPgnWorkbook().catch((error) => {
    console.error(error instanceof Error ? error.message : error);
    process.exitCode = 1;
  });
}
