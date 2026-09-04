import path from "node:path";
import { pathToFileURL } from "node:url";
import type { Workbook } from "exceljs";
import { loadConfig, requireTarget, type AppConfig } from "./config";
import { loadPgnWorkbook } from "./excel/pgn-workbook-loader";
import {
  appendLatestTurnExecution,
  applyScenarioResults,
  appendPostResetDrainTranscript,
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
  validateDiscordWebhookUrl,
  type DiscordRunProgressEvent,
} from "./notifications/discord";
import { parseCliOptions, type CliOptions } from "./pgn-cli";
import { selectScenarios } from "./pgn-selection";
import {
  createRetestRunId,
  needsFinalRetestCleanup,
  retestDriveFolderName,
} from "./retest/retest-run";
import { selectRetestScenarios } from "./retest/retest-selection";
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
): Promise<void> {
  const options = parseCliOptions(args);
  if (mode === "full" && options.resumeRunId) {
    throw new Error("--resume is only available in retest mode");
  }
  const config = loadConfig();
  try {
    const releaseWorkbookLock = await acquireWorkbookLock(
      config.pgnExecutedWorkbookPath,
      mode === "retest" ? "PGN retest runner" : "PGN test runner",
    );
    try {
      await runPgnWorkbookLocked(options, config, mode);
    } finally {
      await releaseWorkbookLock();
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
): Promise<void> {
  const source = await loadPgnWorkbook(config.pgnSourceWorkbookPath);
  assertPgnWorkbookValid(source.parsed);
  const executed = await openExecutedPgnWorkbook(
    config.pgnSourceWorkbookPath,
    config.pgnExecutedWorkbookPath,
  );
  assertPgnWorkbookValid(executed.parsed);

  console.log(
    `[Workbook] ${executed.resumed ? "Resuming" : "Created"}: ${relativeToProject(config, config.pgnExecutedWorkbookPath)}`,
  );
  let selectedScenarios: PgnTestScenario[];
  let runId: string;
  let retestRun: RetestRunMetadata | undefined;
  let skippedByStatusCount = 0;
  let finalCleanupOnly = false;
  if (mode === "retest") {
    if (options.rerunAll || options.rerunIds.size) {
      throw new Error("--rerun is not used in retest mode; use --test instead");
    }
    if (
      options.resumeRunId &&
      (options.testIds.size > 0 || options.sheet !== undefined)
    ) {
      throw new Error("--resume cannot be combined with --test or --sheet");
    }
    const resumedRun = options.resumeRunId
      ? getRetestRunMetadata(executed.workbook, options.resumeRunId)
      : undefined;
    if (options.resumeRunId && !resumedRun) {
      throw new Error(`Retest Run was not found: ${options.resumeRunId}`);
    }
    const retestSelection = selectRetestScenarios(executed.parsed.scenarios, {
      testIds: options.testIds,
      sheet: options.sheet,
      limit: options.limit,
      resumeSelectedIds: resumedRun?.selectedIds,
      completedIds: new Set(resumedRun?.finishedIds ?? []),
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
      if (needsFinalRetestCleanup(resumedRun, selectedScenarios.length)) {
        finalCleanupOnly = true;
        console.log(
          "All selected scenarios are complete; retrying final session cleanup.",
        );
      } else {
        console.log("Nothing to execute.");
        return;
      }
    }
    if (!config.googleDriveEvidenceEnabled) {
      throw new Error(
        "Retest mode requires Google Drive evidence. Configure Drive before launching the selected retests.",
      );
    }
    const startedAt = resumedRun?.startedAt ?? new Date();
    runId =
      resumedRun?.runId ?? uniqueRetestRunId(executed.workbook, startedAt);
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
    await saveExecutedPgnWorkbook(
      executed.workbook,
      config.pgnExecutedWorkbookPath,
    );
    console.log(`Retest Run: ${runId}`);
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
    runId = createRunId();
  }

  console.log(
    `[Test] Selected ${selectedScenarios.length} scenario(s), ${selectedScenarios.reduce((count, scenario) => count + scenario.turns.length, 0)} turn(s)`,
  );
  console.log(
    `[Session] Isolation enabled: send "${config.resetCommand}" and require "${config.resetConfirmation}" before every scenario`,
  );

  const notificationStartedAt = new Date();
  let executedCount = 0;
  let capturedCount = 0;
  let timeoutCount = 0;
  let errorCount = 0;
  let evidenceUploadedCount = 0;
  let evidenceUploadErrorCount = 0;
  let awaitingEvaluationCount = 0;
  let currentScenarioId: string | undefined;
  let workbookProgress = retestRun
    ? "Retest run metadata saved"
    : "No scenario results saved yet";
  let failureStage = "initializing the active test run";
  const notifier = createDiscordNotifier(config);
  const notificationProgress = (updatedAt = new Date()): DiscordRunProgressEvent => ({
    completedScenarios: executedCount,
    totalScenarios: selectedScenarios.length,
    currentScenarioId,
    capturedScenarios: capturedCount,
    timeouts: timeoutCount,
    technicalErrors: errorCount,
    evidenceUploaded: evidenceUploadedCount,
    evidenceUploadErrors: evidenceUploadErrorCount,
    updatedAt,
  });
  const discordHandlesProcessSignals =
    config.discordNotificationsEnabled &&
    (config.discordNotifyFailure ||
      config.discordNotifyStart ||
      config.discordNotifyProgress) &&
    validateDiscordWebhookUrl(config.discordWebhookUrl).valid;
  let progressTimer: NodeJS.Timeout | undefined;
  let periodicProgress: Promise<void> | undefined;
  let activeClient: WhatsAppClient | undefined;
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
  const unregisterInterruptionHandlers = discordHandlesProcessSignals
    ? registerDiscordInterruptionHandlers({
        notifier,
        progress: () => notificationProgress(),
        settle: () =>
          Promise.all([...notificationOperations]).then(() => undefined),
        cleanup: async () => {
          stopProgressTimer();
          await activeClient?.close();
        },
      })
    : () => undefined;

  try {
    await trackNotification(
      notifier.runStarted({
        runId,
        mode,
        selectedScenarios: selectedScenarios.length,
        startedAt: notificationStartedAt,
        googleDriveEvidenceEnabled: config.googleDriveEvidenceEnabled,
        workbookPath: config.pgnExecutedWorkbookPath,
      }),
    );
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
    let driveFolderId = storedEvidenceRun?.folderId ?? "";
    let driveFolderUrl = storedEvidenceRun?.folderUrl ?? "";
    if (config.googleDriveEvidenceEnabled) {
      const publisher = createGoogleDriveEvidencePublisher(config);
      const parent = await publisher.validateParentFolder();
      console.log(`[Evidence] Drive parent ready: ${parent.name} (${parent.id})`);
      const folder = await publisher.ensureRunFolder(
        runId,
        storedEvidenceRun?.folderId,
        mode === "retest"
          ? retestDriveFolderName(config.googleDriveRetestFolderPrefix, runId)
          : undefined,
      );
      driveFolderId = folder.id;
      driveFolderUrl = folder.webViewLink;
      driveEvidence = { publisher, folderId: folder.id };
      console.log(`[Evidence] Drive run folder ready: ${folder.name}`);
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
    const client = new WhatsAppClient(config, {
      handleProcessSignals: !discordHandlesProcessSignals,
    });
    activeClient = client;
    try {
      failureStage = "opening WhatsApp Web";
      await client.open();
      failureStage = "authenticating WhatsApp Web";
      await client.ensureAuthenticated({ allowQrLogin: false });
      failureStage = "opening the configured WhatsApp chat";
      await client.openChat(requireTarget(config));

      for (
        let scenarioIndex = 0;
        scenarioIndex < selectedScenarios.length;
        scenarioIndex += 1
      ) {
        const scenario = selectedScenarios[scenarioIndex];
        currentScenarioId = scenario.testCaseId;
        failureStage = `preparing scenario ${scenario.testCaseId}`;
        const executions: ExecutedTurn[] = [];
        if (retestRun) {
          snapshotRetestHistory(
            executed.workbook,
            runId,
            scenario,
            new Date(),
          );
          await saveExecutedPgnWorkbook(
            executed.workbook,
            config.pgnExecutedWorkbookPath,
          );
        }
        await resetAndDrainSession(
          client,
          config,
          runId,
          scenario,
          executed.workbook,
        );
        console.log(
          `[Scenario] ${scenario.testCaseId} (${scenario.sheetName}, ${scenario.turns.length} turn(s))`,
        );
        failureStage = `executing scenario ${scenario.testCaseId}`;
        for (const turn of scenario.turns) {
          console.log(`[Turn ${turn.turnNumber}] Sending: ${turn.userInput}`);
          const execution = await executeTurn(
            client,
            config,
            runId,
            scenario,
            turn,
            driveEvidence,
          );
          executions.push(execution);
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
        if (
          executions.length === scenario.turns.length &&
          executions.every(
            (execution) => execution.technicalStatus === "CAPTURED",
          )
        ) {
          capturedCount += 1;
          if (retestRun) {
            awaitingEvaluationCount += 1;
          }
        } else if (
          executions.some((execution) => execution.technicalStatus === "TIMEOUT")
        ) {
          timeoutCount += 1;
        } else {
          errorCount += 1;
        }
        requestProgressNotification();
      }

      failureStage = "performing final bot session cleanup";
      const finalCleanupScenario =
        selectedScenarios.at(-1) ??
        (retestRun
          ? executed.parsed.scenarios.find(
              (scenario) =>
                scenario.testCaseId === retestRun.selectedIds.at(-1),
            )
          : undefined);
      if (!finalCleanupScenario) {
        throw new Error(
          "Could not identify a scenario for final session cleanup",
        );
      }
      await resetAndDrainSession(
        client,
        config,
        runId,
        finalCleanupScenario,
        executed.workbook,
        true,
      );
    } finally {
      await client.close();
      if (activeClient === client) activeClient = undefined;
    }

    if (retestRun) {
      failureStage = "saving final retest metadata";
      retestRun.state = retestRun.selectedIds.every((testCaseId) =>
        retestRun!.finishedIds.includes(testCaseId),
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
    } else {
      console.log(
        `[Workbook] COMPLETE: ${relativeToProject(config, config.pgnExecutedWorkbookPath)}`,
      );
    }
    stopProgressTimer();
    const completedAt = new Date();
    await trackNotification(
      notifier.runCompleted({
        ...notificationProgress(completedAt),
        completedAt,
        checkpoint: retestRun?.state === "IN_PROGRESS",
      }),
    );
    if (periodicProgress) await periodicProgress;
  } catch (error) {
    stopProgressTimer();
    const failedAt = new Date();
    await trackNotification(
      notifier.runFailed({
        ...notificationProgress(failedAt),
        failedAt,
        reason: `Technical failure while ${failureStage}.`,
        workbookProgress,
        evidenceProgress: `${evidenceUploadedCount} uploaded, ${evidenceUploadErrorCount} upload errors`,
      }),
    );
    if (periodicProgress) await periodicProgress;
    throw error;
  } finally {
    stopProgressTimer();
    unregisterInterruptionHandlers();
  }
}

const entrypoint = process.argv[1];
if (entrypoint && import.meta.url === pathToFileURL(entrypoint).href) {
  runPgnWorkbook().catch((error) => {
    console.error(error instanceof Error ? error.message : error);
    process.exitCode = 1;
  });
}
