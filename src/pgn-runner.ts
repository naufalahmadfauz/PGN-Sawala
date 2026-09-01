import path from "node:path";
import { pathToFileURL } from "node:url";
import type { Workbook } from "exceljs";
import { loadConfig, requireTarget, type AppConfig } from "./config";
import {
  isScenarioComplete,
  isScenarioPartiallyComplete,
  loadPgnWorkbook,
} from "./excel/pgn-workbook-loader";
import {
  applyLatestScenarioExecution,
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
import { assertPgnWorkbookValid } from "./excel/pgn-workbook-validator";
import type {
  ExecutedTurn,
  PgnSheetKind,
  PgnTestScenario,
  PgnTestTurn,
  TechnicalStatus,
} from "./excel/pgn-types";
import type { BotSessionResetAttempt, SentMessage } from "./types";
import {
  createGoogleDriveEvidencePublisher,
  type EvidenceDrivePublisher,
} from "./evidence/google-drive";
import { evidenceFileName } from "./evidence/evidence-migration";
import { WhatsAppClient } from "./whatsapp/client";
import {
  BotSessionResetError,
  resetBotSession,
  waitForPostResetQuiet,
} from "./whatsapp/session-reset";

interface CliOptions {
  limit?: number;
  sheet?: PgnSheetKind;
  testIds: Set<string>;
  rerunAll: boolean;
  rerunIds: Set<string>;
}

interface RunEvidenceContext {
  publisher: EvidenceDrivePublisher;
  folderId: string;
}

function parseIdList(value: string): string[] {
  return value
    .split(",")
    .map((item) => item.trim())
    .filter(Boolean);
}

export function parseCliOptions(args: string[]): CliOptions {
  const options: CliOptions = {
    testIds: new Set(),
    rerunAll: false,
    rerunIds: new Set(),
  };

  for (let index = 0; index < args.length; index += 1) {
    const argument = args[index];
    if (argument === "--limit") {
      const value = Number(args[++index]);
      if (!Number.isInteger(value) || value < 1) {
        throw new Error("--limit requires a positive integer");
      }
      options.limit = value;
    } else if (argument === "--sheet") {
      const value = args[++index]?.toLowerCase();
      if (value === "kb" || value === "knowledge") {
        options.sheet = "kb";
      } else if (value === "negative" || value === "neg") {
        options.sheet = "negative";
      } else if (value === "all") {
        options.sheet = undefined;
      } else {
        throw new Error("--sheet must be kb, negative, or all");
      }
    } else if (argument === "--test") {
      const value = args[++index];
      if (!value || value.startsWith("--")) {
        throw new Error("--test requires a Test Case ID");
      }
      parseIdList(value).forEach((id) => options.testIds.add(id));
    } else if (argument === "--rerun") {
      const value = args[index + 1];
      if (value && !value.startsWith("--")) {
        index += 1;
        parseIdList(value).forEach((id) => options.rerunIds.add(id));
      } else {
        options.rerunAll = true;
      }
    } else {
      throw new Error(`Unknown argument: ${argument}`);
    }
  }
  return options;
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
          `[Evidence] EVIDENCE_UPLOAD_ERROR: ${error instanceof Error ? error.message : String(error)}`,
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

function selectScenarios(
  scenarios: PgnTestScenario[],
  options: CliOptions,
  workbook: Workbook,
): { runnable: PgnTestScenario[]; skipped: string[] } {
  const allIds = new Set(scenarios.map((scenario) => scenario.testCaseId));
  const requestedIds = new Set([...options.testIds, ...options.rerunIds]);
  for (const id of [...options.testIds, ...options.rerunIds]) {
    if (!allIds.has(id)) {
      throw new Error(`Test Case ID was not found: ${id}`);
    }
  }

  let selected = scenarios.filter(
    (scenario) => !options.sheet || scenario.sheetKind === options.sheet,
  );
  if (requestedIds.size > 0) {
    const selectedIds = new Set(
      selected.map((scenario) => scenario.testCaseId),
    );
    for (const id of requestedIds) {
      if (!selectedIds.has(id)) {
        throw new Error(
          `Test Case ID ${id} is not in the selected ${options.sheet} sheet`,
        );
      }
    }
    selected = selected.filter((scenario) =>
      requestedIds.has(scenario.testCaseId),
    );
  }

  const skipped: string[] = [];
  const runnable = selected.filter((scenario) => {
    const rerun = options.rerunAll || options.rerunIds.has(scenario.testCaseId);
    if (!rerun && isScenarioPartiallyComplete(workbook, scenario)) {
      skipped.push(
        `${scenario.testCaseId}: partially completed; use --rerun ${scenario.testCaseId} to preserve multi-turn context`,
      );
      return false;
    }
    if (!rerun && isScenarioComplete(workbook, scenario)) {
      skipped.push(`${scenario.testCaseId}: already completed`);
      return false;
    }
    return true;
  });

  return {
    runnable: options.limit ? runnable.slice(0, options.limit) : runnable,
    skipped,
  };
}

export async function runPgnWorkbook(args = process.argv.slice(2)): Promise<void> {
  const options = parseCliOptions(args);
  const config = loadConfig();
  const source = await loadPgnWorkbook(config.pgnSourceWorkbookPath);
  assertPgnWorkbookValid(source.parsed);
  const executed = await openExecutedPgnWorkbook(
    config.pgnSourceWorkbookPath,
    config.pgnExecutedWorkbookPath,
  );
  assertPgnWorkbookValid(executed.parsed);

  const selection = selectScenarios(
    executed.parsed.scenarios,
    options,
    executed.workbook,
  );
  console.log(
    `[Workbook] ${executed.resumed ? "Resuming" : "Created"}: ${relativeToProject(config, config.pgnExecutedWorkbookPath)}`,
  );
  selection.skipped.forEach((message) => console.log(`[Skip] ${message}`));
  if (selection.runnable.length === 0) {
    console.log("[Test] No scenarios require execution");
    return;
  }
  console.log(
    `[Test] Selected ${selection.runnable.length} scenario(s), ${selection.runnable.reduce((count, scenario) => count + scenario.turns.length, 0)} turn(s)`,
  );
  console.log(
    `[Session] Isolation enabled: send "${config.resetCommand}" and require "${config.resetConfirmation}" before every scenario`,
  );

  const runId = createRunId();
  let driveEvidence: RunEvidenceContext | undefined;
  if (config.googleDriveEvidenceEnabled) {
    const publisher = createGoogleDriveEvidencePublisher(config);
    await publisher.validateParentFolder();
    const storedRun = getEvidenceRunMetadata(executed.workbook, runId);
    const folder = await publisher.ensureRunFolder(runId, storedRun?.folderId);
    upsertEvidenceRunMetadata(executed.workbook, {
      runId,
      folderId: folder.id,
      folderUrl: folder.webViewLink,
      migrationVersion: EVIDENCE_MIGRATION_VERSION,
      timestamp: new Date(),
      mode: "FUTURE",
    });
    await saveExecutedPgnWorkbook(
      executed.workbook,
      config.pgnExecutedWorkbookPath,
    );
    driveEvidence = { publisher, folderId: folder.id };
    console.log(`[Evidence] Drive run folder ready: ${folder.name}`);
  }
  const client = new WhatsAppClient(config);
  try {
    await client.open();
    await client.ensureAuthenticated({ allowQrLogin: false });
    await client.openChat(requireTarget(config));

    for (let scenarioIndex = 0; scenarioIndex < selection.runnable.length; scenarioIndex += 1) {
      const scenario = selection.runnable[scenarioIndex];
      const executions: ExecutedTurn[] = [];
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
        if (driveEvidence) {
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
            localCleanPath: execution.evidencePath,
            status: execution.evidenceStatus ?? "EVIDENCE_CAPTURE_ERROR",
          });
        }
        applyLatestScenarioExecution(
          executed.workbook,
          runId,
          scenario,
          executions,
        );
        await saveExecutedPgnWorkbook(
          executed.workbook,
          config.pgnExecutedWorkbookPath,
        );
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

    }

    await resetAndDrainSession(
      client,
      config,
      runId,
      selection.runnable.at(-1)!,
      executed.workbook,
      true,
    );
  } finally {
    await client.close();
  }

  console.log(
    `[Workbook] COMPLETE: ${relativeToProject(config, config.pgnExecutedWorkbookPath)}`,
  );
}

const entrypoint = process.argv[1];
if (entrypoint && import.meta.url === pathToFileURL(entrypoint).href) {
  runPgnWorkbook().catch((error) => {
    console.error(error instanceof Error ? error.message : error);
    process.exitCode = 1;
  });
}
