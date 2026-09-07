import { access, readdir } from "node:fs/promises";
import path from "node:path";
import type { Workbook } from "exceljs";
import type { AppConfig } from "../config";
import { acquireWorkbookLock } from "../excel/workbook-lock";
import {
  getEvidenceRunMetadata,
} from "../excel/evidence-workbook";
import {
  getRetestRunMetadata,
  upsertRetestRunMetadata,
} from "../excel/retest-workbook";
import {
  isScenarioComplete,
  loadPgnWorkbook,
} from "../excel/pgn-workbook-loader";
import {
  appendRecoveryTranscriptEvent,
  openExecutedPgnWorkbook,
  saveExecutedPgnWorkbook,
} from "../excel/pgn-workbook-writer";
import {
  TRANSCRIPT_SHEET_NAME,
  type PgnTestScenario,
} from "../excel/pgn-types";
import {
  createGoogleDriveEvidencePublisher,
  type EvidenceDrivePublisher,
} from "../evidence/google-drive";
import { safeGoogleCredentialError } from "../evidence/google-service-account";
import { validateDiscordWebhookUrl } from "../notifications/discord";
import { retestDriveFolderName } from "../retest/retest-run";
import { recoveryDemoConfig } from "./demo-safety";
import { getRunConfiguration } from "../excel/run-configuration";
import { CONTINUOUS_RECOVERY_WARNING, readSessionMode, readExecutionTransport, sessionModeLabel } from "../session-mode";
import { fieldCell, optionalFieldCell, EVIDENCE_FILE_SCHEMA, getWorksheetSchema, normalizeWorkbookHeader } from "../excel/workbook-schema";
import {
  acquireRunProcessLock,
  createRecoveryManifest,
  discoverRecoveryRun,
  hashFile,
  inspectRunProcessLock,
  openRecoveryCheckpoint,
  readRecoveryRun,
  recoveryManifestMatches,
  type RecoveryDiscovery,
  type RecoveryRunManifest,
  type RecoveryRunState,
  type RunLockInspection,
} from "./run-state";

export type RecoveryCheckStatus = "ok" | "warn" | "error" | "info";

export interface RecoveryValidationCheck {
  id: string;
  label: string;
  status: RecoveryCheckStatus;
  detail: string;
}

export interface RecoveryReconciliation {
  checkpointCompletedIds: string[];
  workbookCompletedIds: string[];
  transcriptCompletedIds: string[];
  artifactConfirmedIds: string[];
  safeCompletedIds: string[];
  reconciledScenarioIds: string[];
  mismatchedScenarioIds: string[];
  evidenceCapturedScenarioIds: string[];
  evidenceUploadedScenarioIds: string[];
  nextScenarioId?: string;
  interruptedScenarioId?: string;
  restartInterruptedScenarioFromTurnOne: boolean;
}

export type RecoverySourceDrift =
  | "unchanged"
  | "formatting-only"
  | "structural"
  | "unavailable";

export interface RecoveryValidation {
  runId: string;
  mode: "full" | "retest";
  state: RecoveryRunState;
  manifest: RecoveryRunManifest;
  lock: RunLockInspection;
  sourceDrift: RecoverySourceDrift;
  reconciliation?: RecoveryReconciliation;
  checks: RecoveryValidationCheck[];
  ready: boolean;
  restartReady?: boolean;
}

export interface RecoveryValidationDependencies {
  checkDriveAccess?: boolean;
  drivePublisher?: EvidenceDrivePublisher;
  profileEntries?: (profilePath: string) => Promise<string[]>;
  fileExists?: (filePath: string) => Promise<boolean>;
}

export type RecoveryRepairStrategy = "rerun" | "checkpoint" | "artifacts";

export interface RecoveryMutationResult {
  runId: string;
  scenarioId?: string;
  warning?: string;
}

async function defaultExists(filePath: string): Promise<boolean> {
  return access(filePath)
    .then(() => true)
    .catch(() => false);
}

function samePath(left: string, right: string): boolean {
  return path.resolve(left) === path.resolve(right);
}

function expectedRunFolderName(config: AppConfig, state: RecoveryRunState): string {
  return state.mode === "retest"
    ? retestDriveFolderName(config.googleDriveRetestFolderPrefix, state.runId)
    : `${config.googleDriveEvidenceFolderPrefix}-${state.runId}`;
}

function selectedScenarios(
  scenarios: readonly PgnTestScenario[],
  state: RecoveryRunState,
): PgnTestScenario[] {
  const byId = new Map(scenarios.map((scenario) => [scenario.testCaseId, scenario]));
  return state.selectedScenarioIds.map((testCaseId) => {
    const scenario = byId.get(testCaseId);
    if (!scenario) {
      throw new Error(`Recovery scenario is missing from the workbook: ${testCaseId}`);
    }
    return scenario;
  });
}

export function selectRecoveryScenarios(
  scenarios: readonly PgnTestScenario[],
  state: RecoveryRunState,
): PgnTestScenario[] {
  const completed = new Set([
    ...state.completedScenarioIds,
    ...state.skippedScenarioIds,
  ]);
  return selectedScenarios(scenarios, state).filter(
    (scenario) => !completed.has(scenario.testCaseId),
  );
}

function transcriptCompletedScenarioIds(
  workbook: Workbook,
  state: RecoveryRunState,
  manifest: RecoveryRunManifest,
  technicalOnly = false,
): Set<string> {
  const transcript = workbook.getWorksheet(TRANSCRIPT_SHEET_NAME);
  if (!transcript) return new Set();
  const selected = new Set(state.selectedScenarioIds);
  const latestAttemptTurns = new Map<string, Set<number>>();
  const hasAttemptMarker = new Set<string>();
  const explicitlyCompleted = new Set<string>();
  for (let rowNumber = 2; rowNumber <= transcript.rowCount; rowNumber += 1) {
    if (fieldCell(transcript, rowNumber, "runId").text !== state.runId) continue;
    const scenarioId = fieldCell(transcript, rowNumber, "testCaseId").text;
    if (!selected.has(scenarioId)) continue;
    const event = fieldCell(transcript, rowNumber, "status").text;
    if (event === "SCENARIO_ATTEMPT_STARTED") {
      hasAttemptMarker.add(scenarioId);
      latestAttemptTurns.set(scenarioId, new Set());
      explicitlyCompleted.delete(scenarioId);
      continue;
    }
    if (event === "SCENARIO_ATTEMPT_COMPLETED") {
      if (technicalOnly) explicitlyCompleted.delete(scenarioId);
      else explicitlyCompleted.add(scenarioId);
      continue;
    }
    if (event === "SCENARIO_ATTEMPT_FAILED" && state.mode === "full") {
      explicitlyCompleted.add(scenarioId);
      continue;
    }
    if (technicalOnly || fieldCell(transcript, rowNumber, "role").text !== "USER" || event !== "CAPTURED") continue;
    const turns = latestAttemptTurns.get(scenarioId) ?? new Set<number>();
    const turn = Number(fieldCell(transcript, rowNumber, "turn").value);
    if (Number.isInteger(turn) && turn > 0) turns.add(turn);
    latestAttemptTurns.set(scenarioId, turns);
  }
  const completed = new Set(explicitlyCompleted);
  if (technicalOnly) return completed;
  for (const item of manifest.scenarios) {
    const turns = latestAttemptTurns.get(item.testCaseId);
    if (
      turns?.size === item.turnCount &&
      (!hasAttemptMarker.has(item.testCaseId) || turns.has(item.turnCount))
    ) {
      completed.add(item.testCaseId);
    }
  }
  return completed;
}

function workbookCompletedScenarioIds(
  workbook: Workbook,
  scenarios: readonly PgnTestScenario[],
  state: RecoveryRunState,
  manifest: RecoveryRunManifest,
): Set<string> {
  if (state.mode === "retest") {
    return new Set(getRetestRunMetadata(workbook, state.runId)?.finishedIds ?? []);
  }
  const technicalCompleted = transcriptCompletedScenarioIds(workbook, state, manifest, true);
  return new Set(
    scenarios
      .filter((scenario) => {
        if (isScenarioComplete(workbook, scenario)) return true;
        const worksheet = workbook.getWorksheet(scenario.sheetName);
        if (!worksheet) return false;
        const mapping = getWorksheetSchema(worksheet);
        if (!mapping.fields.notes || !mapping.fields.testDate) {
          // Optional reporting fields cannot be required to recover a terminal technical attempt.
          return technicalCompleted.has(scenario.testCaseId);
        }
        const rows =
          scenario.sheetKind === "kb"
            ? scenario.turns.map((turn) => turn.rowNumber)
            : [scenario.sourceRowNumber];
        return rows.some((rowNumber) => {
          const executedAt = optionalFieldCell(worksheet, rowNumber, "testDate")?.value;
          // Notes survive reruns; only the current execution's marker counts.
          return (
            executedAt instanceof Date &&
            !Number.isNaN(executedAt.getTime()) &&
            Boolean(optionalFieldCell(worksheet, rowNumber, "notes")?.text.includes(
              `[Technical execution ${executedAt.toISOString()}]`,
            ))
          );
        });
      })
      .map((scenario) => scenario.testCaseId),
  );
}

function evidenceScenarioIds(
  workbook: Workbook,
  state: RecoveryRunState,
  manifest: RecoveryRunManifest,
): { captured: Set<string>; uploaded: Set<string> } {
  const metadata = workbook.getWorksheet("Execution Metadata");
  if (!metadata) return { captured: new Set(), uploaded: new Set() };
  const capturedTurns = new Map<string, Set<number>>();
  const uploadedTurns = new Map<string, Set<number>>();
  const selected = new Set(state.selectedScenarioIds);
  for (let rowNumber = 2; rowNumber <= metadata.rowCount; rowNumber += 1) {
    if (fieldCell(metadata, rowNumber, "runId", EVIDENCE_FILE_SCHEMA).text !== state.runId) continue;
    const scenarioId = fieldCell(metadata, rowNumber, "testCaseId", EVIDENCE_FILE_SCHEMA).text;
    if (!selected.has(scenarioId)) continue;
    const turn = Number(fieldCell(metadata, rowNumber, "turn", EVIDENCE_FILE_SCHEMA).value);
    if (!Number.isInteger(turn) || turn < 1) continue;
    const status = fieldCell(metadata, rowNumber, "evidenceStatus", EVIDENCE_FILE_SCHEMA).text;
    if (
      status &&
      status !== "EVIDENCE_CAPTURE_ERROR" &&
      status !== "EVIDENCE_MISSING" &&
      status !== "EVIDENCE_REQUIRES_RERUN"
    ) {
      const turns = capturedTurns.get(scenarioId) ?? new Set<number>();
      turns.add(turn);
      capturedTurns.set(scenarioId, turns);
    }
    if (status === "EVIDENCE_SYNCED" || status === "EVIDENCE_ALREADY_SYNCED") {
      const turns = uploadedTurns.get(scenarioId) ?? new Set<number>();
      turns.add(turn);
      uploadedTurns.set(scenarioId, turns);
    }
  }
  const captured = new Set<string>();
  const uploaded = new Set<string>();
  for (const scenario of manifest.scenarios) {
    if (capturedTurns.get(scenario.testCaseId)?.size === scenario.turnCount) {
      captured.add(scenario.testCaseId);
    }
    if (uploadedTurns.get(scenario.testCaseId)?.size === scenario.turnCount) {
      uploaded.add(scenario.testCaseId);
    }
  }
  return { captured, uploaded };
}

function orderedSubset(order: readonly string[], values: ReadonlySet<string>): string[] {
  return order.filter((value) => values.has(value));
}

export function reconcileRecoveryArtifacts(
  workbook: Workbook,
  scenarios: readonly PgnTestScenario[],
  state: RecoveryRunState,
  manifest: RecoveryRunManifest,
): RecoveryReconciliation {
  const checkpoint = new Set(state.completedScenarioIds);
  const workbookCompleted = workbookCompletedScenarioIds(workbook, scenarios, state, manifest);
  const transcriptCompleted = transcriptCompletedScenarioIds(workbook, state, manifest);
  const skipped = new Set(state.skippedScenarioIds);
  const artifactConfirmed = new Set(
    [...workbookCompleted].filter((id) => transcriptCompleted.has(id)),
  );
  const safeCompleted = new Set(
    [...checkpoint].filter((id) => artifactConfirmed.has(id)),
  );
  const mismatched = new Set<string>();
  const reconciled = new Set(
    state.reconciliationDecisions.map((decision) => decision.scenarioId),
  );
  for (const id of state.selectedScenarioIds) {
    if (skipped.has(id)) continue;
    const values = [
      checkpoint.has(id),
      workbookCompleted.has(id),
      transcriptCompleted.has(id),
    ];
    if (
      values.some(Boolean) &&
      !values.every(Boolean) &&
      !reconciled.has(id)
    ) {
      mismatched.add(id);
    }
  }
  const evidence = evidenceScenarioIds(workbook, state, manifest);
  const finished = new Set([...state.completedScenarioIds, ...state.skippedScenarioIds]);
  const nextScenarioId = state.selectedScenarioIds.find((id) => !finished.has(id));
  return {
    checkpointCompletedIds: orderedSubset(state.selectedScenarioIds, checkpoint),
    workbookCompletedIds: orderedSubset(state.selectedScenarioIds, workbookCompleted),
    transcriptCompletedIds: orderedSubset(state.selectedScenarioIds, transcriptCompleted),
    artifactConfirmedIds: orderedSubset(state.selectedScenarioIds, artifactConfirmed),
    safeCompletedIds: orderedSubset(state.selectedScenarioIds, safeCompleted),
    reconciledScenarioIds: orderedSubset(state.selectedScenarioIds, reconciled),
    mismatchedScenarioIds: orderedSubset(state.selectedScenarioIds, mismatched),
    evidenceCapturedScenarioIds: orderedSubset(state.selectedScenarioIds, evidence.captured),
    evidenceUploadedScenarioIds: orderedSubset(state.selectedScenarioIds, evidence.uploaded),
    nextScenarioId,
    interruptedScenarioId: state.activeScenarioId,
    restartInterruptedScenarioFromTurnOne: Boolean(
      state.activeScenarioId && !finished.has(state.activeScenarioId),
    ),
  };
}

function lockCheck(lock: RunLockInspection): RecoveryValidationCheck {
  if (lock.status === "unlocked") {
    return { id: "lock", label: "Process lock", status: "ok", detail: "available" };
  }
  if (lock.status === "active") {
    return {
      id: "lock",
      label: "Process lock",
      status: "error",
      detail: `run is active in PID ${lock.record.pid}`,
    };
  }
  if (lock.status === "stale") {
    return {
      id: "lock",
      label: "Process lock",
      status: "warn",
      detail: `stale lock from PID ${lock.record.pid}; it will be recovered on resume`,
    };
  }
  return { id: "lock", label: "Process lock", status: "error", detail: lock.reason };
}

function legacyRecoverySchemasMatch(workbook: Workbook, manifest: RecoveryRunManifest): boolean {
  // Shipped checkpoints without fingerprints came from the canonical fixed-layout engine.
  return manifest.scenarios.filter((item) => !item.schemaFingerprint).every((item) => {
    const sheet = workbook.getWorksheet(item.sheetName);
    if (!sheet) return false;
    const mapping = getWorksheetSchema(sheet);
    return mapping.definition.fields.filter((field) => field.field !== "evidence").every((field, index) =>
      mapping.fields[field.field]?.columnIndex === index + 1 &&
      normalizeWorkbookHeader(mapping.fields[field.field]?.header ?? "") === normalizeWorkbookHeader(field.header),
    );
  });
}

export async function validateRecoveryRun(
  config: AppConfig,
  runId?: string,
  dependencies: RecoveryValidationDependencies = {},
): Promise<RecoveryValidation> {
  let recovered: Awaited<ReturnType<typeof readRecoveryRun>>;
  if (runId) {
    recovered = await readRecoveryRun(config.projectRoot, runId);
  } else {
    const discovered = await discoverRecoveryRun(config.projectRoot);
    if (discovered.kind === "none") throw new Error("No recoverable PGN run was found");
    if (discovered.kind === "unreadable") {
      throw new Error(`Recovery state is unreadable: ${discovered.reason}`);
    }
    recovered = { state: discovered.state, manifest: discovered.manifest };
  }
  const { state, manifest } = recovered;
  config = await recoveryDemoConfig(config, state);
  const sessionMode = readSessionMode(state.sessionMode);
  const checks: RecoveryValidationCheck[] = [];
  const lock = await inspectRunProcessLock(config.projectRoot);
  checks.push(lockCheck(lock));
  checks.push({
    id: "state",
    label: "Checkpoint",
    status:
      state.status === "COMPLETED" || state.status === "ABANDONED" ? "error" : "ok",
    detail: `${state.status}; updated ${state.updatedAt}`,
  });
  if (sessionMode === "continuous") {
    checks.push({ id: "session-continuity", label: "Continuous recovery", status: "error", detail: CONTINUOUS_RECOVERY_WARNING });
  }
  if (state.mode !== "full" && state.mode !== "retest") {
    checks.push({ id: "mode", label: "Run mode", status: "error", detail: "invalid" });
  }

  const storedSourcePath = path.resolve(config.projectRoot, state.sourceWorkbookPath);
  const storedExecutedPath = path.resolve(config.projectRoot, state.executedWorkbookPath);
  const sourcePathMatches = samePath(
    storedSourcePath,
    config.pgnSourceWorkbookPath,
  );
  const executedPathMatches = samePath(
    storedExecutedPath,
    config.pgnExecutedWorkbookPath,
  );
  if (!sourcePathMatches) {
    checks.push({
      id: "source-path",
      label: "Source workbook path",
      status: "error",
      detail: "current configuration points to a different source workbook",
    });
  }
  if (!executedPathMatches) {
    checks.push({
      id: "executed-path",
      label: "Executed workbook path",
      status: "error",
      detail: "current configuration points to a different executed workbook",
    });
  }

  const exists = dependencies.fileExists ?? defaultExists;
  let sourceDrift: RecoverySourceDrift = "unavailable";
  if (!sourcePathMatches) {
    checks.push({
      id: "source",
      label: "Source workbook",
      status: "error",
      detail: "not inspected because the stored path differs from configuration",
    });
  } else if (!(await exists(storedSourcePath))) {
    checks.push({
      id: "source",
      label: "Source workbook",
      status: "error",
      detail: "missing",
    });
  } else {
    try {
      const [sourceHash, source] = await Promise.all([
        hashFile(storedSourcePath),
        loadPgnWorkbook(storedSourcePath),
      ]);
      const currentSourceScenarios = selectedScenarios(
        source.parsed.scenarios,
        state,
      );
      const currentManifest = createRecoveryManifest(
        state.runId, sourceHash, currentSourceScenarios, new Date(manifest.createdAt),
      );
      const schemaMatches = legacyRecoverySchemasMatch(source.workbook, manifest) && recoveryManifestMatches(manifest, currentManifest);
      if (sourceHash === state.sourceWorkbookHash && schemaMatches) {
        sourceDrift = "unchanged";
        checks.push({
          id: "source",
          label: "Source workbook",
          status: "ok",
          detail: "hash and scenario manifest unchanged",
        });
      } else {
        sourceDrift = schemaMatches
          ? "formatting-only"
          : "structural";
        checks.push({
          id: "source",
          label: "Source workbook",
          status: sourceDrift === "formatting-only" ? "warn" : "error",
          detail:
            sourceDrift === "formatting-only"
              ? "file hash changed, but selected scenario inputs are unchanged"
              : "selected scenario inputs, order, rows, turns, or column mapping changed; resume is blocked",
        });
      }
    } catch (error) {
      sourceDrift = "structural";
      checks.push({
        id: "source",
        label: "Source workbook",
        status: "error",
        detail: error instanceof Error ? error.message : "cannot be parsed",
      });
    }
  }

  let reconciliation: RecoveryReconciliation | undefined;
  if (!executedPathMatches) {
    checks.push({
      id: "workbook",
      label: "Executed workbook",
      status: "error",
      detail: "not inspected because the stored path differs from configuration",
    });
  } else if (!(await exists(storedExecutedPath))) {
    checks.push({
      id: "workbook",
      label: "Executed workbook",
      status: "error",
      detail: "missing",
    });
  } else {
    try {
      const executed = await loadPgnWorkbook(storedExecutedPath);
      const executionContext = getRunConfiguration(executed.workbook, state.runId);
      if (executionContext && (executionContext.sessionMode !== sessionMode || executionContext.transport !== readExecutionTransport(state.transport))) {
        throw new Error("Session mode or transport differs between workbook and recovery checkpoint");
      }
      const scenarios = selectedScenarios(executed.parsed.scenarios, state);
      if (!legacyRecoverySchemasMatch(executed.workbook, manifest)) {
        throw new Error("Older recovery checkpoint has no schema snapshot; changed column mapping requires explicit recovery review");
      }
      const executedManifest = createRecoveryManifest(
        state.runId,
        state.sourceWorkbookHash,
        scenarios,
        new Date(manifest.createdAt),
      );
      if (!recoveryManifestMatches(manifest, executedManifest)) {
        throw new Error(
          "Executed workbook scenario inputs, order, rows, turns, or column mapping differ from the recovery manifest",
        );
      }
      if (state.mode === "retest") {
        const retest = getRetestRunMetadata(executed.workbook, state.runId);
        if (!retest) {
          throw new Error("Executed workbook is missing the original retest selection");
        }
        if (
          JSON.stringify(retest.selectedIds) !==
          JSON.stringify(state.selectedScenarioIds)
        ) {
          throw new Error(
            "Retest selection differs between the workbook and recovery checkpoint",
          );
        }
      }
      const evidenceRun = getEvidenceRunMetadata(executed.workbook, state.runId);
      const folderIds = new Set(
        [
          state.driveRunFolderId,
          evidenceRun?.folderId || undefined,
          state.mode === "retest"
            ? getRetestRunMetadata(executed.workbook, state.runId)?.folderId
            : undefined,
        ].filter((value): value is string => Boolean(value)),
      );
      if (folderIds.size > 1) {
        throw new Error(
          "Google Drive run folder differs between recovery and workbook metadata",
        );
      }
      reconciliation = reconcileRecoveryArtifacts(
        executed.workbook,
        scenarios,
        state,
        manifest,
      );
      checks.push({
        id: "workbook",
        label: "Executed workbook",
        status: "ok",
        detail: "readable",
      });
      if (reconciliation.reconciledScenarioIds.length) {
        checks.push({
          id: "reconciliation-decisions",
          label: "Operator reconciliation",
          status: "warn",
          detail: `${reconciliation.reconciledScenarioIds.length} scenario(s) follow an explicit operator decision until successfully rerun or skipped`,
        });
      }
      checks.push({
        id: "reconciliation",
        label: "Progress reconciliation",
        status: reconciliation.mismatchedScenarioIds.length ? "error" : "ok",
        detail: reconciliation.mismatchedScenarioIds.length
          ? `${reconciliation.mismatchedScenarioIds.length} scenario(s) disagree across checkpoint, workbook, and transcript; choose a repair strategy before resume`
          : `${reconciliation.checkpointCompletedIds.length} completed, ${reconciliation.evidenceCapturedScenarioIds.length} with complete local evidence, ${reconciliation.evidenceUploadedScenarioIds.length} fully uploaded`,
      });
    } catch (error) {
      checks.push({
        id: "workbook",
        label: "Executed workbook",
        status: "error",
        detail: error instanceof Error ? error.message : "cannot be parsed",
      });
    }
  }

  if (state.isDemo) {
    checks.push(
      {
        id: "whatsapp-profile", label: "WhatsApp execution", status: "info",
        detail: "blocked in demo mode; no browser or Playwright launched",
      },
      {
        id: "drive", label: "Google Drive", status: "info",
        detail: "skipped in demo mode; no API calls",
      },
      {
        id: "discord", label: "Discord notifications", status: "info",
        detail: "suppressed in demo mode; no notifications sent",
      },
    );
  } else {
    const profileEntries = dependencies.profileEntries ?? readdir;
    try {
      const entries = await profileEntries(config.profileDir);
      checks.push({
        id: "whatsapp-profile",
        label: "WhatsApp session",
        status: entries.length ? "ok" : "error",
        detail: entries.length
          ? "saved browser profile is present; no testcase was sent"
          : "saved browser profile is empty",
      });
    } catch {
      checks.push({
        id: "whatsapp-profile",
        label: "WhatsApp session",
        status: "error",
        detail: "saved browser profile is missing",
      });
    }
    checks.push({
      id: "whatsapp-target",
      label: "WhatsApp target",
      status: config.target ? "ok" : "error",
      detail: config.target ? `configured by ${config.target.kind}` : "not configured",
    });

    if (state.mode === "retest" && !config.googleDriveEvidenceEnabled) {
      checks.push({
        id: "drive",
        label: "Google Drive",
        status: "error",
        detail: "retest recovery requires Drive evidence",
      });
    } else if (state.driveRunFolderId && !config.googleDriveEvidenceEnabled) {
      checks.push({
        id: "drive",
        label: "Google Drive",
        status: "error",
        detail: "the existing run folder cannot be reused while Drive evidence is disabled",
      });
    } else if (!config.googleDriveEvidenceEnabled) {
      checks.push({
        id: "drive",
        label: "Google Drive",
        status: "info",
        detail: "disabled for this full run",
      });
    } else if (dependencies.checkDriveAccess === false) {
      checks.push({
        id: "drive",
        label: "Google Drive",
        status: "info",
        detail: state.driveRunFolderId
          ? "stored run folder ID present; access not checked"
          : "configuration present; access not checked",
      });
    } else {
      try {
        const publisher =
          dependencies.drivePublisher ?? createGoogleDriveEvidencePublisher(config);
        await publisher.validateParentFolder();
        if (state.driveRunFolderId) {
          await publisher.validateRunFolder(
            state.driveRunFolderId,
            expectedRunFolderName(config, state),
          );
        }
        checks.push({
          id: "drive",
          label: "Google Drive",
          status: "ok",
          detail: state.driveRunFolderId
            ? "stored run folder is accessible and will be reused"
            : "parent folder is accessible; no run folder existed at interruption",
        });
      } catch (error) {
        checks.push({
          id: "drive",
          label: "Google Drive",
          status: "error",
          detail: safeGoogleCredentialError(error, config.googleServiceAccount?.value),
        });
      }
    }

    const discord = validateDiscordWebhookUrl(config.discordWebhookUrl);
    if (config.discordConfigurationIssues.length) {
      checks.push({
        id: "discord",
        label: "Discord",
        status: "warn",
        detail: "configuration has invalid settings; recovery remains available",
      });
    } else if (!config.discordNotificationsEnabled) {
      checks.push({
        id: "discord",
        label: "Discord",
        status: "info",
        detail: "notifications disabled",
      });
    } else {
      checks.push({
        id: "discord",
        label: "Discord",
        status: discord.valid ? "ok" : "warn",
        detail: discord.valid
          ? "configured; no notification sent by validation"
          : "enabled but webhook is unavailable; recovery remains fail-open",
      });
    }
  }

  return {
    runId: state.runId,
    mode: state.mode,
    state,
    manifest,
    lock,
    sourceDrift,
    reconciliation,
    checks,
    ready: !checks.some((check) => check.status === "error"),
    restartReady: sessionMode === "continuous" && !checks.some((check) =>
      check.status === "error" && check.id !== "session-continuity" && check.id !== "reconciliation",
    ),
  };
}

function lockDescription(lock: RunLockInspection): string {
  if (lock.status === "unlocked") return "no active process";
  if (lock.status === "active") return `active PID ${lock.record.pid}`;
  if (lock.status === "stale") return `stale PID ${lock.record.pid}`;
  return `unreadable (${lock.reason})`;
}

export function formatRecoveryDiscovery(discovery: RecoveryDiscovery): string {
  if (discovery.kind === "none") {
    return `Recoverable run: none\nProcess lock: ${lockDescription(discovery.lock)}`;
  }
  if (discovery.kind === "unreadable") {
    return [
      `Recovery state: unreadable${discovery.runId ? ` for ${discovery.runId}` : ""}`,
      `Reason: ${discovery.reason}`,
      `Process lock: ${lockDescription(discovery.lock)}`,
    ].join("\n");
  }
  const state = discovery.state;
  const finished = new Set([
    ...state.completedScenarioIds,
    ...state.skippedScenarioIds,
  ]).size;
  return [
    `Run ID: ${state.runId}`,
    ...(state.isDemo ? ["Recovery type: DEMO (local UI/testing only; live execution disabled)"] : []),
    `Mode: ${state.mode}`,
    "Transport: WhatsApp",
    `Session Mode: ${sessionModeLabel(readSessionMode(state.sessionMode))}`,
    ...(readSessionMode(state.sessionMode) === "continuous" ? [CONTINUOUS_RECOVERY_WARNING] : []),
    `State: ${discovery.kind === "running" ? "RUNNING" : state.status}`,
    `Progress: ${state.completedScenarioIds.length} completed, ${state.skippedScenarioIds.length} skipped, ${state.totalScenarios - finished} remaining`,
    `Last completed: ${state.lastCompletedScenarioId ?? "none"}`,
    `Interrupted scenario: ${state.activeScenarioId ?? "none"}`,
    `Interrupted attempt: ${state.activeScenarioAttempt ?? "n/a"}`,
    ...(state.isDemo && state.interruptionReason ? [state.interruptionReason] : []),
    ...(state.isDemo && state.activeScenarioId
      ? [`Recommended: Restart ${state.activeScenarioId} from Turn 1`]
      : []),
    `Workbook: ${state.executedWorkbookPath} (${state.workbookProgress})`,
    `Drive folder: ${state.driveRunFolderId ? "stored for reuse" : "not created"}`,
    `Process lock: ${lockDescription(discovery.lock)}`,
    `Last checkpoint: ${state.updatedAt}`,
  ].join("\n");
}

export function formatRecoveryValidation(validation: RecoveryValidation): string {
  const labels: Record<RecoveryCheckStatus, string> = {
    ok: "OK",
    warn: "WARN",
    error: "ERROR",
    info: "INFO",
  };
  const lines = [
    `Run ID: ${validation.runId}`,
    ...(validation.state.isDemo ? ["Recovery type: DEMO"] : []),
    `Mode: ${validation.mode}`,
    "Transport: WhatsApp",
    `Session Mode: ${sessionModeLabel(readSessionMode(validation.state.sessionMode))}`,
    ...validation.checks.map(
      (check) => `${labels[check.status].padEnd(5)} ${check.label}: ${check.detail}`,
    ),
  ];
  if (validation.state.isDemo && validation.reconciliation?.mismatchedScenarioIds.length) {
    lines.push(
      `Recovery state mismatch: ${validation.reconciliation.mismatchedScenarioIds.join(", ")}`,
      "Recommended: re-run uncertain scenarios from Turn 1 after reconciliation (preview only).",
    );
  }
  if (validation.reconciliation?.nextScenarioId && readSessionMode(validation.state.sessionMode) === "isolated") {
    lines.push(`Resume at: ${validation.reconciliation.nextScenarioId} from Turn 1`);
  }
  lines.push(`Resume readiness: ${validation.ready ? "READY" : "BLOCKED"}${validation.state.isDemo ? " (preview only)" : ""}`);
  if (readSessionMode(validation.state.sessionMode) === "continuous") {
    lines.push(`Full restart readiness: ${validation.restartReady ? "READY" : "BLOCKED"}; all ${validation.state.totalScenarios} original scenarios, new Run ID and clean initial reset`);
  }
  if (validation.state.isDemo) {
    lines.push("DEMO MODE: Real WhatsApp execution is disabled. No testcase messages were sent.");
  }
  return lines.join("\n");
}

async function appendRecoveryAudit(
  config: AppConfig,
  options: {
    runId: string;
    event:
      | "RUN_ABANDONED"
      | "RECOVERY_RECONCILED"
      | "SCENARIO_SKIPPED_BY_OPERATOR";
    message: string;
    scenarioId?: string;
    timestamp: Date;
  },
): Promise<string | undefined> {
  try {
    if (!(await defaultExists(config.pgnExecutedWorkbookPath))) {
      return "Recovery state was updated, but the missing executed workbook prevented an audit event from being appended.";
    }
    const opened = await openExecutedPgnWorkbook(
      config.pgnSourceWorkbookPath,
      config.pgnExecutedWorkbookPath,
    );
    const scenario = options.scenarioId
      ? opened.parsed.scenarios.find(
          (item) => item.testCaseId === options.scenarioId,
        )
      : undefined;
    appendRecoveryTranscriptEvent(opened.workbook, {
      runId: options.runId,
      event: options.event,
      message: options.message,
      timestamp: options.timestamp,
      scenario,
    });
    await saveExecutedPgnWorkbook(
      opened.workbook,
      config.pgnExecutedWorkbookPath,
    );
    return undefined;
  } catch (error) {
    return `Recovery state was updated, but the workbook audit event could not be saved: ${error instanceof Error ? error.message : String(error)}`;
  }
}

async function withRecoveryOwnership<Value>(
  config: AppConfig,
  runId: string,
  purpose: string,
  operation: (config: AppConfig) => Promise<Value>,
): Promise<Value> {
  const { state } = await readRecoveryRun(config.projectRoot, runId);
  config = await recoveryDemoConfig(config, state);
  const runLock = await acquireRunProcessLock(config.projectRoot, purpose, {
    recoverStale: !state.isDemo,
  });
  let workbookRelease: (() => Promise<void>) | undefined;
  try {
    if (state.isDemo) {
      const discovery = await discoverRecoveryRun(config.projectRoot);
      if (
        discovery.kind === "unreadable" ||
        (discovery.kind !== "none" && discovery.state.runId !== runId)
      ) {
        throw new Error("Another recovery run exists; demo operations will not alter it");
      }
    }
    await runLock.heartbeat({ runId, mode: (await readRecoveryRun(config.projectRoot, runId)).state.mode });
    workbookRelease = await acquireWorkbookLock(
      config.pgnExecutedWorkbookPath,
      purpose,
    );
    return await operation(config);
  } finally {
    await workbookRelease?.().catch(() => undefined);
    await runLock.release().catch(() => undefined);
  }
}

function updateAttempt(
  state: RecoveryRunState,
  scenarioId: string,
  status: "COMPLETED" | "INTERRUPTED" | "SKIPPED_BY_OPERATOR",
  now: Date,
  reason: string,
): void {
  const attempt = [...state.scenarioAttempts]
    .reverse()
    .find(
      (item) => item.scenarioId === scenarioId && item.status === "RUNNING",
    );
  if (attempt) {
    attempt.status = status;
    attempt.finishedAt = now.toISOString();
    attempt.reason = reason;
  }
}

export async function skipRecoveryScenario(
  config: AppConfig,
  runId: string,
): Promise<RecoveryMutationResult> {
  if (readSessionMode((await readRecoveryRun(config.projectRoot, runId)).state.sessionMode) === "continuous") {
    throw new Error(CONTINUOUS_RECOVERY_WARNING);
  }
  return withRecoveryOwnership(
    config,
    runId,
    "PGN recovery scenario skip",
    async (config) => {
      const { checkpoint } = await openRecoveryCheckpoint(config.projectRoot, runId);
      const before = checkpoint.snapshot();
      const finished = new Set([
        ...before.completedScenarioIds,
        ...before.skippedScenarioIds,
      ]);
      const scenarioId =
        before.activeScenarioId ??
        before.selectedScenarioIds.find((id) => !finished.has(id));
      if (!scenarioId) throw new Error("No incomplete recovery scenario is available to skip");
      const now = new Date();
      const warning = await appendRecoveryAudit(config, {
        runId,
        event: "SCENARIO_SKIPPED_BY_OPERATOR",
        message: `Operator skipped ${scenarioId}; the run will continue with the next incomplete scenario.`,
        scenarioId,
        timestamp: now,
      });
      await checkpoint.update((state) => {
        if (!state.skippedScenarioIds.includes(scenarioId)) {
          state.skippedScenarioIds.push(scenarioId);
        }
        state.reconciliationDecisions = state.reconciliationDecisions.filter(
          (decision) => decision.scenarioId !== scenarioId,
        );
        updateAttempt(
          state,
          scenarioId,
          "SKIPPED_BY_OPERATOR",
          now,
          "Explicitly skipped by operator",
        );
        state.activeScenarioId = undefined;
        state.activeScenarioAttempt = undefined;
        state.activeScenarioStartedAt = undefined;
        state.status = "RECOVERABLE";
        state.interruptionReason = `Scenario ${scenarioId} skipped by operator`;
        state.updatedAt = now.toISOString();
        state.heartbeatAt = now.toISOString();
      });
      return { runId, scenarioId, warning };
    },
  );
}

export async function repairRecoveryProgress(
  config: AppConfig,
  runId: string,
  strategy: RecoveryRepairStrategy,
): Promise<RecoveryMutationResult> {
  if (readSessionMode((await readRecoveryRun(config.projectRoot, runId)).state.sessionMode) === "continuous") {
    throw new Error(CONTINUOUS_RECOVERY_WARNING);
  }
  return withRecoveryOwnership(
    config,
    runId,
    "PGN recovery reconciliation",
    async (config) => {
      const { checkpoint, manifest } = await openRecoveryCheckpoint(
        config.projectRoot,
        runId,
      );
      const before = checkpoint.snapshot();
      const executed = await openExecutedPgnWorkbook(
        config.pgnSourceWorkbookPath,
        config.pgnExecutedWorkbookPath,
      );
      const scenarios = selectedScenarios(executed.parsed.scenarios, before);
      const reconciliation = reconcileRecoveryArtifacts(
        executed.workbook,
        scenarios,
        before,
        manifest,
      );
      const completed =
        strategy === "checkpoint"
          ? reconciliation.checkpointCompletedIds
          : strategy === "artifacts"
            ? reconciliation.artifactConfirmedIds
            : reconciliation.safeCompletedIds;
      const completedSet = new Set(completed);
      const mismatchSet = new Set(reconciliation.mismatchedScenarioIds);
      const now = new Date();
      const descriptions: Record<RecoveryRepairStrategy, string> = {
        rerun: "re-run every mismatched scenario (safest)",
        checkpoint: "trust checkpoint completion records",
        artifacts: "trust scenarios confirmed by both workbook and transcript",
      };
      if (before.mode === "retest") {
        const retest = getRetestRunMetadata(executed.workbook, runId);
        if (!retest) {
          throw new Error(`Retest Run was not found in the workbook: ${runId}`);
        }
        retest.finishedIds = retest.selectedIds.filter((id) =>
          completedSet.has(id),
        );
        retest.state = "IN_PROGRESS";
        retest.updatedAt = now;
        upsertRetestRunMetadata(executed.workbook, retest);
      }
      appendRecoveryTranscriptEvent(executed.workbook, {
        runId,
        event: "RECOVERY_RECONCILED",
        message: `Operator reconciled recovery progress: ${descriptions[strategy]}.`,
        timestamp: now,
      });
      await saveExecutedPgnWorkbook(
        executed.workbook,
        config.pgnExecutedWorkbookPath,
      );
      await checkpoint.update((state) => {
        state.completedScenarioIds = state.selectedScenarioIds.filter(
          (id) => completedSet.has(id) && !state.skippedScenarioIds.includes(id),
        );
        state.lastCompletedScenarioId = state.completedScenarioIds.at(-1);
        state.reconciliationDecisions = [
          ...state.reconciliationDecisions.filter(
            (decision) => !mismatchSet.has(decision.scenarioId),
          ),
          ...reconciliation.mismatchedScenarioIds.map((scenarioId) => ({
            scenarioId,
            strategy,
            decidedAt: now.toISOString(),
          })),
        ];
        if (state.activeScenarioId && completedSet.has(state.activeScenarioId)) {
          updateAttempt(
            state,
            state.activeScenarioId,
            "COMPLETED",
            now,
            `Reconciled using ${strategy} strategy`,
          );
          state.activeScenarioId = undefined;
          state.activeScenarioAttempt = undefined;
          state.activeScenarioStartedAt = undefined;
        }
        state.status = "RECOVERABLE";
        state.interruptionReason = `Progress reconciled: ${descriptions[strategy]}`;
        state.updatedAt = now.toISOString();
        state.heartbeatAt = now.toISOString();
      });
      return { runId };
    },
  );
}

export async function abandonRecoveryRun(
  config: AppConfig,
  runId: string,
): Promise<RecoveryMutationResult> {
  return withRecoveryOwnership(
    config,
    runId,
    "PGN recovery abandonment",
    async (config) => {
      const { checkpoint } = await openRecoveryCheckpoint(config.projectRoot, runId);
      const before = checkpoint.snapshot();
      if (before.status === "COMPLETED" || before.status === "ABANDONED") {
        throw new Error(`Run ${runId} is already ${before.status}`);
      }
      const now = new Date();
      const warning = await appendRecoveryAudit(config, {
        runId,
        event: "RUN_ABANDONED",
        message: "Operator explicitly abandoned recovery. Existing workbook, evidence, and run history were preserved.",
        scenarioId: before.activeScenarioId,
        timestamp: now,
      });
      await checkpoint.update((state) => {
        if (state.activeScenarioId) {
          updateAttempt(
            state,
            state.activeScenarioId,
            "INTERRUPTED",
            now,
            "Run abandoned by operator",
          );
        }
        state.status = "ABANDONED";
        state.interruptedAt = now.toISOString();
        state.interruptionReason = "Explicitly abandoned by operator";
        state.activeScenarioId = undefined;
        state.activeScenarioAttempt = undefined;
        state.activeScenarioStartedAt = undefined;
        state.updatedAt = now.toISOString();
        state.heartbeatAt = now.toISOString();
      });
      return { runId, warning };
    },
  );
}
