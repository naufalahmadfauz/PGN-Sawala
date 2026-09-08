import { createHash, randomUUID } from "node:crypto";
import {
  mkdir,
  open,
  readFile,
  readdir,
  rename,
  rm,
  rmdir,
  stat,
  unlink,
} from "node:fs/promises";
import path from "node:path";
import type { PgnTestScenario } from "../excel/pgn-types";
import { isProcessAlive } from "../process-liveness";
import { readExecutionTransport, readSessionMode, type RunExecutionContext, type SessionMode, type ExecutionTransport } from "../session-mode";

export const RECOVERY_SCHEMA_VERSION = 1;
export const RUN_LOCK_STALE_MS = 45_000;

export type RecoveryRunStatus =
  | "PREPARING"
  | "RUNNING"
  | "INTERRUPTED"
  | "RECOVERABLE"
  | "COMPLETED"
  | "FAILED"
  | "ABANDONED";

export type RecoveryAttemptStatus =
  | "RUNNING"
  | "INTERRUPTED"
  | "COMPLETED"
  | "FAILED"
  | "SKIPPED_BY_OPERATOR";

export interface RecoveryScenarioAttempt {
  scenarioId: string;
  attempt: number;
  status: RecoveryAttemptStatus;
  startedAt: string;
  finishedAt?: string;
  reason?: string;
}

export interface RecoveryRunMetrics {
  executedScenarios: number;
  capturedScenarios: number;
  timeouts: number;
  technicalErrors: number;
  evidenceCaptured: number;
  evidenceUploaded: number;
  evidenceUploadErrors: number;
}

export interface RecoveryReconciliationDecision {
  scenarioId: string;
  strategy: "rerun" | "checkpoint" | "artifacts";
  decidedAt: string;
}

export interface RecoveryRunState {
  schemaVersion: typeof RECOVERY_SCHEMA_VERSION;
  runId: string;
  isDemo?: true;
  mode: "full" | "retest";
  sessionMode?: SessionMode;
  transport?: ExecutionTransport;
  sessionResetAttempts?: number;
  restartedFromRunId?: string;
  restTarget?: { accountId: string; skillId: string };
  status: RecoveryRunStatus;
  sourceWorkbookPath: string;
  sourceWorkbookHash: string;
  executedWorkbookPath: string;
  selectedScenarioIds: string[];
  completedScenarioIds: string[];
  skippedScenarioIds: string[];
  totalScenarios: number;
  lastCompletedScenarioId?: string;
  activeScenarioId?: string;
  activeScenarioAttempt?: number;
  activeScenarioStartedAt?: string;
  scenarioAttempts: RecoveryScenarioAttempt[];
  reconciliationDecisions: RecoveryReconciliationDecision[];
  driveRunFolderId?: string;
  driveRunFolderUrl?: string;
  finalCleanupComplete: boolean;
  workbookProgress: string;
  metrics: RecoveryRunMetrics;
  startedAt: string;
  updatedAt: string;
  heartbeatAt: string;
  interruptedAt?: string;
  interruptionReason?: string;
  resumedAt?: string;
  resumeCount: number;
}

export interface RecoveryManifestScenario {
  testCaseId: string;
  sheetKind: "kb" | "negative";
  sheetName: string;
  sourceRowNumber: number;
  order: number;
  turnCount: number;
  inputHash: string;
  schemaFingerprint?: string;
}

export interface RecoveryRunManifest {
  schemaVersion: typeof RECOVERY_SCHEMA_VERSION;
  runId: string;
  isDemo?: true;
  sessionMode?: SessionMode;
  transport?: ExecutionTransport;
  sourceWorkbookHash: string;
  createdAt: string;
  scenarios: RecoveryManifestScenario[];
}

interface ActiveRunPointer {
  schemaVersion: typeof RECOVERY_SCHEMA_VERSION;
  runId: string;
  updatedAt: string;
}

export interface RunLockRecord {
  schemaVersion: typeof RECOVERY_SCHEMA_VERSION;
  token: string;
  pid: number;
  purpose: string;
  createdAt: string;
  heartbeatAt: string;
  runId?: string;
  mode?: "full" | "retest";
}

export type RunLockInspection =
  | { status: "unlocked" }
  | { status: "active"; record: RunLockRecord; ageMs: number }
  | { status: "stale"; record: RunLockRecord; ageMs: number }
  | { status: "unreadable"; reason: string; ageMs?: number };

export interface RunProcessLock {
  readonly path: string;
  readonly token: string;
  heartbeat(run?: { runId: string; mode: "full" | "retest" }): Promise<void>;
  release(): Promise<void>;
}

export type RecoveryDiscovery =
  | { kind: "none"; lock: RunLockInspection }
  | {
      kind: "running" | "recoverable";
      state: RecoveryRunState;
      manifest: RecoveryRunManifest;
      lock: RunLockInspection;
    }
  | {
      kind: "unreadable";
      runId?: string;
      reason: string;
      lock: RunLockInspection;
    };

export interface AtomicWriteHooks {
  beforeRename?: (temporaryPath: string, targetPath: string) => Promise<void>;
}

function runtimeRoot(projectRoot: string): string {
  return path.join(projectRoot, ".runtime", "pgn");
}

export function recoveryPaths(projectRoot: string, runId?: string): {
  root: string;
  runs: string;
  active: string;
  lock: string;
  runDirectory?: string;
  state?: string;
  manifest?: string;
} {
  const root = runtimeRoot(projectRoot);
  const runs = path.join(root, "runs");
  if (!runId) {
    return {
      root,
      runs,
      active: path.join(root, "active-run.json"),
      lock: path.join(root, "run.lock"),
    };
  }
  assertSafeRunId(runId);
  const runDirectory = path.join(runs, runId);
  return {
    root,
    runs,
    active: path.join(root, "active-run.json"),
    lock: path.join(root, "run.lock"),
    runDirectory,
    state: path.join(runDirectory, "state.json"),
    manifest: path.join(runDirectory, "manifest.json"),
  };
}

function assertSafeRunId(runId: string): void {
  if (!/^[A-Za-z0-9][A-Za-z0-9._-]{0,159}$/.test(runId)) {
    throw new Error(`Unsafe recovery Run ID: ${runId}`);
  }
}

export function assertRecoveryRunExecutable(
  state: Pick<RecoveryRunState, "runId" | "isDemo">,
): void {
  if (state.isDemo || state.runId.startsWith("DEMO-RECOVERY-")) {
    throw new Error(
      "Recovery demo detected. This run is for UI/testing purposes only. Real WhatsApp execution is disabled.",
    );
  }
}

function assertDemoIdentity(runId: string, isDemo: unknown): void {
  if (
    (isDemo !== undefined && isDemo !== true) ||
    runId.startsWith("DEMO-RECOVERY-") !== (isDemo === true)
  ) {
    throw new Error("Recovery demo identity is invalid; demo runs cannot become real runs");
  }
}

async function syncDirectory(directory: string): Promise<void> {
  let handle;
  try {
    handle = await open(directory, "r");
    await handle.sync();
  } catch {
    // Directory fsync is not supported consistently on Windows.
  } finally {
    await handle?.close().catch(() => undefined);
  }
}

export async function atomicWriteJson(
  targetPath: string,
  value: unknown,
  hooks: AtomicWriteHooks = {},
): Promise<void> {
  await mkdir(path.dirname(targetPath), { recursive: true });
  const temporaryPath = path.join(
    path.dirname(targetPath),
    `.${path.basename(targetPath)}.${process.pid}.${randomUUID()}.tmp`,
  );
  let handle;
  try {
    handle = await open(temporaryPath, "wx", 0o600);
    await handle.writeFile(`${JSON.stringify(value, null, 2)}\n`, "utf8");
    await handle.sync();
    await handle.close();
    handle = undefined;
    await hooks.beforeRename?.(temporaryPath, targetPath);
    await rename(temporaryPath, targetPath);
    await syncDirectory(path.dirname(targetPath));
  } catch (error) {
    await handle?.close().catch(() => undefined);
    await rm(temporaryPath, { force: true }).catch(() => undefined);
    throw error;
  }
}

export async function hashFile(filePath: string): Promise<string> {
  return createHash("sha256").update(await readFile(filePath)).digest("hex");
}

function scenarioInputHash(scenario: PgnTestScenario): string {
  return createHash("sha256")
    .update(
      JSON.stringify({
        testCaseId: scenario.testCaseId,
        sheetKind: scenario.sheetKind,
        sheetName: scenario.sheetName,
        sourceRowNumber: scenario.sourceRowNumber,
        turns: scenario.turns.map((turn) => ({
          rowNumber: turn.rowNumber,
          turnNumber: turn.turnNumber,
          userInput: turn.userInput,
        })),
      }),
    )
    .digest("hex");
}

export function createRecoveryManifest(
  runId: string,
  sourceWorkbookHash: string,
  scenarios: readonly PgnTestScenario[],
  createdAt = new Date(),
  execution: Partial<RunExecutionContext> = {},
): RecoveryRunManifest {
  assertSafeRunId(runId);
  return {
    schemaVersion: RECOVERY_SCHEMA_VERSION,
    runId,
    sessionMode: readSessionMode(execution.sessionMode),
    transport: readExecutionTransport(execution.transport),
    sourceWorkbookHash,
    createdAt: createdAt.toISOString(),
    scenarios: scenarios.map((scenario, order) => ({
      testCaseId: scenario.testCaseId,
      sheetKind: scenario.sheetKind,
      sheetName: scenario.sheetName,
      sourceRowNumber: scenario.sourceRowNumber,
      order,
      turnCount: scenario.turns.length,
      inputHash: scenarioInputHash(scenario),
      ...(scenario.schemaFingerprint ? { schemaFingerprint: scenario.schemaFingerprint } : {}),
    })),
  };
}

export function recoveryManifestMatches(
  expected: RecoveryRunManifest,
  current: RecoveryRunManifest,
): boolean {
  return expected.scenarios.length === current.scenarios.length && expected.scenarios.every((scenario, index) => {
    const actual = current.scenarios[index];
    const { schemaFingerprint, ...identity } = scenario;
    const { schemaFingerprint: currentSchema, ...currentIdentity } = actual;
    return JSON.stringify(identity) === JSON.stringify(currentIdentity) &&
      (schemaFingerprint === undefined || schemaFingerprint === currentSchema);
  });
}

function stringArray(value: unknown, name: string): string[] {
  if (!Array.isArray(value) || !value.every((item) => typeof item === "string")) {
    throw new Error(`${name} must be an array of strings`);
  }
  if (new Set(value).size !== value.length) {
    throw new Error(`${name} must not contain duplicates`);
  }
  return value;
}

function validDate(value: unknown, name: string): string {
  if (typeof value !== "string" || Number.isNaN(Date.parse(value))) {
    throw new Error(`${name} must be an ISO date`);
  }
  return value;
}

function parseRecoveryState(value: unknown): RecoveryRunState {
  if (!value || typeof value !== "object") {
    throw new Error("Recovery state must be an object");
  }
  const candidate = value as Partial<RecoveryRunState>;
  const statuses: RecoveryRunStatus[] = [
    "PREPARING",
    "RUNNING",
    "INTERRUPTED",
    "RECOVERABLE",
    "COMPLETED",
    "FAILED",
    "ABANDONED",
  ];
  if (candidate.schemaVersion !== RECOVERY_SCHEMA_VERSION) {
    throw new Error("Recovery state schema version is not supported");
  }
  if (typeof candidate.runId !== "string") throw new Error("Run ID is missing");
  assertSafeRunId(candidate.runId);
  assertDemoIdentity(candidate.runId, candidate.isDemo);
  readSessionMode(candidate.sessionMode);
  readExecutionTransport(candidate.transport);
  if (candidate.sessionResetAttempts !== undefined && (!Number.isInteger(candidate.sessionResetAttempts) || candidate.sessionResetAttempts < 0)) {
    throw new Error("Recovery session reset count is invalid");
  }
  if (candidate.restartedFromRunId !== undefined) {
    assertSafeRunId(candidate.restartedFromRunId);
    if (candidate.restartedFromRunId === candidate.runId) throw new Error("A restarted run must have a new Run ID");
  }
  if (candidate.restTarget !== undefined && (
    !candidate.restTarget || typeof candidate.restTarget.accountId !== "string" ||
    typeof candidate.restTarget.skillId !== "string" || !/^[A-Za-z0-9_-]+$/.test(candidate.restTarget.accountId) || !/^\d+$/.test(candidate.restTarget.skillId)
  )) throw new Error("Recovery REST target is invalid");
  if (candidate.mode !== "full" && candidate.mode !== "retest") {
    throw new Error("Recovery mode is invalid");
  }
  if (!candidate.status || !statuses.includes(candidate.status)) {
    throw new Error("Recovery status is invalid");
  }
  if (
    typeof candidate.sourceWorkbookPath !== "string" ||
    !candidate.sourceWorkbookPath ||
    typeof candidate.sourceWorkbookHash !== "string" ||
    !/^[a-f0-9]{64}$/.test(candidate.sourceWorkbookHash) ||
    typeof candidate.executedWorkbookPath !== "string" ||
    !candidate.executedWorkbookPath ||
    typeof candidate.workbookProgress !== "string"
  ) {
    throw new Error("Recovery workbook metadata is invalid");
  }
  const selectedScenarioIds = stringArray(
    candidate.selectedScenarioIds,
    "selectedScenarioIds",
  );
  const completedScenarioIds = stringArray(
    candidate.completedScenarioIds,
    "completedScenarioIds",
  );
  const skippedScenarioIds = stringArray(
    candidate.skippedScenarioIds,
    "skippedScenarioIds",
  );
  if (
    !Number.isInteger(candidate.totalScenarios) ||
    candidate.totalScenarios !== selectedScenarioIds.length
  ) {
    throw new Error("Recovery total scenario count is invalid");
  }
  if (!Array.isArray(candidate.scenarioAttempts)) {
    throw new Error("Recovery scenario attempts are invalid");
  }
  if (!Array.isArray(candidate.reconciliationDecisions)) {
    throw new Error("Recovery reconciliation decisions are invalid");
  }
  const selectedIds = new Set(selectedScenarioIds);
  const completedIds = new Set(completedScenarioIds);
  const skippedIds = new Set(skippedScenarioIds);
  if (
    completedScenarioIds.some((id) => !selectedIds.has(id)) ||
    skippedScenarioIds.some((id) => !selectedIds.has(id)) ||
    completedScenarioIds.some((id) => skippedIds.has(id))
  ) {
    throw new Error("Recovery progress contains unknown or conflicting scenario IDs");
  }
  if (
    candidate.activeScenarioId !== undefined &&
    (!selectedIds.has(candidate.activeScenarioId) ||
      completedIds.has(candidate.activeScenarioId) ||
      skippedIds.has(candidate.activeScenarioId))
  ) {
    throw new Error("Recovery active scenario is invalid");
  }
  const attemptStatuses: RecoveryAttemptStatus[] = [
    "RUNNING",
    "INTERRUPTED",
    "COMPLETED",
    "FAILED",
    "SKIPPED_BY_OPERATOR",
  ];
  const attemptKeys = new Set<string>();
  for (const attempt of candidate.scenarioAttempts) {
    if (!attempt || typeof attempt !== "object") {
      throw new Error("Recovery scenario attempt is invalid");
    }
    const parsed = attempt as RecoveryScenarioAttempt;
    const attemptKey = `${parsed.scenarioId}|${parsed.attempt}`;
    if (
      typeof parsed.scenarioId !== "string" ||
      !selectedIds.has(parsed.scenarioId) ||
      !Number.isInteger(parsed.attempt) ||
      parsed.attempt < 1 ||
      !attemptStatuses.includes(parsed.status) ||
      attemptKeys.has(attemptKey)
    ) {
      throw new Error("Recovery scenario attempt is invalid");
    }
    validDate(parsed.startedAt, "attempt startedAt");
    if (parsed.finishedAt !== undefined) {
      validDate(parsed.finishedAt, "attempt finishedAt");
    }
    attemptKeys.add(attemptKey);
  }
  const decisionIds = new Set<string>();
  for (const decision of candidate.reconciliationDecisions) {
    const parsed = decision as RecoveryReconciliationDecision;
    if (
      !decision ||
      typeof decision !== "object" ||
      typeof parsed.scenarioId !== "string" ||
      !selectedIds.has(parsed.scenarioId) ||
      (parsed.strategy !== "rerun" &&
        parsed.strategy !== "checkpoint" &&
        parsed.strategy !== "artifacts") ||
      decisionIds.has(parsed.scenarioId)
    ) {
      throw new Error("Recovery reconciliation decision is invalid");
    }
    validDate(parsed.decidedAt, "reconciliation decidedAt");
    decisionIds.add(parsed.scenarioId);
  }
  const metrics = candidate.metrics;
  if (!metrics || typeof metrics !== "object") {
    throw new Error("Recovery metrics are invalid");
  }
  const metricNames: Array<keyof RecoveryRunMetrics> = [
    "executedScenarios",
    "capturedScenarios",
    "timeouts",
    "technicalErrors",
    "evidenceCaptured",
    "evidenceUploaded",
    "evidenceUploadErrors",
  ];
  for (const metric of metricNames.map((name) => metrics[name])) {
    if (!Number.isInteger(metric) || metric < 0) {
      throw new Error("Recovery metrics must be non-negative integers");
    }
  }
  if (
    typeof candidate.finalCleanupComplete !== "boolean" ||
    !Number.isInteger(candidate.resumeCount) ||
    (candidate.resumeCount ?? -1) < 0
  ) {
    throw new Error("Recovery lifecycle metadata is invalid");
  }
  for (const [name, value] of [
    ["activeScenarioStartedAt", candidate.activeScenarioStartedAt],
    ["interruptedAt", candidate.interruptedAt],
    ["resumedAt", candidate.resumedAt],
  ] as const) {
    if (value !== undefined) validDate(value, name);
  }
  if (
    candidate.activeScenarioAttempt !== undefined &&
    (!Number.isInteger(candidate.activeScenarioAttempt) ||
      candidate.activeScenarioAttempt < 1)
  ) {
    throw new Error("Recovery active scenario attempt is invalid");
  }
  const hasAnyActiveField =
    candidate.activeScenarioId !== undefined ||
    candidate.activeScenarioAttempt !== undefined ||
    candidate.activeScenarioStartedAt !== undefined;
  if (
    hasAnyActiveField &&
    (!candidate.activeScenarioId ||
      candidate.activeScenarioAttempt === undefined ||
      candidate.activeScenarioStartedAt === undefined ||
      !attemptKeys.has(
        `${candidate.activeScenarioId}|${candidate.activeScenarioAttempt}`,
      ))
  ) {
    throw new Error("Recovery active scenario fields do not agree");
  }
  if (
    candidate.lastCompletedScenarioId !== undefined &&
    !completedIds.has(candidate.lastCompletedScenarioId)
  ) {
    throw new Error("Recovery last completed scenario is invalid");
  }
  return {
    ...(candidate as RecoveryRunState),
    selectedScenarioIds,
    completedScenarioIds,
    skippedScenarioIds,
    startedAt: validDate(candidate.startedAt, "startedAt"),
    updatedAt: validDate(candidate.updatedAt, "updatedAt"),
    heartbeatAt: validDate(candidate.heartbeatAt, "heartbeatAt"),
  };
}

function parseManifest(value: unknown): RecoveryRunManifest {
  if (!value || typeof value !== "object") {
    throw new Error("Recovery manifest must be an object");
  }
  const candidate = value as Partial<RecoveryRunManifest>;
  if (
    candidate.schemaVersion !== RECOVERY_SCHEMA_VERSION ||
    typeof candidate.runId !== "string" ||
    typeof candidate.sourceWorkbookHash !== "string" ||
    !Array.isArray(candidate.scenarios)
  ) {
    throw new Error("Recovery manifest is invalid");
  }
  assertSafeRunId(candidate.runId);
  assertDemoIdentity(candidate.runId, candidate.isDemo);
  readSessionMode(candidate.sessionMode);
  readExecutionTransport(candidate.transport);
  validDate(candidate.createdAt, "manifest createdAt");
  for (const scenario of candidate.scenarios) {
    const parsed = scenario as RecoveryManifestScenario;
    if (
      !scenario ||
      typeof scenario !== "object" ||
      typeof parsed.testCaseId !== "string" ||
      (parsed.sheetKind !== "kb" && parsed.sheetKind !== "negative") ||
      typeof parsed.sheetName !== "string" ||
      !Number.isInteger(parsed.sourceRowNumber) ||
      parsed.sourceRowNumber < 2 ||
      !Number.isInteger(parsed.order) ||
      parsed.order < 0 ||
      !Number.isInteger(parsed.turnCount) ||
      parsed.turnCount < 1 ||
      !/^[a-f0-9]{64}$/.test(parsed.inputHash)
      || (parsed.schemaFingerprint !== undefined && !/^[a-f0-9]{64}$/.test(parsed.schemaFingerprint))
    ) {
      throw new Error("Recovery manifest scenario is invalid");
    }
  }
  return candidate as RecoveryRunManifest;
}

async function readJson(filePath: string): Promise<unknown> {
  return JSON.parse(await readFile(filePath, "utf8")) as unknown;
}

export async function readRecoveryRun(
  projectRoot: string,
  runId: string,
): Promise<{ state: RecoveryRunState; manifest: RecoveryRunManifest }> {
  const paths = recoveryPaths(projectRoot, runId);
  const [state, manifest] = await Promise.all([
    readJson(paths.state!),
    readJson(paths.manifest!),
  ]);
  const parsedState = parseRecoveryState(state);
  const parsedManifest = parseManifest(manifest);
  if (parsedState.runId !== runId || parsedManifest.runId !== runId) {
    throw new Error("Recovery files do not match the requested Run ID");
  }
  const manifestIds = parsedManifest.scenarios.map((scenario, index) => {
    if (scenario.order !== index) {
      throw new Error("Recovery manifest scenario order is invalid");
    }
    return scenario.testCaseId;
  });
  if (
    new Set(manifestIds).size !== manifestIds.length ||
    JSON.stringify(manifestIds) !==
      JSON.stringify(parsedState.selectedScenarioIds) ||
    parsedManifest.sourceWorkbookHash !== parsedState.sourceWorkbookHash ||
    parsedManifest.isDemo !== parsedState.isDemo ||
    readSessionMode(parsedManifest.sessionMode) !== readSessionMode(parsedState.sessionMode) ||
    readExecutionTransport(parsedManifest.transport) !== readExecutionTransport(parsedState.transport)
  ) {
    throw new Error("Recovery state and manifest do not agree");
  }
  return { state: parsedState, manifest: parsedManifest };
}

async function activeRunId(projectRoot: string): Promise<string | undefined> {
  try {
    const value = (await readJson(recoveryPaths(projectRoot).active)) as Partial<ActiveRunPointer>;
    if (
      value.schemaVersion !== RECOVERY_SCHEMA_VERSION ||
      typeof value.runId !== "string"
    ) {
      throw new Error("Active run pointer is invalid");
    }
    assertSafeRunId(value.runId);
    return value.runId;
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code === "ENOENT") return undefined;
    throw error;
  }
}

function isRecoverableStatus(status: RecoveryRunStatus): boolean {
  return status !== "COMPLETED" && status !== "ABANDONED";
}

async function newestRecoverableRunId(projectRoot: string): Promise<string | undefined> {
  const paths = recoveryPaths(projectRoot);
  let entries;
  try {
    entries = await readdir(paths.runs, { withFileTypes: true });
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code === "ENOENT") return undefined;
    throw error;
  }
  const candidates: RecoveryRunState[] = [];
  for (const entry of entries) {
    if (!entry.isDirectory()) continue;
    try {
      const state = parseRecoveryState(
        await readJson(path.join(paths.runs, entry.name, "state.json")),
      );
      if (isRecoverableStatus(state.status)) candidates.push(state);
    } catch {
      // Historical unreadable entries are ignored unless the active pointer selects one.
    }
  }
  candidates.sort((left, right) => right.updatedAt.localeCompare(left.updatedAt));
  return candidates[0]?.runId;
}

function parseLockRecord(value: unknown): RunLockRecord {
  if (!value || typeof value !== "object") throw new Error("lock record is invalid");
  const record = value as Partial<RunLockRecord>;
  if (
    record.schemaVersion !== RECOVERY_SCHEMA_VERSION ||
    typeof record.token !== "string" ||
    !record.token ||
    !Number.isInteger(record.pid) ||
    (record.pid ?? 0) <= 0 ||
    typeof record.purpose !== "string"
  ) {
    throw new Error("lock record is invalid");
  }
  validDate(record.createdAt, "lock createdAt");
  validDate(record.heartbeatAt, "lock heartbeatAt");
  if (record.runId !== undefined) assertSafeRunId(record.runId);
  if (
    record.mode !== undefined &&
    record.mode !== "full" &&
    record.mode !== "retest"
  ) {
    throw new Error("lock mode is invalid");
  }
  return record as RunLockRecord;
}

async function pathAge(filePath: string, now: Date): Promise<number | undefined> {
  try {
    return Math.max(0, now.getTime() - (await stat(filePath)).mtimeMs);
  } catch {
    return undefined;
  }
}

export async function inspectRunProcessLock(
  projectRoot: string,
  options: {
    now?: Date;
    processAlive?: (pid: number) => boolean;
    staleAfterMs?: number;
  } = {},
): Promise<RunLockInspection> {
  const now = options.now ?? new Date();
  const lockPath = recoveryPaths(projectRoot).lock;
  let entries: string[];
  try {
    entries = await readdir(lockPath);
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code === "ENOENT") {
      return { status: "unlocked" };
    }
    return {
      status: "unreadable",
      reason: "Run lock cannot be read",
      ageMs: await pathAge(lockPath, now),
    };
  }
  const ownerNames = entries.filter(
    (entry) => entry.startsWith("owner-") && entry.endsWith(".json"),
  );
  if (ownerNames.length !== 1) {
    return {
      status: "unreadable",
      reason:
        ownerNames.length === 0
          ? "Run lock has no owner record"
          : "Run lock has multiple owner records",
      ageMs: await pathAge(lockPath, now),
    };
  }
  const ownerName = ownerNames[0];
  try {
    const owner = parseLockRecord(await readJson(path.join(lockPath, ownerName)));
    if (ownerName !== `owner-${owner.token}.json`) {
      throw new Error("lock token does not match owner file");
    }
    const heartbeatPath = path.join(lockPath, `heartbeat-${owner.token}.json`);
    let record = owner;
    try {
      const heartbeat = parseLockRecord(await readJson(heartbeatPath));
      if (heartbeat.token !== owner.token || heartbeat.pid !== owner.pid) {
        throw new Error("lock heartbeat owner does not match");
      }
      record = heartbeat;
    } catch (error) {
      if ((error as NodeJS.ErrnoException).code !== "ENOENT") throw error;
    }
    const ageMs = Math.max(0, now.getTime() - Date.parse(record.heartbeatAt));
    const alive = (options.processAlive ?? isProcessAlive)(record.pid);
    return { status: alive ? "active" : "stale", record, ageMs };
  } catch (error) {
    return {
      status: "unreadable",
      reason: error instanceof Error ? error.message : "Run lock is invalid",
      ageMs: await pathAge(lockPath, now),
    };
  }
}

export async function discoverRecoveryRun(
  projectRoot: string,
): Promise<RecoveryDiscovery> {
  const lock = await inspectRunProcessLock(projectRoot);
  let runId: string | undefined;
  try {
    runId = (await activeRunId(projectRoot)) ?? (await newestRecoverableRunId(projectRoot));
  } catch (error) {
    return {
      kind: "unreadable",
      reason: error instanceof Error ? error.message : "Active recovery pointer is unreadable",
      lock,
    };
  }
  if (!runId) {
    if (lock.status === "active") {
      return {
        kind: "unreadable",
        runId: lock.record.runId,
        reason: `PGN process is active in PID ${lock.record.pid}, but no recoverable checkpoint is available yet`,
        lock,
      };
    }
    if (lock.status === "unreadable") {
      return {
        kind: "unreadable",
        reason: `PGN process lock is unreadable: ${lock.reason}`,
        lock,
      };
    }
    return { kind: "none", lock };
  }
  try {
    const recovered = await readRecoveryRun(projectRoot, runId);
    if (!isRecoverableStatus(recovered.state.status)) {
      const fallbackRunId = await newestRecoverableRunId(projectRoot);
      if (!fallbackRunId || fallbackRunId === runId) {
        return { kind: "none", lock };
      }
      const fallback = await readRecoveryRun(projectRoot, fallbackRunId);
      const matchingFallbackLock =
        lock.status === "active" &&
        (!lock.record.runId || lock.record.runId === fallback.state.runId);
      return {
        kind: matchingFallbackLock ? "running" : "recoverable",
        ...fallback,
        lock,
      };
    }
    const matchingActiveLock =
      lock.status === "active" &&
      (!lock.record.runId || lock.record.runId === recovered.state.runId);
    return {
      kind: matchingActiveLock ? "running" : "recoverable",
      ...recovered,
      lock,
    };
  } catch (error) {
    return {
      kind: "unreadable",
      runId,
      reason: error instanceof Error ? error.message : "Recovery files are unreadable",
      lock,
    };
  }
}

async function writeActivePointer(projectRoot: string, runId: string, now: Date): Promise<void> {
  await atomicWriteJson(recoveryPaths(projectRoot).active, {
    schemaVersion: RECOVERY_SCHEMA_VERSION,
    runId,
    updatedAt: now.toISOString(),
  } satisfies ActiveRunPointer);
}

async function clearActivePointer(projectRoot: string, runId: string): Promise<void> {
  try {
    if ((await activeRunId(projectRoot)) === runId) {
      await rm(recoveryPaths(projectRoot).active, { force: true });
      await syncDirectory(recoveryPaths(projectRoot).root);
    }
  } catch {
    // The terminal state remains authoritative if pointer cleanup cannot finish.
  }
}

function cloneState(state: RecoveryRunState): RecoveryRunState {
  return structuredClone(state);
}

export class RecoveryCheckpoint {
  private pending: Promise<void> = Promise.resolve();

  constructor(
    readonly projectRoot: string,
    private current: RecoveryRunState,
  ) {}

  snapshot(): RecoveryRunState {
    return cloneState(this.current);
  }

  async update(
    change: (
      current: RecoveryRunState,
    ) => RecoveryRunState | void | Promise<RecoveryRunState | void>,
  ): Promise<RecoveryRunState> {
    let result!: RecoveryRunState;
    const operation = this.pending.then(async () => {
      const draft = cloneState(this.current);
      const changed = await change(draft);
      const next = parseRecoveryState(changed ?? draft);
      if (this.current.isDemo && !next.isDemo) {
        throw new Error("Recovery demo runs cannot become real runs");
      }
      if (readSessionMode(this.current.sessionMode) !== readSessionMode(next.sessionMode) ||
          readExecutionTransport(this.current.transport) !== readExecutionTransport(next.transport)) {
        throw new Error("A run's session mode and transport cannot be changed; start a new run");
      }
      if (JSON.stringify(this.current.restTarget) !== JSON.stringify(next.restTarget)) {
        throw new Error("A run's REST account and skill cannot be changed; start a new run");
      }
      const paths = recoveryPaths(this.projectRoot, next.runId);
      await atomicWriteJson(paths.state!, next);
      this.current = next;
      result = cloneState(next);
      if (isRecoverableStatus(next.status)) {
        await writeActivePointer(this.projectRoot, next.runId, new Date(next.updatedAt));
      } else {
        await clearActivePointer(this.projectRoot, next.runId);
      }
    });
    this.pending = operation.catch(() => undefined);
    await operation;
    return result;
  }

  heartbeat(now = new Date()): Promise<RecoveryRunState> {
    return this.update((state) => {
      state.heartbeatAt = now.toISOString();
      state.updatedAt = now.toISOString();
    });
  }

  async flush(): Promise<void> {
    await this.pending;
  }
}

export async function initializeRecoveryCheckpoint(options: {
  projectRoot: string;
  runId: string;
  isDemo?: true;
  mode: "full" | "retest";
  sessionMode?: SessionMode;
  transport?: ExecutionTransport;
  restartedFromRunId?: string;
  restTarget?: { accountId: string; skillId: string };
  sourceWorkbookPath: string;
  executedWorkbookPath: string;
  sourceWorkbookHash: string;
  scenarios: readonly PgnTestScenario[];
  startedAt?: Date;
}): Promise<{ checkpoint: RecoveryCheckpoint; manifest: RecoveryRunManifest }> {
  const now = options.startedAt ?? new Date();
  const relative = (filePath: string): string =>
    path.relative(options.projectRoot, filePath).replaceAll(path.sep, "/");
  const manifest = createRecoveryManifest(
    options.runId,
    options.sourceWorkbookHash,
    options.scenarios,
    now,
    options,
  );
  if (options.isDemo) manifest.isDemo = true;
  const state: RecoveryRunState = {
    schemaVersion: RECOVERY_SCHEMA_VERSION,
    runId: options.runId,
    ...(options.isDemo ? { isDemo: true as const } : {}),
    mode: options.mode,
    sessionMode: readSessionMode(options.sessionMode),
    transport: readExecutionTransport(options.transport),
    sessionResetAttempts: 0,
    ...(options.restartedFromRunId ? { restartedFromRunId: options.restartedFromRunId } : {}),
    ...(options.restTarget ? { restTarget: options.restTarget } : {}),
    status: "PREPARING",
    sourceWorkbookPath: relative(options.sourceWorkbookPath),
    sourceWorkbookHash: options.sourceWorkbookHash,
    executedWorkbookPath: relative(options.executedWorkbookPath),
    selectedScenarioIds: options.scenarios.map((scenario) => scenario.testCaseId),
    completedScenarioIds: [],
    skippedScenarioIds: [],
    totalScenarios: options.scenarios.length,
    scenarioAttempts: [],
    reconciliationDecisions: [],
    workbookProgress: "Run checkpoint created",
    finalCleanupComplete: false,
    metrics: {
      executedScenarios: 0,
      capturedScenarios: 0,
      timeouts: 0,
      technicalErrors: 0,
      evidenceCaptured: 0,
      evidenceUploaded: 0,
      evidenceUploadErrors: 0,
    },
    startedAt: now.toISOString(),
    updatedAt: now.toISOString(),
    heartbeatAt: now.toISOString(),
    resumeCount: 0,
  };
  const validatedState = parseRecoveryState(state);
  const validatedManifest = parseManifest(manifest);
  const paths = recoveryPaths(options.projectRoot, options.runId);
  await mkdir(paths.runs, { recursive: true });
  try {
    await mkdir(paths.runDirectory!);
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code === "EEXIST") {
      throw new Error(`Recovery history already exists for Run ${options.runId}`);
    }
    throw error;
  }
  await atomicWriteJson(paths.manifest!, validatedManifest);
  await atomicWriteJson(paths.state!, validatedState);
  await writeActivePointer(options.projectRoot, options.runId, now);
  return {
    checkpoint: new RecoveryCheckpoint(options.projectRoot, validatedState),
    manifest: validatedManifest,
  };
}

export async function openRecoveryCheckpoint(
  projectRoot: string,
  runId: string,
): Promise<{ checkpoint: RecoveryCheckpoint; manifest: RecoveryRunManifest }> {
  const recovered = await readRecoveryRun(projectRoot, runId);
  return {
    checkpoint: new RecoveryCheckpoint(projectRoot, recovered.state),
    manifest: recovered.manifest,
  };
}

async function removeOwnedRunLock(lockPath: string, token: string): Promise<void> {
  await Promise.all([
    unlink(path.join(lockPath, `owner-${token}.json`)).catch((error) => {
      if ((error as NodeJS.ErrnoException).code !== "ENOENT") throw error;
    }),
    unlink(path.join(lockPath, `heartbeat-${token}.json`)).catch((error) => {
      if ((error as NodeJS.ErrnoException).code !== "ENOENT") throw error;
    }),
  ]);
  await rmdir(lockPath).catch((error) => {
    const code = (error as NodeJS.ErrnoException).code;
    if (code !== "ENOENT" && code !== "ENOTEMPTY" && code !== "EEXIST") {
      throw error;
    }
  });
}

export async function acquireRunProcessLock(
  projectRoot: string,
  purpose: string,
  options: {
    now?: () => Date;
    processAlive?: (pid: number) => boolean;
    recoverStale?: boolean;
  } = {},
): Promise<RunProcessLock> {
  const paths = recoveryPaths(projectRoot);
  await mkdir(paths.root, { recursive: true });
  const token = randomUUID();
  const pending = path.join(paths.root, `.run-lock-${token}.pending`);
  const ownerPath = path.join(paths.lock, `owner-${token}.json`);
  const heartbeatPath = path.join(paths.lock, `heartbeat-${token}.json`);
  const now = options.now ?? (() => new Date());
  let record: RunLockRecord = {
    schemaVersion: RECOVERY_SCHEMA_VERSION,
    token,
    pid: process.pid,
    purpose,
    createdAt: now().toISOString(),
    heartbeatAt: now().toISOString(),
  };
  await atomicWriteJson(pending, record);
  try {
    for (let attempt = 0; attempt < 5; attempt += 1) {
      try {
        await mkdir(paths.lock);
      } catch (error) {
        if ((error as NodeJS.ErrnoException).code !== "EEXIST") throw error;
        if (options.recoverStale === false) {
          throw new Error("A PGN process lock already exists; demo operations will not alter it");
        }
        const inspection = await inspectRunProcessLock(projectRoot, {
          now: now(),
          processAlive: options.processAlive,
        });
        if (inspection.status === "active") {
          throw new Error(
            `Another PGN process is active (PID ${inspection.record.pid}, ${inspection.record.purpose})`,
          );
        }
        const recoverableUnreadable =
          inspection.status === "unreadable" &&
          inspection.ageMs !== undefined &&
          inspection.ageMs >= RUN_LOCK_STALE_MS;
        if (inspection.status !== "stale" && !recoverableUnreadable) {
          throw new Error(
            inspection.status === "unreadable"
              ? `PGN run lock is unreadable: ${inspection.reason}`
              : "PGN run lock could not be acquired",
          );
        }
        const quarantine = `${paths.lock}.stale-${randomUUID()}`;
        try {
          await rename(paths.lock, quarantine);
          await rm(quarantine, { recursive: true, force: true });
        } catch (takeoverError) {
          if ((takeoverError as NodeJS.ErrnoException).code === "ENOENT") continue;
          throw takeoverError;
        }
        continue;
      }

      let installed = false;
      try {
        await rename(pending, ownerPath);
        installed = true;
        await atomicWriteJson(heartbeatPath, record);
        const entries = (await readdir(paths.lock)).sort();
        const expectedEntries = [
          `heartbeat-${token}.json`,
          `owner-${token}.json`,
        ].sort();
        if (JSON.stringify(entries) !== JSON.stringify(expectedEntries)) {
          throw new Error("PGN run lock ownership changed during acquisition");
        }
      } catch (error) {
        if (installed) await removeOwnedRunLock(paths.lock, token);
        else await rmdir(paths.lock).catch(() => undefined);
        throw error;
      }

      let released = false;
      let pendingOperation: Promise<void> = Promise.resolve();
      const enqueue = (operation: () => Promise<void>): Promise<void> => {
        const result = pendingOperation.then(operation);
        pendingOperation = result.catch(() => undefined);
        return result;
      };
      return {
        path: paths.lock,
        token,
        heartbeat: (run) =>
          enqueue(async () => {
            if (released) return;
            record = {
              ...record,
              ...run,
              heartbeatAt: now().toISOString(),
            };
            await atomicWriteJson(heartbeatPath, record);
          }),
        release: () =>
          enqueue(async () => {
            if (released) return;
            released = true;
            await removeOwnedRunLock(paths.lock, token);
          }),
      };
    }
    throw new Error("PGN run lock could not be acquired");
  } finally {
    await rm(pending, { force: true }).catch(() => undefined);
  }
}
