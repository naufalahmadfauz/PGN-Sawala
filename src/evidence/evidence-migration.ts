import {
  constants as fsConstants,
  access,
  copyFile,
  mkdir,
  realpath,
} from "node:fs/promises";
import { createHash } from "node:crypto";
import path from "node:path";
import ExcelJS from "exceljs";
import type { AppConfig } from "../config";
import {
  EVIDENCE_MIGRATION_VERSION,
  MAIN_EVIDENCE_COLUMN,
  TRANSCRIPT_EVIDENCE_STATUS_COLUMN,
  TRANSCRIPT_EVIDENCE_URL_COLUMN,
  getEvidenceFileMetadata,
  getEvidenceRunMetadata,
  removeEvidenceFileMetadata,
  readEvidenceHyperlink,
  upsertEvidenceFileMetadata,
  upsertEvidenceRunMetadata,
  writeEvidenceHyperlink,
  writeMainEvidenceHyperlink,
} from "../excel/evidence-workbook";
import { cellText, parsePgnWorkbook } from "../excel/pgn-workbook-loader";
import { updateRetestHistoryEvidence } from "../excel/retest-workbook";
import {
  NEGATIVE_SHEET_NAME,
  TRANSCRIPT_SHEET_NAME,
  type EvidenceStatus,
} from "../excel/pgn-types";
import {
  openExecutedPgnWorkbook,
  saveExecutedPgnWorkbook,
} from "../excel/pgn-workbook-writer";
import { retestDriveFolderName } from "../retest/retest-run";
import type { EvidenceDrivePublisher } from "./google-drive";
import { safeGoogleCredentialError } from "./google-service-account";
import {
  cropLegacyEvidence,
  validateLegacyCropStrategy,
  type CleanedEvidenceResult,
  type LegacyCropDefinition,
} from "./legacy-evidence-crop";

const MIGRATABLE_ROLES = new Set(["USER", "BOT", "SYSTEM"]);

export interface TranscriptEvidenceRecord {
  evidenceKey: string;
  runId: string;
  testCaseId: string;
  sheetName: string;
  rowNumber: number;
  turnNumber: number;
  transcriptRows: number[];
  localEvidencePath?: string;
  evidenceUrl?: string;
  evidenceStatus?: EvidenceStatus;
  invalidReason?: string;
}

export interface EvidenceInventory {
  records: TranscriptEvidenceRecord[];
  runIds: string[];
  missingCompletedTurns: MissingCompletedEvidence[];
  finalTurnNumbers: Record<string, number>;
}

export interface MissingCompletedEvidence {
  testCaseId: string;
  sheetName: string;
  rowNumber: number;
  turnNumber: number;
}

export interface EvidencePreviewResult {
  crop: LegacyCropDefinition;
  previews: Array<
    CleanedEvidenceResult & { testCaseId: string; turnNumber: number }
  >;
}

export interface EvidenceMigrationSummary {
  totalEvidenceDiscovered: number;
  successfullyCleaned: number;
  successfullyUploaded: number;
  alreadySynced: number;
  missing: number;
  requiresRerun: number;
  uploadErrors: number;
  missingIds: string[];
  requiresRerunIds: string[];
  uploadErrorIds: string[];
  backupPath: string;
  runId: string;
  driveParentId: string;
  driveFolderId: string;
  driveFolderUrl: string;
  driveFolderReused: boolean;
  driveFolders: Array<{
    runId: string;
    id: string;
    url: string;
    reused: boolean;
  }>;
}

function numericCell(value: unknown): number | undefined {
  const number = Number(value);
  return Number.isInteger(number) && number > 0 ? number : undefined;
}

function uniqueValue(values: Set<string>): string | undefined {
  return values.size === 1 ? [...values][0] : undefined;
}

function safePathSegment(value: string): string {
  const sanitized = value
    .replace(/[^A-Za-z0-9_-]+/g, "-")
    .replace(/^-+|-+$/g, "")
    .slice(0, 100);
  if (!sanitized) {
    throw new Error("Evidence identifier cannot be converted to a safe filename");
  }
  if (sanitized === value) {
    return sanitized;
  }
  const suffix = createHash("sha256").update(value).digest("hex").slice(0, 8);
  return `${sanitized.slice(0, 90)}-${suffix}`;
}

export function evidenceFileName(testCaseId: string, turnNumber: number): string {
  return `${safePathSegment(testCaseId)}-turn-${turnNumber}.png`;
}

export function discoverEvidenceInventory(
  workbook: ExcelJS.Workbook,
): EvidenceInventory {
  const transcript = workbook.getWorksheet(TRANSCRIPT_SHEET_NAME);
  if (!transcript) {
    throw new Error(`Worksheet "${TRANSCRIPT_SHEET_NAME}" was not found`);
  }
  const groups = new Map<
    string,
    {
      runId: string;
      testCaseId: string;
      sheetName: string;
      rowNumber: number;
      turnNumber: number;
      transcriptRows: number[];
      locations: Set<string>;
      paths: Set<string>;
      urls: Set<string>;
      statuses: Set<string>;
    }
  >();

  for (let rowNumber = 2; rowNumber <= transcript.rowCount; rowNumber += 1) {
    const row = transcript.getRow(rowNumber);
    if (!MIGRATABLE_ROLES.has(row.getCell(6).text)) {
      continue;
    }
    const runId = row.getCell(1).text.trim();
    const testCaseId = row.getCell(2).text.trim();
    const sheetName = row.getCell(3).text.trim();
    const excelRow = numericCell(row.getCell(4).value);
    const turnNumber = numericCell(row.getCell(5).value);
    if (!runId || !testCaseId || !sheetName || !excelRow || !turnNumber) {
      continue;
    }
    const key = `${runId}|${testCaseId}|${turnNumber}`;
    const group = groups.get(key) ?? {
      runId,
      testCaseId,
      sheetName,
      rowNumber: excelRow,
      turnNumber,
      transcriptRows: [],
      locations: new Set<string>(),
      paths: new Set<string>(),
      urls: new Set<string>(),
      statuses: new Set<string>(),
    };
    group.transcriptRows.push(rowNumber);
    group.locations.add(`${sheetName}|${excelRow}`);
    const localPath = row.getCell(13).text.trim();
    const url = readEvidenceHyperlink(
      row.getCell(TRANSCRIPT_EVIDENCE_URL_COLUMN),
    );
    const status = row.getCell(TRANSCRIPT_EVIDENCE_STATUS_COLUMN).text.trim();
    if (localPath) {
      group.paths.add(localPath);
    }
    if (url) {
      group.urls.add(url);
    }
    if (status) {
      group.statuses.add(status);
    }
    groups.set(key, group);
  }

  const parsed = parsePgnWorkbook(workbook);
  const scenariosById = new Map(
    parsed.scenarios.map((scenario) => [scenario.testCaseId, scenario]),
  );
  const records = [...groups.entries()]
    .map(([evidenceKey, group]): TranscriptEvidenceRecord => {
      const invalidReasons: string[] = [];
      if (group.locations.size > 1) {
        invalidReasons.push("transcript rows contain multiple result locations");
      }
      if (group.paths.size > 1) {
        invalidReasons.push("transcript rows contain multiple local evidence paths");
      }
      if (group.urls.size > 1) {
        invalidReasons.push("transcript rows contain multiple Evidence URLs");
      }
      const scenario = scenariosById.get(group.testCaseId);
      const expectedTurn = scenario?.turns.find(
        (turn) => turn.turnNumber === group.turnNumber,
      );
      if (!scenario || !expectedTurn) {
        invalidReasons.push("transcript record does not map to a workbook scenario turn");
      } else if (
        scenario.sheetName !== group.sheetName ||
        expectedTurn.rowNumber !== group.rowNumber
      ) {
        invalidReasons.push("transcript result location does not match the workbook scenario");
      }
      return {
        evidenceKey,
        runId: group.runId,
        testCaseId: group.testCaseId,
        sheetName: group.sheetName,
        rowNumber: group.rowNumber,
        turnNumber: group.turnNumber,
        transcriptRows: group.transcriptRows,
        localEvidencePath: uniqueValue(group.paths),
        evidenceUrl: uniqueValue(group.urls),
        evidenceStatus: uniqueValue(group.statuses) as
          | EvidenceStatus
          | undefined,
        invalidReason: invalidReasons.length
          ? invalidReasons.join("; ")
          : undefined,
      };
    })
    .sort((left, right) => left.transcriptRows[0] - right.transcriptRows[0]);

  const coveredTurns = new Set(
    records.map((record) => `${record.testCaseId}|${record.turnNumber}`),
  );
  const missingCompletedTurns: MissingCompletedEvidence[] = [];
  const finalTurnNumbers: Record<string, number> = {};
  for (const scenario of parsed.scenarios) {
    finalTurnNumbers[scenario.testCaseId] =
      scenario.turns.at(-1)?.turnNumber ?? 1;
    const worksheet = workbook.getWorksheet(scenario.sheetName);
    if (!worksheet) {
      continue;
    }
    const scenarioComplete = scenario.sheetKind === "negative"
      ? Boolean(cellText(worksheet.getCell(scenario.sourceRowNumber, 8)).trim())
      : undefined;
    for (const turn of scenario.turns) {
      const completed =
        scenarioComplete ??
        Boolean(cellText(worksheet.getCell(turn.rowNumber, 9)).trim());
      const turnKey = `${scenario.testCaseId}|${turn.turnNumber}`;
      if (completed && !coveredTurns.has(turnKey)) {
        missingCompletedTurns.push({
          testCaseId: scenario.testCaseId,
          sheetName: scenario.sheetName,
          rowNumber: turn.rowNumber,
          turnNumber: turn.turnNumber,
        });
      }
    }
  }

  return {
    records,
    runIds: [...new Set(records.map((record) => record.runId))],
    missingCompletedTurns,
    finalTurnNumbers,
  };
}

function representativeRecords(
  inventory: EvidenceInventory,
): TranscriptEvidenceRecord[] {
  const requested = [
    ["PGN-KB-003", 1],
    ["PGN-KB-031", 2],
    ["PGN-KB-075", 1],
  ] as const;
  const selected = requested.flatMap(([testCaseId, turnNumber]) => {
    const record = inventory.records.find(
      (candidate) =>
        candidate.testCaseId === testCaseId &&
        candidate.turnNumber === turnNumber &&
        candidate.localEvidencePath,
    );
    return record ? [record] : [];
  });
  for (const record of inventory.records) {
    if (
      selected.length >= 3 ||
      !record.localEvidencePath ||
      selected.some((candidate) => candidate.evidenceKey === record.evidenceKey)
    ) {
      continue;
    }
    selected.push(record);
  }
  if (selected.length < 3) {
    throw new Error("At least three mapped local screenshots are required for crop validation");
  }
  return selected.slice(0, 3);
}

function resolveEvidenceSourcePath(
  config: AppConfig,
  relativePath: string,
): string {
  const absolutePath = path.resolve(config.projectRoot, relativePath);
  const relative = path.relative(config.evidenceDir, absolutePath);
  if (relative.startsWith("..") || path.isAbsolute(relative)) {
    throw new Error(
      `Evidence path is outside the active evidence directory: ${relativePath}`,
    );
  }
  return absolutePath;
}

async function assertRealPathInside(
  directory: string,
  candidate: string,
): Promise<void> {
  const [realDirectory, realCandidate] = await Promise.all([
    realpath(directory),
    realpath(candidate),
  ]);
  const relative = path.relative(realDirectory, realCandidate);
  if (relative.startsWith("..") || path.isAbsolute(relative)) {
    throw new Error(`Evidence path resolves outside ${directory}`);
  }
}

function relativeToProject(projectRoot: string, absolutePath: string): string {
  return path.relative(projectRoot, absolutePath).replaceAll(path.sep, "/");
}

function resolveInsideDirectory(directory: string, ...segments: string[]): string {
  const absolutePath = path.resolve(directory, ...segments);
  const relative = path.relative(directory, absolutePath);
  if (!relative || relative.startsWith("..") || path.isAbsolute(relative)) {
    throw new Error(`Evidence output escapes the configured clean directory`);
  }
  return absolutePath;
}

export async function createRepresentativeEvidencePreviews(
  config: AppConfig,
  inventory: EvidenceInventory,
  log: (line: string) => void = console.log,
): Promise<EvidencePreviewResult> {
  const records = representativeRecords(inventory);
  const sourcePaths = records.map((record) =>
    resolveEvidenceSourcePath(config, record.localEvidencePath!),
  );
  await Promise.all(
    sourcePaths.map((sourcePath) =>
      assertRealPathInside(config.evidenceDir, sourcePath),
    ),
  );
  const crop = await validateLegacyCropStrategy(
    sourcePaths,
    config.legacyEvidenceCropLeft,
  );
  const previewDirectory = path.join(config.evidenceCleanDir, "previews");
  await mkdir(previewDirectory, { recursive: true });
  await assertRealPathInside(config.evidenceCleanDir, previewDirectory);
  const previews = [];
  for (let index = 0; index < records.length; index += 1) {
    const record = records[index];
    const outputPath = path.join(
      previewDirectory,
      evidenceFileName(record.testCaseId, record.turnNumber),
    );
    const result = await cropLegacyEvidence(
      sourcePaths[index],
      outputPath,
      crop.left,
      config.legacyEvidenceCropLeft,
    );
    if (!result.usable) {
      throw new Error(
        `Representative evidence is unusable for ${record.testCaseId}: ${result.reason}`,
      );
    }
    log("[Evidence Migration]");
    log(`Test: ${record.testCaseId} turn ${record.turnNumber}`);
    log(`Original: ${result.originalWidth}x${result.originalHeight}`);
    log(`Cleaned: ${result.width}x${result.height}`);
    previews.push({
      ...result,
      testCaseId: record.testCaseId,
      turnNumber: record.turnNumber,
    });
  }
  return { crop, previews };
}

async function fileExists(filePath: string): Promise<boolean> {
  return access(filePath)
    .then(() => true)
    .catch(() => false);
}

export async function createEvidenceWorkbookBackup(
  config: AppConfig,
): Promise<string> {
  await mkdir(config.reportArchiveDir, { recursive: true });
  const extension = path.extname(config.pgnExecutedWorkbookPath);
  const baseName = path.basename(config.pgnExecutedWorkbookPath, extension);
  const baseBackupPath = path.join(
    config.reportArchiveDir,
    `${baseName}-before-evidence${extension}`,
  );
  let backupPath = baseBackupPath;
  if (await fileExists(baseBackupPath)) {
    const timestamp = new Date().toISOString().replace(/[:.]/g, "-");
    let suffix = 0;
    do {
      backupPath = path.join(
        config.reportArchiveDir,
        `${baseName}-before-evidence-${timestamp}${suffix ? `-${suffix}` : ""}${extension}`,
      );
      suffix += 1;
    } while (await fileExists(backupPath));
  }
  await copyFile(
    config.pgnExecutedWorkbookPath,
    backupPath,
    fsConstants.COPYFILE_EXCL,
  );
  return backupPath;
}

function writeTranscriptEvidence(
  workbook: ExcelJS.Workbook,
  record: TranscriptEvidenceRecord,
  status: EvidenceStatus,
  url?: string,
): boolean {
  const transcript = workbook.getWorksheet(TRANSCRIPT_SHEET_NAME)!;
  let changed = false;
  for (const rowNumber of record.transcriptRows) {
    const row = transcript.getRow(rowNumber);
    if (url) {
      changed =
        writeEvidenceHyperlink(
          row.getCell(TRANSCRIPT_EVIDENCE_URL_COLUMN),
          url,
        ) || changed;
    } else if (row.getCell(TRANSCRIPT_EVIDENCE_URL_COLUMN).value) {
      row.getCell(TRANSCRIPT_EVIDENCE_URL_COLUMN).value = null;
      changed = true;
    }
    if (row.getCell(TRANSCRIPT_EVIDENCE_STATUS_COLUMN).text !== status) {
      row.getCell(TRANSCRIPT_EVIDENCE_STATUS_COLUMN).value = status;
      changed = true;
    }
  }
  return changed;
}

function shouldWriteMainEvidence(
  record: TranscriptEvidenceRecord,
  inventory: EvidenceInventory,
): boolean {
  if (record.sheetName !== NEGATIVE_SHEET_NAME) {
    return true;
  }
  return record.turnNumber === inventory.finalTurnNumbers[record.testCaseId];
}

function applyEvidenceUrl(
  workbook: ExcelJS.Workbook,
  inventory: EvidenceInventory,
  record: TranscriptEvidenceRecord,
  status: EvidenceStatus,
  url: string,
): boolean {
  let changed = writeTranscriptEvidence(workbook, record, status, url);
  if (shouldWriteMainEvidence(record, inventory)) {
    changed =
      writeMainEvidenceHyperlink(
        workbook,
        record.sheetName,
        record.rowNumber,
        url,
      ) || changed;
    changed =
      updateRetestHistoryEvidence(
        workbook,
        record.runId,
        record.testCaseId,
        record.sheetName === NEGATIVE_SHEET_NAME
          ? undefined
          : record.turnNumber,
        url,
      ) || changed;
  }
  return changed;
}

function clearMainEvidence(
  workbook: ExcelJS.Workbook,
  inventory: EvidenceInventory,
  record: TranscriptEvidenceRecord,
): void {
  if (!shouldWriteMainEvidence(record, inventory)) {
    return;
  }
  const worksheet = workbook.getWorksheet(record.sheetName);
  if (!worksheet) {
    throw new Error(`Worksheet "${record.sheetName}" was not found`);
  }
  worksheet.getCell(record.rowNumber, MAIN_EVIDENCE_COLUMN).value = null;
}

function recordLabel(record: TranscriptEvidenceRecord): string {
  return `${record.testCaseId} turn ${record.turnNumber}`;
}

function driveFileNameForRecord(
  record: TranscriptEvidenceRecord,
): string {
  return evidenceFileName(record.testCaseId, record.turnNumber);
}

function fallbackRunId(now: Date): string {
  return now.toISOString().replace(/[:.]/g, "-");
}

export async function runEvidenceMigration(options: {
  config: AppConfig;
  publisher: EvidenceDrivePublisher;
  now?: () => Date;
  log?: (line: string) => void;
}): Promise<EvidenceMigrationSummary> {
  const { config, publisher } = options;
  const now = options.now ?? (() => new Date());
  const log = options.log ?? console.log;
  await publisher.validateParentFolder();
  const preflightWorkbook = new ExcelJS.Workbook();
  await preflightWorkbook.xlsx.readFile(config.pgnExecutedWorkbookPath);
  const preflightInventory = discoverEvidenceInventory(preflightWorkbook);
  const preflightLegacyRecords = preflightInventory.records.filter((record) => {
    const mode = getEvidenceRunMetadata(preflightWorkbook, record.runId)?.mode;
    return !mode || mode === "MIGRATION";
  });
  const preview = preflightLegacyRecords.length
    ? await createRepresentativeEvidencePreviews(
        config,
        {
          ...preflightInventory,
          records: preflightLegacyRecords,
          runIds: [
            ...new Set(preflightLegacyRecords.map((record) => record.runId)),
          ],
        },
        log,
      )
    : undefined;
  const backupPath = await createEvidenceWorkbookBackup(config);
  const opened = await openExecutedPgnWorkbook(
    config.pgnSourceWorkbookPath,
    config.pgnExecutedWorkbookPath,
  );
  const inventory = discoverEvidenceInventory(opened.workbook);
  const runIds = inventory.runIds.length
    ? inventory.runIds
    : [fallbackRunId(now())];
  const driveFolders = [] as EvidenceMigrationSummary["driveFolders"];
  const driveFoldersByRun = new Map<string, (typeof driveFolders)[number]>();
  const runModesByRun = new Map<
    string,
    "MIGRATION" | "FUTURE" | "RETEST"
  >();
  for (const runId of runIds) {
    const storedRun = getEvidenceRunMetadata(opened.workbook, runId);
    runModesByRun.set(runId, storedRun?.mode ?? "MIGRATION");
    const folder = await publisher.ensureRunFolder(
      runId,
      storedRun?.folderId,
      storedRun?.mode === "RETEST"
        ? retestDriveFolderName(publisher.retestFolderPrefix, runId)
        : undefined,
    );
    const folderMetadata = {
      runId,
      id: folder.id,
      url: folder.webViewLink,
      reused: folder.reused,
    };
    driveFolders.push(folderMetadata);
    driveFoldersByRun.set(runId, folderMetadata);
    if (
      !storedRun ||
      storedRun.folderId !== folder.id ||
      storedRun.folderUrl !== folder.webViewLink
    ) {
      upsertEvidenceRunMetadata(opened.workbook, {
        runId,
        folderId: folder.id,
        folderUrl: folder.webViewLink,
        migrationVersion:
          storedRun?.migrationVersion || EVIDENCE_MIGRATION_VERSION,
        timestamp: storedRun?.timestamp ?? now(),
        mode: storedRun?.mode ?? "MIGRATION",
      });
      await saveExecutedPgnWorkbook(
        opened.workbook,
        config.pgnExecutedWorkbookPath,
      );
    }
  }

  for (const missing of inventory.missingCompletedTurns) {
    upsertEvidenceFileMetadata(opened.workbook, {
      evidenceKey: `MISSING|${missing.testCaseId}|${missing.turnNumber}`,
      runId: "",
      testCaseId: missing.testCaseId,
      turnNumber: missing.turnNumber,
      driveFileName: evidenceFileName(
        missing.testCaseId,
        missing.turnNumber,
      ),
      status: "EVIDENCE_MISSING",
    });
  }
  if (inventory.missingCompletedTurns.length) {
    await saveExecutedPgnWorkbook(
      opened.workbook,
      config.pgnExecutedWorkbookPath,
    );
  }

  let successfullyCleaned = 0;
  let successfullyUploaded = 0;
  let alreadySynced = 0;
  let requiresRerun = 0;
  let uploadErrors = 0;
  const missingIds = inventory.missingCompletedTurns.map(
    (missing) => `${missing.testCaseId} turn ${missing.turnNumber}`,
  );
  const requiresRerunIds: string[] = [];
  const uploadErrorIds: string[] = [];

  for (const record of inventory.records) {
    const label = recordLabel(record);
    const driveFileName = driveFileNameForRecord(record);
    log(`[Evidence] ${label}`);
    if (!record.invalidReason) {
      clearMainEvidence(opened.workbook, inventory, record);
      removeEvidenceFileMetadata(
        opened.workbook,
        `MISSING|${record.testCaseId}|${record.turnNumber}`,
      );
    }
    const storedFile = getEvidenceFileMetadata(
      opened.workbook,
      record.evidenceKey,
    );
    const invalidReason =
      record.invalidReason ??
      (record.evidenceStatus === "EVIDENCE_CAPTURE_ERROR"
        ? "The stored path is a full-page failure diagnostic, not conversation evidence"
        : undefined);
    if (invalidReason) {
      requiresRerun += 1;
      requiresRerunIds.push(label);
      writeTranscriptEvidence(
        opened.workbook,
        record,
        "EVIDENCE_REQUIRES_RERUN",
      );
      upsertEvidenceFileMetadata(opened.workbook, {
        evidenceKey: record.evidenceKey,
        runId: record.runId,
        testCaseId: record.testCaseId,
        turnNumber: record.turnNumber,
        driveFileName,
        status: "EVIDENCE_REQUIRES_RERUN",
      });
      log(`[Evidence] EVIDENCE_REQUIRES_RERUN: ${invalidReason}`);
      await saveExecutedPgnWorkbook(
        opened.workbook,
        config.pgnExecutedWorkbookPath,
      );
      continue;
    }

    if (!record.localEvidencePath) {
      missingIds.push(label);
      writeTranscriptEvidence(opened.workbook, record, "EVIDENCE_MISSING");
      upsertEvidenceFileMetadata(opened.workbook, {
        evidenceKey: record.evidenceKey,
        runId: record.runId,
        testCaseId: record.testCaseId,
        turnNumber: record.turnNumber,
        driveFileName,
        status: "EVIDENCE_MISSING",
      });
      await saveExecutedPgnWorkbook(
        opened.workbook,
        config.pgnExecutedWorkbookPath,
      );
      continue;
    }

    let originalPath: string;
    try {
      originalPath = resolveEvidenceSourcePath(
        config,
        record.localEvidencePath,
      );
    } catch (error) {
      requiresRerun += 1;
      requiresRerunIds.push(label);
      writeTranscriptEvidence(
        opened.workbook,
        record,
        "EVIDENCE_REQUIRES_RERUN",
      );
      upsertEvidenceFileMetadata(opened.workbook, {
        evidenceKey: record.evidenceKey,
        runId: record.runId,
        testCaseId: record.testCaseId,
        turnNumber: record.turnNumber,
        driveFileName,
        status: "EVIDENCE_REQUIRES_RERUN",
      });
      log(
        `[Evidence] EVIDENCE_REQUIRES_RERUN: ${error instanceof Error ? error.message : String(error)}`,
      );
      await saveExecutedPgnWorkbook(
        opened.workbook,
        config.pgnExecutedWorkbookPath,
      );
      continue;
    }
    if (!(await fileExists(originalPath))) {
      missingIds.push(label);
      writeTranscriptEvidence(opened.workbook, record, "EVIDENCE_MISSING");
      upsertEvidenceFileMetadata(opened.workbook, {
        evidenceKey: record.evidenceKey,
        runId: record.runId,
        testCaseId: record.testCaseId,
        turnNumber: record.turnNumber,
        driveFileName,
        status: "EVIDENCE_MISSING",
      });
      await saveExecutedPgnWorkbook(
        opened.workbook,
        config.pgnExecutedWorkbookPath,
      );
      continue;
    }
    try {
      await assertRealPathInside(config.evidenceDir, originalPath);
    } catch (error) {
      requiresRerun += 1;
      requiresRerunIds.push(label);
      writeTranscriptEvidence(
        opened.workbook,
        record,
        "EVIDENCE_REQUIRES_RERUN",
      );
      upsertEvidenceFileMetadata(opened.workbook, {
        evidenceKey: record.evidenceKey,
        runId: record.runId,
        testCaseId: record.testCaseId,
        turnNumber: record.turnNumber,
        driveFileName,
        status: "EVIDENCE_REQUIRES_RERUN",
      });
      log(
        `[Evidence] EVIDENCE_REQUIRES_RERUN: ${error instanceof Error ? error.message : String(error)}`,
      );
      await saveExecutedPgnWorkbook(
        opened.workbook,
        config.pgnExecutedWorkbookPath,
      );
      continue;
    }

    let uploadPath = originalPath;
    let localCleanPath = record.localEvidencePath;
    if (runModesByRun.get(record.runId) === "MIGRATION") {
      if (!preview) {
        throw new Error("Legacy evidence crop was not validated");
      }
      const cleanedPath = resolveInsideDirectory(
        config.evidenceCleanDir,
        safePathSegment(record.runId),
        driveFileName,
      );
      await mkdir(path.dirname(cleanedPath), { recursive: true });
      await assertRealPathInside(
        config.evidenceCleanDir,
        path.dirname(cleanedPath),
      );
      let cleaned: CleanedEvidenceResult;
      try {
        cleaned = await cropLegacyEvidence(
          originalPath,
          cleanedPath,
          preview.crop.left,
          config.legacyEvidenceCropLeft,
        );
      } catch (error) {
        requiresRerun += 1;
        requiresRerunIds.push(label);
        writeTranscriptEvidence(
          opened.workbook,
          record,
          "EVIDENCE_REQUIRES_RERUN",
        );
        upsertEvidenceFileMetadata(opened.workbook, {
          evidenceKey: record.evidenceKey,
          runId: record.runId,
          testCaseId: record.testCaseId,
          turnNumber: record.turnNumber,
          driveFileName,
          status: "EVIDENCE_REQUIRES_RERUN",
        });
        log(
          `[Evidence] EVIDENCE_REQUIRES_RERUN: ${error instanceof Error ? error.message : String(error)}`,
        );
        await saveExecutedPgnWorkbook(
          opened.workbook,
          config.pgnExecutedWorkbookPath,
        );
        continue;
      }
      if (!cleaned.usable) {
        requiresRerun += 1;
        requiresRerunIds.push(label);
        writeTranscriptEvidence(
          opened.workbook,
          record,
          "EVIDENCE_REQUIRES_RERUN",
        );
        upsertEvidenceFileMetadata(opened.workbook, {
          evidenceKey: record.evidenceKey,
          runId: record.runId,
          testCaseId: record.testCaseId,
          turnNumber: record.turnNumber,
          driveFileName,
          localCleanPath: relativeToProject(config.projectRoot, cleanedPath),
          status: "EVIDENCE_REQUIRES_RERUN",
        });
        await saveExecutedPgnWorkbook(
          opened.workbook,
          config.pgnExecutedWorkbookPath,
        );
        continue;
      }
      if (!cleaned.reused) {
        successfullyCleaned += 1;
      }
      log("[Evidence] Cropped");
      uploadPath = cleanedPath;
      localCleanPath = relativeToProject(config.projectRoot, cleanedPath);
    } else {
      log("[Evidence] Conversation-pane screenshot retained");
    }

    try {
      const uploaded = await publisher.uploadPng({
        folderId: driveFoldersByRun.get(record.runId)!.id,
        localPath: uploadPath,
        fileName: driveFileName,
        existingFileId: storedFile?.driveFileId,
      });
      if (!uploaded.reused) {
        successfullyUploaded += 1;
      } else {
        alreadySynced += 1;
      }
      applyEvidenceUrl(
        opened.workbook,
        inventory,
        record,
        "EVIDENCE_SYNCED",
        uploaded.webViewLink,
      );
      upsertEvidenceFileMetadata(opened.workbook, {
        evidenceKey: record.evidenceKey,
        runId: record.runId,
        testCaseId: record.testCaseId,
        turnNumber: record.turnNumber,
        driveFileId: uploaded.id,
        driveFileName: uploaded.name,
        evidenceUrl: uploaded.webViewLink,
        localCleanPath,
        status: "EVIDENCE_SYNCED",
      });
      log("[Evidence] Uploaded");
      log("[Evidence] Hyperlink written");
    } catch (error) {
      uploadErrors += 1;
      uploadErrorIds.push(label);
      writeTranscriptEvidence(
        opened.workbook,
        record,
        "EVIDENCE_UPLOAD_ERROR",
      );
      upsertEvidenceFileMetadata(opened.workbook, {
        evidenceKey: record.evidenceKey,
        runId: record.runId,
        testCaseId: record.testCaseId,
        turnNumber: record.turnNumber,
        driveFileName,
        localCleanPath,
        status: "EVIDENCE_UPLOAD_ERROR",
      });
      log(
        `[Evidence] EVIDENCE_UPLOAD_ERROR: ${safeGoogleCredentialError(error)}`,
      );
    }
    await saveExecutedPgnWorkbook(
      opened.workbook,
      config.pgnExecutedWorkbookPath,
    );
    log("[Workbook] Saved");
  }

  const uniqueMissingIds = [...new Set(missingIds)];
  const primaryFolder = driveFolders.at(-1)!;
  return {
    totalEvidenceDiscovered: inventory.records.length,
    successfullyCleaned,
    successfullyUploaded,
    alreadySynced,
    missing: uniqueMissingIds.length,
    requiresRerun,
    uploadErrors,
    missingIds: uniqueMissingIds,
    requiresRerunIds: [...new Set(requiresRerunIds)],
    uploadErrorIds: [...new Set(uploadErrorIds)],
    backupPath,
    runId: primaryFolder.runId,
    driveParentId: publisher.parentFolderId,
    driveFolderId: primaryFolder.id,
    driveFolderUrl: primaryFolder.url,
    driveFolderReused: primaryFolder.reused,
    driveFolders,
  };
}
