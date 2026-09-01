import {
  constants as fsConstants,
  access,
  copyFile,
  mkdir,
} from "node:fs/promises";
import path from "node:path";
import ExcelJS from "exceljs";
import type { AppConfig } from "../config";
import {
  EVIDENCE_MIGRATION_VERSION,
  TRANSCRIPT_EVIDENCE_STATUS_COLUMN,
  TRANSCRIPT_EVIDENCE_URL_COLUMN,
  getEvidenceFileMetadata,
  getEvidenceRunMetadata,
  readEvidenceHyperlink,
  upsertEvidenceFileMetadata,
  upsertEvidenceRunMetadata,
  writeEvidenceHyperlink,
  writeMainEvidenceHyperlink,
} from "../excel/evidence-workbook";
import { cellText, parsePgnWorkbook } from "../excel/pgn-workbook-loader";
import {
  NEGATIVE_SHEET_NAME,
  TRANSCRIPT_SHEET_NAME,
  type EvidenceStatus,
} from "../excel/pgn-types";
import {
  openExecutedPgnWorkbook,
  saveExecutedPgnWorkbook,
} from "../excel/pgn-workbook-writer";
import type { EvidenceDrivePublisher } from "./google-drive";
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
  missingCompletedTurns: string[];
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
}

function numericCell(value: unknown): number | undefined {
  const number = Number(value);
  return Number.isInteger(number) && number > 0 ? number : undefined;
}

function uniqueValue(values: Set<string>): string | undefined {
  return values.size === 1 ? [...values][0] : undefined;
}

export function evidenceFileName(testCaseId: string, turnNumber: number): string {
  return `${testCaseId}-turn-${turnNumber}.png`;
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
      paths: new Set<string>(),
      urls: new Set<string>(),
      statuses: new Set<string>(),
    };
    group.transcriptRows.push(rowNumber);
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

  const records = [...groups.entries()]
    .map(([evidenceKey, group]): TranscriptEvidenceRecord => {
      const invalidReasons: string[] = [];
      if (group.paths.size > 1) {
        invalidReasons.push("transcript rows contain multiple local evidence paths");
      }
      if (group.urls.size > 1) {
        invalidReasons.push("transcript rows contain multiple Evidence URLs");
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
  const parsed = parsePgnWorkbook(workbook);
  const missingCompletedTurns: string[] = [];
  for (const scenario of parsed.scenarios) {
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
        missingCompletedTurns.push(
          `${scenario.testCaseId} turn ${turn.turnNumber}`,
        );
      }
    }
  }

  return {
    records,
    runIds: [...new Set(records.map((record) => record.runId))],
    missingCompletedTurns,
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
  return requested.map(([testCaseId, turnNumber]) => {
    const record = inventory.records.find(
      (candidate) =>
        candidate.testCaseId === testCaseId &&
        candidate.turnNumber === turnNumber,
    );
    if (!record?.localEvidencePath) {
      throw new Error(
        `Representative evidence was not found for ${testCaseId} turn ${turnNumber}`,
      );
    }
    return record;
  });
}

function resolveProjectFile(projectRoot: string, relativePath: string): string {
  const absolutePath = path.resolve(projectRoot, relativePath);
  const relative = path.relative(projectRoot, absolutePath);
  if (relative.startsWith("..") || path.isAbsolute(relative)) {
    throw new Error(`Evidence path escapes the project directory: ${relativePath}`);
  }
  return absolutePath;
}

function relativeToProject(projectRoot: string, absolutePath: string): string {
  return path.relative(projectRoot, absolutePath).replaceAll(path.sep, "/");
}

export async function createRepresentativeEvidencePreviews(
  config: AppConfig,
  inventory: EvidenceInventory,
  log: (line: string) => void = console.log,
): Promise<EvidencePreviewResult> {
  const records = representativeRecords(inventory);
  const sourcePaths = records.map((record) =>
    resolveProjectFile(config.projectRoot, record.localEvidencePath!),
  );
  const crop = await validateLegacyCropStrategy(
    sourcePaths,
    config.legacyEvidenceCropLeft,
  );
  const previewDirectory = path.join(config.evidenceCleanDir, "previews");
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
  const backupPath = path.join(
    config.reportArchiveDir,
    `${baseName}-before-evidence${extension}`,
  );
  if (!(await fileExists(backupPath))) {
    await copyFile(
      config.pgnExecutedWorkbookPath,
      backupPath,
      fsConstants.COPYFILE_EXCL,
    );
  }
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
  records: TranscriptEvidenceRecord[],
): boolean {
  if (record.sheetName !== NEGATIVE_SHEET_NAME) {
    return true;
  }
  const finalTurn = Math.max(
    ...records
      .filter((candidate) => candidate.testCaseId === record.testCaseId)
      .map((candidate) => candidate.turnNumber),
  );
  return record.turnNumber === finalTurn;
}

function applyEvidenceUrl(
  workbook: ExcelJS.Workbook,
  inventory: EvidenceInventory,
  record: TranscriptEvidenceRecord,
  status: EvidenceStatus,
  url: string,
): boolean {
  let changed = writeTranscriptEvidence(workbook, record, status, url);
  if (shouldWriteMainEvidence(record, inventory.records)) {
    changed =
      writeMainEvidenceHyperlink(
        workbook,
        record.sheetName,
        record.rowNumber,
        url,
      ) || changed;
  }
  return changed;
}

function recordLabel(record: TranscriptEvidenceRecord): string {
  return `${record.testCaseId} turn ${record.turnNumber}`;
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
  const backupPath = await createEvidenceWorkbookBackup(config);
  const opened = await openExecutedPgnWorkbook(
    config.pgnSourceWorkbookPath,
    config.pgnExecutedWorkbookPath,
  );
  const inventory = discoverEvidenceInventory(opened.workbook);
  const preview = await createRepresentativeEvidencePreviews(config, inventory, log);
  const runId = inventory.runIds.at(-1) ?? fallbackRunId(now());
  const storedRun = getEvidenceRunMetadata(opened.workbook, runId);
  await publisher.validateParentFolder();
  const driveFolder = await publisher.ensureRunFolder(
    runId,
    storedRun?.folderId,
  );
  upsertEvidenceRunMetadata(opened.workbook, {
    runId,
    folderId: driveFolder.id,
    folderUrl: driveFolder.webViewLink,
    migrationVersion: EVIDENCE_MIGRATION_VERSION,
    timestamp: now(),
    mode: "MIGRATION",
  });
  await saveExecutedPgnWorkbook(
    opened.workbook,
    config.pgnExecutedWorkbookPath,
  );

  let successfullyCleaned = 0;
  let successfullyUploaded = 0;
  let alreadySynced = 0;
  let requiresRerun = 0;
  let uploadErrors = 0;
  const missingIds = [...inventory.missingCompletedTurns];
  const requiresRerunIds: string[] = [];
  const uploadErrorIds: string[] = [];

  for (const record of inventory.records) {
    const label = recordLabel(record);
    const driveFileName = evidenceFileName(
      record.testCaseId,
      record.turnNumber,
    );
    log(`[Evidence] ${label}`);
    const storedFile = getEvidenceFileMetadata(
      opened.workbook,
      record.evidenceKey,
    );
    const existingUrl = record.evidenceUrl ?? storedFile?.evidenceUrl;
    if (existingUrl) {
      applyEvidenceUrl(
        opened.workbook,
        inventory,
        record,
        "EVIDENCE_ALREADY_SYNCED",
        existingUrl,
      );
      upsertEvidenceFileMetadata(opened.workbook, {
        evidenceKey: record.evidenceKey,
        runId: record.runId,
        testCaseId: record.testCaseId,
        turnNumber: record.turnNumber,
        driveFileId: storedFile?.driveFileId,
        driveFileName: storedFile?.driveFileName || driveFileName,
        evidenceUrl: existingUrl,
        localCleanPath: storedFile?.localCleanPath,
        status: "EVIDENCE_ALREADY_SYNCED",
      });
      await saveExecutedPgnWorkbook(
        opened.workbook,
        config.pgnExecutedWorkbookPath,
      );
      alreadySynced += 1;
      continue;
    }

    if (record.invalidReason) {
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

    const originalPath = resolveProjectFile(
      config.projectRoot,
      record.localEvidencePath,
    );
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

    const cleanedPath = path.join(
      config.evidenceCleanDir,
      record.runId,
      driveFileName,
    );
    const cleaned = await cropLegacyEvidence(
      originalPath,
      cleanedPath,
      preview.crop.left,
      config.legacyEvidenceCropLeft,
    );
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
    const localCleanPath = relativeToProject(config.projectRoot, cleanedPath);

    try {
      const uploaded = await publisher.uploadPng({
        folderId: driveFolder.id,
        localPath: cleanedPath,
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
        `[Evidence] EVIDENCE_UPLOAD_ERROR: ${error instanceof Error ? error.message : String(error)}`,
      );
    }
    await saveExecutedPgnWorkbook(
      opened.workbook,
      config.pgnExecutedWorkbookPath,
    );
    log("[Workbook] Saved");
  }

  const uniqueMissingIds = [...new Set(missingIds)];
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
    runId,
    driveParentId: publisher.parentFolderId,
    driveFolderId: driveFolder.id,
    driveFolderUrl: driveFolder.webViewLink,
    driveFolderReused: driveFolder.reused,
  };
}
