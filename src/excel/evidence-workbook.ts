import ExcelJS, { type Cell, type Worksheet } from "exceljs";
import {
  KB_SHEET_NAME,
  NEGATIVE_SHEET_NAME,
  TRANSCRIPT_SHEET_NAME,
  type EvidenceStatus,
} from "./pgn-types";

export const EXECUTION_METADATA_SHEET_NAME = "Execution Metadata";
export const EVIDENCE_MIGRATION_VERSION = "1";
export const MAIN_EVIDENCE_COLUMN = 14;
export const TRANSCRIPT_EVIDENCE_URL_COLUMN = 14;
export const TRANSCRIPT_EVIDENCE_STATUS_COLUMN = 15;

const RUN_HEADERS = [
  "Run ID",
  "Evidence Drive Folder ID",
  "Evidence Drive Folder URL",
  "Evidence Migration Version",
  "Migration Timestamp",
  "Mode",
];

const FILE_HEADERS = [
  "Evidence Key",
  "Run ID",
  "Test Case ID",
  "Turn",
  "Drive File ID",
  "Drive File Name",
  "Evidence URL",
  "Local Clean Path",
  "Evidence Status",
];

export interface EvidenceRunMetadata {
  runId: string;
  folderId: string;
  folderUrl: string;
  migrationVersion: string;
  timestamp: Date;
  mode: "MIGRATION" | "FUTURE";
}

export interface EvidenceFileMetadata {
  evidenceKey: string;
  runId: string;
  testCaseId: string;
  turnNumber: number;
  driveFileId?: string;
  driveFileName: string;
  evidenceUrl?: string;
  localCleanPath?: string;
  status: EvidenceStatus;
}

function cloneStyle(source: Cell, target: Cell): void {
  target.style = structuredClone(source.style);
}

function styleHeaderCell(cell: Cell): void {
  cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
  cell.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FF1F4E78" },
  };
  cell.alignment = { vertical: "middle", wrapText: true };
}

function ensureMainEvidenceColumn(worksheet: Worksheet): boolean {
  const header = worksheet.getCell(1, MAIN_EVIDENCE_COLUMN);
  if (header.text && header.text !== "Evidence") {
    throw new Error(
      `${worksheet.name}!N1 is already used by "${header.text}"; cannot add Evidence`,
    );
  }
  let changed = false;
  if (header.text !== "Evidence") {
    header.value = "Evidence";
    cloneStyle(worksheet.getCell(1, MAIN_EVIDENCE_COLUMN - 1), header);
    changed = true;
  }
  if ((worksheet.getColumn(MAIN_EVIDENCE_COLUMN).width ?? 0) < 18) {
    worksheet.getColumn(MAIN_EVIDENCE_COLUMN).width = 18;
    changed = true;
  }
  return changed;
}

function ensureTranscriptEvidenceColumns(worksheet: Worksheet): boolean {
  const urlHeader = worksheet.getCell(1, TRANSCRIPT_EVIDENCE_URL_COLUMN);
  const statusHeader = worksheet.getCell(1, TRANSCRIPT_EVIDENCE_STATUS_COLUMN);
  if (urlHeader.text && urlHeader.text !== "Evidence URL") {
    throw new Error(
      `${worksheet.name}!N1 is already used by "${urlHeader.text}"`,
    );
  }
  if (statusHeader.text && statusHeader.text !== "Evidence Status") {
    throw new Error(
      `${worksheet.name}!O1 is already used by "${statusHeader.text}"`,
    );
  }
  let changed = false;
  if (urlHeader.text !== "Evidence URL") {
    urlHeader.value = "Evidence URL";
    cloneStyle(worksheet.getCell(1, 13), urlHeader);
    changed = true;
  }
  if (statusHeader.text !== "Evidence Status") {
    statusHeader.value = "Evidence Status";
    cloneStyle(worksheet.getCell(1, 13), statusHeader);
    changed = true;
  }
  if ((worksheet.getColumn(TRANSCRIPT_EVIDENCE_URL_COLUMN).width ?? 0) < 20) {
    worksheet.getColumn(TRANSCRIPT_EVIDENCE_URL_COLUMN).width = 20;
    changed = true;
  }
  if ((worksheet.getColumn(TRANSCRIPT_EVIDENCE_STATUS_COLUMN).width ?? 0) < 24) {
    worksheet.getColumn(TRANSCRIPT_EVIDENCE_STATUS_COLUMN).width = 24;
    changed = true;
  }
  return changed;
}

function ensureMetadataWorksheet(workbook: ExcelJS.Workbook): {
  worksheet: Worksheet;
  changed: boolean;
} {
  let worksheet = workbook.getWorksheet(EXECUTION_METADATA_SHEET_NAME);
  let changed = false;
  if (!worksheet) {
    worksheet = workbook.addWorksheet(EXECUTION_METADATA_SHEET_NAME);
    worksheet.views = [{ state: "frozen", ySplit: 1 }];
    changed = true;
  }
  RUN_HEADERS.forEach((header, index) => {
    const cell = worksheet!.getCell(1, index + 1);
    if (!cell.text) {
      cell.value = header;
      styleHeaderCell(cell);
      changed = true;
    } else if (cell.text !== header) {
      throw new Error(
        `${EXECUTION_METADATA_SHEET_NAME}!${cell.address} must be "${header}"`,
      );
    }
  });
  FILE_HEADERS.forEach((header, index) => {
    const cell = worksheet!.getCell(1, index + 8);
    if (!cell.text) {
      cell.value = header;
      styleHeaderCell(cell);
      changed = true;
    } else if (cell.text !== header) {
      throw new Error(
        `${EXECUTION_METADATA_SHEET_NAME}!${cell.address} must be "${header}"`,
      );
    }
  });
  [24, 30, 45, 28, 24, 14, 3, 55, 24, 18, 10, 28, 30, 20, 55, 24].forEach(
    (width, index) => {
      if ((worksheet!.getColumn(index + 1).width ?? 0) < width) {
        worksheet!.getColumn(index + 1).width = width;
        changed = true;
      }
    },
  );
  return { worksheet, changed };
}

export function ensureEvidenceWorkbookSchema(workbook: ExcelJS.Workbook): boolean {
  const kb = workbook.getWorksheet(KB_SHEET_NAME);
  const negative = workbook.getWorksheet(NEGATIVE_SHEET_NAME);
  const transcript = workbook.getWorksheet(TRANSCRIPT_SHEET_NAME);
  if (!kb || !negative || !transcript) {
    throw new Error("PGN workbook evidence schema requires both result sheets and transcript");
  }
  const metadata = ensureMetadataWorksheet(workbook);
  const kbChanged = ensureMainEvidenceColumn(kb);
  const negativeChanged = ensureMainEvidenceColumn(negative);
  const transcriptChanged = ensureTranscriptEvidenceColumns(transcript);
  return kbChanged || negativeChanged || transcriptChanged || metadata.changed;
}

export function readEvidenceHyperlink(cell: Cell): string | undefined {
  const value = cell.value;
  if (
    value &&
    typeof value === "object" &&
    "hyperlink" in value &&
    typeof value.hyperlink === "string"
  ) {
    return value.hyperlink;
  }
  return undefined;
}

export function writeEvidenceHyperlink(cell: Cell, url: string): boolean {
  if (readEvidenceHyperlink(cell) === url) {
    return false;
  }
  cell.value = { text: "View Evidence", hyperlink: url };
  cell.font = {
    ...cell.font,
    color: { argb: "FF0563C1" },
    underline: true,
  };
  cell.alignment = { ...cell.alignment, vertical: "top", wrapText: true };
  return true;
}

export function writeMainEvidenceHyperlink(
  workbook: ExcelJS.Workbook,
  sheetName: string,
  rowNumber: number,
  url: string,
): boolean {
  const worksheet = workbook.getWorksheet(sheetName);
  if (!worksheet) {
    throw new Error(`Worksheet "${sheetName}" was not found`);
  }
  return writeEvidenceHyperlink(
    worksheet.getCell(rowNumber, MAIN_EVIDENCE_COLUMN),
    url,
  );
}

export function getEvidenceRunMetadata(
  workbook: ExcelJS.Workbook,
  runId: string,
): EvidenceRunMetadata | undefined {
  const worksheet = workbook.getWorksheet(EXECUTION_METADATA_SHEET_NAME);
  if (!worksheet) {
    return undefined;
  }
  for (let rowNumber = 2; rowNumber <= worksheet.rowCount; rowNumber += 1) {
    const row = worksheet.getRow(rowNumber);
    if (row.getCell(1).text !== runId) {
      continue;
    }
    const folderUrl = readEvidenceHyperlink(row.getCell(3));
    const timestampValue = row.getCell(5).value;
    return {
      runId,
      folderId: row.getCell(2).text,
      folderUrl: folderUrl ?? "",
      migrationVersion: row.getCell(4).text,
      timestamp:
        timestampValue instanceof Date ? timestampValue : new Date(row.getCell(5).text),
      mode: row.getCell(6).text === "FUTURE" ? "FUTURE" : "MIGRATION",
    };
  }
  return undefined;
}

export function upsertEvidenceRunMetadata(
  workbook: ExcelJS.Workbook,
  metadata: EvidenceRunMetadata,
): void {
  const worksheet = ensureMetadataWorksheet(workbook).worksheet;
  let rowNumber = 2;
  while (rowNumber <= worksheet.rowCount && worksheet.getCell(rowNumber, 1).text) {
    if (worksheet.getCell(rowNumber, 1).text === metadata.runId) {
      break;
    }
    rowNumber += 1;
  }
  const row = worksheet.getRow(rowNumber);
  row.getCell(1).value = metadata.runId;
  row.getCell(2).value = metadata.folderId;
  writeEvidenceHyperlink(row.getCell(3), metadata.folderUrl);
  row.getCell(4).value = metadata.migrationVersion;
  row.getCell(5).value = metadata.timestamp;
  row.getCell(5).numFmt = "yyyy-mm-dd hh:mm:ss";
  row.getCell(6).value = metadata.mode;
}

export function getEvidenceFileMetadata(
  workbook: ExcelJS.Workbook,
  evidenceKey: string,
): EvidenceFileMetadata | undefined {
  const worksheet = workbook.getWorksheet(EXECUTION_METADATA_SHEET_NAME);
  if (!worksheet) {
    return undefined;
  }
  for (let rowNumber = 2; rowNumber <= worksheet.rowCount; rowNumber += 1) {
    const row = worksheet.getRow(rowNumber);
    if (row.getCell(8).text !== evidenceKey) {
      continue;
    }
    return {
      evidenceKey,
      runId: row.getCell(9).text,
      testCaseId: row.getCell(10).text,
      turnNumber: Number(row.getCell(11).value),
      driveFileId: row.getCell(12).text || undefined,
      driveFileName: row.getCell(13).text,
      evidenceUrl: readEvidenceHyperlink(row.getCell(14)),
      localCleanPath: row.getCell(15).text || undefined,
      status: (row.getCell(16).text || "EVIDENCE_PENDING") as EvidenceStatus,
    };
  }
  return undefined;
}

export function upsertEvidenceFileMetadata(
  workbook: ExcelJS.Workbook,
  metadata: EvidenceFileMetadata,
): void {
  const worksheet = ensureMetadataWorksheet(workbook).worksheet;
  let rowNumber = 2;
  while (rowNumber <= worksheet.rowCount && worksheet.getCell(rowNumber, 8).text) {
    if (worksheet.getCell(rowNumber, 8).text === metadata.evidenceKey) {
      break;
    }
    rowNumber += 1;
  }
  const row = worksheet.getRow(rowNumber);
  row.getCell(8).value = metadata.evidenceKey;
  row.getCell(9).value = metadata.runId;
  row.getCell(10).value = metadata.testCaseId;
  row.getCell(11).value = metadata.turnNumber;
  row.getCell(12).value = metadata.driveFileId ?? "";
  row.getCell(13).value = metadata.driveFileName;
  if (metadata.evidenceUrl) {
    writeEvidenceHyperlink(row.getCell(14), metadata.evidenceUrl);
  } else {
    row.getCell(14).value = "";
  }
  row.getCell(15).value = metadata.localCleanPath ?? "";
  row.getCell(16).value = metadata.status;
}
