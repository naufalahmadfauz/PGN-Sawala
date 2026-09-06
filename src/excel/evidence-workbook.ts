import type { Workbook, Cell, Worksheet } from "exceljs";
import { KB_SHEET_NAME, NEGATIVE_SHEET_NAME, TRANSCRIPT_SHEET_NAME, type EvidenceStatus } from "./pgn-types";
import { EVIDENCE_RUN_SCHEMA, EVIDENCE_FILE_SCHEMA, ensureEvidenceField, fieldCell, getWorksheetSchema } from "./workbook-schema";

export const EXECUTION_METADATA_SHEET_NAME = "Execution Metadata";
export const EVIDENCE_MIGRATION_VERSION = "1";
export interface EvidenceRunMetadata {
  runId: string; folderId: string; folderUrl: string; migrationVersion: string;
  timestamp: Date; mode: "MIGRATION" | "FUTURE" | "RETEST";
}
export interface EvidenceFileMetadata {
  evidenceKey: string; runId: string; testCaseId: string; turnNumber: number;
  driveFileId?: string; driveFileName: string; evidenceUrl?: string; localCleanPath?: string;
  status: EvidenceStatus;
}

function ensureMetadataWorksheet(workbook: Workbook): Worksheet {
  let worksheet = workbook.getWorksheet(EXECUTION_METADATA_SHEET_NAME);
  if (!worksheet) {
    worksheet = workbook.addWorksheet(EXECUTION_METADATA_SHEET_NAME);
    worksheet.addRow([
      ...EVIDENCE_RUN_SCHEMA.fields.map((item) => item.header), null,
      ...EVIDENCE_FILE_SCHEMA.fields.map((item) => item.header),
    ]);
    worksheet.getRow(1).font = { bold: true };
    worksheet.views = [{ state: "frozen", ySplit: 1 }];
    worksheet.columns.forEach((column) => { column.width = 24; });
  }
  getWorksheetSchema(worksheet, EVIDENCE_RUN_SCHEMA);
  getWorksheetSchema(worksheet, EVIDENCE_FILE_SCHEMA);
  return worksheet;
}

export function ensureEvidenceWorkbookSchema(workbook: Workbook): boolean {
  const sheets = [KB_SHEET_NAME, NEGATIVE_SHEET_NAME, TRANSCRIPT_SHEET_NAME].map((name) => workbook.getWorksheet(name));
  if (sheets.some((sheet) => !sheet)) throw new Error("PGN workbook evidence schema requires both result sheets and transcript");
  sheets.forEach((sheet) => getWorksheetSchema(sheet!));
  let changed = !workbook.getWorksheet(EXECUTION_METADATA_SHEET_NAME);
  ensureMetadataWorksheet(workbook);
  for (const worksheet of sheets) {
    const names = worksheet!.name === TRANSCRIPT_SHEET_NAME ? ["evidenceUrl", "evidenceStatus"] as const : ["evidence"] as const;
    for (const name of names) changed = ensureEvidenceField(worksheet!, name) || changed;
  }
  return changed;
}

export function readEvidenceHyperlink(cell: Cell | undefined): string | undefined {
  const value = cell?.value;
  return value && typeof value === "object" && "hyperlink" in value && typeof value.hyperlink === "string" ? value.hyperlink : undefined;
}
export function writeEvidenceHyperlink(cell: Cell, url: string): boolean {
  if (readEvidenceHyperlink(cell) === url) return false;
  cell.value = { text: "View Evidence", hyperlink: url };
  cell.style = structuredClone(cell.style);
  cell.font = { ...cell.font, color: { argb: "FF0563C1" }, underline: true };
  cell.alignment = { ...cell.alignment, vertical: "top", wrapText: true };
  return true;
}
export function writeMainEvidenceHyperlink(workbook: Workbook, sheetName: string, rowNumber: number, url: string): boolean {
  const worksheet = workbook.getWorksheet(sheetName);
  if (!worksheet) throw new Error(`Worksheet "${sheetName}" was not found`);
  return writeEvidenceHyperlink(fieldCell(worksheet, rowNumber, "evidence"), url);
}

export function getEvidenceRunMetadata(workbook: Workbook, runId: string): EvidenceRunMetadata | undefined {
  const sheet = workbook.getWorksheet(EXECUTION_METADATA_SHEET_NAME);
  if (!sheet) return undefined;
  for (let row = 2; row <= sheet.rowCount; row += 1) {
    const cell = (field: Parameters<typeof fieldCell>[2]) => fieldCell(sheet, row, field, EVIDENCE_RUN_SCHEMA);
    if (cell("runId").text !== runId) continue;
    const date = cell("timestamp");
    const mode = cell("mode").text;
    return {
      runId, folderId: cell("folderId").text, folderUrl: readEvidenceHyperlink(cell("folderUrl")) ?? "",
      migrationVersion: cell("migrationVersion").text,
      timestamp: date.value instanceof Date ? date.value : new Date(date.text),
      mode: mode === "FUTURE" || mode === "RETEST" ? mode : "MIGRATION",
    };
  }
  return undefined;
}
export function upsertEvidenceRunMetadata(workbook: Workbook, metadata: EvidenceRunMetadata): void {
  const sheet = ensureMetadataWorksheet(workbook);
  let row = 2;
  while (row <= sheet.rowCount && fieldCell(sheet, row, "runId", EVIDENCE_RUN_SCHEMA).text) {
    if (fieldCell(sheet, row, "runId", EVIDENCE_RUN_SCHEMA).text === metadata.runId) break;
    row += 1;
  }
  const cell = (field: Parameters<typeof fieldCell>[2]) => fieldCell(sheet, row, field, EVIDENCE_RUN_SCHEMA);
  cell("runId").value = metadata.runId;
  cell("folderId").value = metadata.folderId;
  if (metadata.folderUrl) writeEvidenceHyperlink(cell("folderUrl"), metadata.folderUrl);
  else cell("folderUrl").value = null;
  cell("migrationVersion").value = metadata.migrationVersion;
  cell("timestamp").value = metadata.timestamp;
  cell("timestamp").numFmt = "yyyy-mm-dd hh:mm:ss";
  cell("mode").value = metadata.mode;
}
export function getEvidenceFileMetadata(workbook: Workbook, evidenceKey: string): EvidenceFileMetadata | undefined {
  const sheet = workbook.getWorksheet(EXECUTION_METADATA_SHEET_NAME);
  if (!sheet) return undefined;
  for (let row = 2; row <= sheet.rowCount; row += 1) {
    const cell = (field: Parameters<typeof fieldCell>[2]) => fieldCell(sheet, row, field, EVIDENCE_FILE_SCHEMA);
    if (cell("evidenceKey").text !== evidenceKey) continue;
    return {
      evidenceKey, runId: cell("runId").text, testCaseId: cell("testCaseId").text,
      turnNumber: Number(cell("turn").value), driveFileId: cell("driveFileId").text || undefined,
      driveFileName: cell("driveFileName").text, evidenceUrl: readEvidenceHyperlink(cell("evidenceUrl")),
      localCleanPath: cell("localCleanPath").text || undefined,
      status: (cell("evidenceStatus").text || "EVIDENCE_PENDING") as EvidenceStatus,
    };
  }
  return undefined;
}
export function upsertEvidenceFileMetadata(workbook: Workbook, metadata: EvidenceFileMetadata): void {
  const sheet = ensureMetadataWorksheet(workbook);
  let row: number | undefined;
  let empty: number | undefined;
  for (let candidate = 2; candidate <= sheet.rowCount; candidate += 1) {
    const key = fieldCell(sheet, candidate, "evidenceKey", EVIDENCE_FILE_SCHEMA).text;
    if (!key && empty === undefined) empty = candidate;
    if (key === metadata.evidenceKey) { row = candidate; break; }
  }
  row ??= empty ?? sheet.rowCount + 1;
  const cell = (field: Parameters<typeof fieldCell>[2]) => fieldCell(sheet, row!, field, EVIDENCE_FILE_SCHEMA);
  cell("evidenceKey").value = metadata.evidenceKey;
  cell("runId").value = metadata.runId;
  cell("testCaseId").value = metadata.testCaseId;
  cell("turn").value = metadata.turnNumber;
  cell("driveFileId").value = metadata.driveFileId ?? "";
  cell("driveFileName").value = metadata.driveFileName;
  if (metadata.evidenceUrl) writeEvidenceHyperlink(cell("evidenceUrl"), metadata.evidenceUrl);
  else cell("evidenceUrl").value = "";
  cell("localCleanPath").value = metadata.localCleanPath ?? "";
  cell("evidenceStatus").value = metadata.status;
}
export function removeEvidenceFileMetadata(workbook: Workbook, evidenceKey: string): boolean {
  const sheet = workbook.getWorksheet(EXECUTION_METADATA_SHEET_NAME);
  if (!sheet) return false;
  for (let row = 2; row <= sheet.rowCount; row += 1) {
    if (fieldCell(sheet, row, "evidenceKey", EVIDENCE_FILE_SCHEMA).text !== evidenceKey) continue;
    for (const item of EVIDENCE_FILE_SCHEMA.fields) fieldCell(sheet, row, item.field, EVIDENCE_FILE_SCHEMA).value = null;
    return true;
  }
  return false;
}
