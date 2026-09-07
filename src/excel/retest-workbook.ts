import type { Workbook, Cell, Worksheet } from "exceljs";
import type { RunExecutionContext } from "../session-mode";
import { runExecutionContext } from "./run-configuration";
import { readEvidenceHyperlink, writeEvidenceHyperlink } from "./evidence-workbook";
import { PGN_TEST_STATUSES, type PgnTestStatus } from "./pgn-test-status";
import { TRANSCRIPT_SHEET_NAME, type ExecutedTurn, type PgnTestScenario } from "./pgn-types";
import {
  RETEST_HISTORY_SCHEMA, RETEST_METADATA_SCHEMA, ensureOwnedWorksheet,
  fieldCell, fieldColumn, optionalFieldCell, appendSchemaRow,
} from "./workbook-schema";

export const RETEST_HISTORY_SHEET_NAME = "Retest History";
export const RETEST_METADATA_SHEET_NAME = "Retest Metadata";
export type RetestRunState = "IN_PROGRESS" | "COMPLETE";
export interface RetestRunMetadata extends Partial<RunExecutionContext> {
  runId: string; startedAt: Date; state: RetestRunState; selectedIds: string[];
  finishedIds: string[]; updatedAt: Date; folderId?: string; folderUrl?: string;
}

function ensureRetestHistoryWorksheet(workbook: Workbook): Worksheet {
  const worksheet = ensureOwnedWorksheet(workbook, RETEST_HISTORY_SCHEMA);
  worksheet.getColumn(fieldColumn(worksheet, "historyKey")).hidden = true;
  return worksheet;
}
export function ensureRetestWorkbookSchema(workbook: Workbook): boolean {
  const changed = !workbook.getWorksheet(RETEST_HISTORY_SHEET_NAME) || !workbook.getWorksheet(RETEST_METADATA_SHEET_NAME);
  ensureRetestHistoryWorksheet(workbook);
  ensureOwnedWorksheet(workbook, RETEST_METADATA_SCHEMA);
  return changed;
}
function parseIdList(value: string): string[] {
  if (!value.trim()) return [];
  try {
    const parsed: unknown = JSON.parse(value);
    if (Array.isArray(parsed) && parsed.every((item) => typeof item === "string")) return parsed;
  } catch { /* Invalid persisted metadata must not be guessed. */ }
  throw new Error("Retest Metadata contains an invalid Test Case ID list");
}
function readDate(cell: Cell, field: string): Date {
  const parsed = cell.value instanceof Date ? cell.value : new Date(cell.text);
  if (Number.isNaN(parsed.getTime())) throw new Error(`Retest Metadata contains an invalid ${field}`);
  return parsed;
}
export function getRetestRunMetadata(workbook: Workbook, runId: string): RetestRunMetadata | undefined {
  const sheet = workbook.getWorksheet(RETEST_METADATA_SHEET_NAME);
  if (!sheet) return undefined;
  for (let row = 2; row <= sheet.rowCount; row += 1) {
    const cell = (field: Parameters<typeof fieldCell>[2]) => fieldCell(sheet, row, field);
    if (cell("runId").text !== runId) continue;
    const state = cell("state").text;
    if (state !== "IN_PROGRESS" && state !== "COMPLETE") throw new Error(`Retest Run ${runId} has invalid state "${state}"`);
    return {
      ...runExecutionContext(workbook, runId),
      runId, startedAt: readDate(cell("startedAt"), "Started At"), state,
      selectedIds: parseIdList(cell("selectedIds").text), finishedIds: parseIdList(cell("finishedIds").text),
      updatedAt: readDate(cell("updatedAt"), "Last Updated"),
      folderId: cell("folderId").text || undefined, folderUrl: readEvidenceHyperlink(cell("folderUrl")),
    };
  }
  return undefined;
}
export function upsertRetestRunMetadata(workbook: Workbook, metadata: RetestRunMetadata): void {
  const sheet = ensureOwnedWorksheet(workbook, RETEST_METADATA_SCHEMA);
  let row = sheet.rowCount + 1;
  for (let candidate = 2; candidate <= sheet.rowCount; candidate += 1) {
    if (fieldCell(sheet, candidate, "runId").text === metadata.runId) { row = candidate; break; }
  }
  const cell = (field: Parameters<typeof fieldCell>[2]) => fieldCell(sheet, row, field);
  cell("runId").value = metadata.runId;
  cell("startedAt").value = metadata.startedAt;
  cell("startedAt").numFmt = "yyyy-mm-dd hh:mm:ss";
  cell("state").value = metadata.state;
  cell("selectedIds").value = JSON.stringify(metadata.selectedIds);
  cell("finishedIds").value = JSON.stringify(metadata.finishedIds);
  cell("updatedAt").value = metadata.updatedAt;
  cell("updatedAt").numFmt = "yyyy-mm-dd hh:mm:ss";
  cell("folderId").value = metadata.folderId ?? "";
  if (metadata.folderUrl) writeEvidenceHyperlink(cell("folderUrl"), metadata.folderUrl);
  else cell("folderUrl").value = null;
}
export function setScenarioStatus(workbook: Workbook, scenario: PgnTestScenario, status: PgnTestStatus): void {
  const sheet = workbook.getWorksheet(scenario.sheetName);
  if (!sheet) throw new Error(`Worksheet "${scenario.sheetName}" was not found`);
  fieldCell(sheet, scenario.sourceRowNumber, "status").value = status;
}
export function applyRetestStatusTransition(workbook: Workbook, scenario: PgnTestScenario, executions: ExecutedTurn[]): boolean {
  const captured = executions.length === scenario.turns.length && executions.every((execution) => execution.technicalStatus === "CAPTURED");
  if (captured) setScenarioStatus(workbook, scenario, PGN_TEST_STATUSES.PendingEvaluation);
  return captured;
}
function historyKey(runId: string, scenario: PgnTestScenario, turnNumber?: number): string {
  return `${runId}|${scenario.testCaseId}|${turnNumber ?? 0}`;
}
function findHistoryRow(sheet: Worksheet, key: string): number | undefined {
  for (let row = 2; row <= sheet.rowCount; row += 1) {
    if (fieldCell(sheet, row, "historyKey").text === key) return row;
  }
  return undefined;
}
function previousRunId(workbook: Workbook, scenario: PgnTestScenario, turnNumber?: number): string {
  const sheet = workbook.getWorksheet(TRANSCRIPT_SHEET_NAME);
  if (!sheet) return "";
  for (let row = sheet.rowCount; row >= 2; row -= 1) {
    if (fieldCell(sheet, row, "testCaseId").text === scenario.testCaseId && (!turnNumber || Number(fieldCell(sheet, row, "turn").value) === turnNumber)) return fieldCell(sheet, row, "runId").text;
  }
  return "";
}
function historyTargets(scenario: PgnTestScenario): Array<{ rowNumber: number; turnNumber?: number }> {
  return scenario.sheetKind === "kb" ? scenario.turns.map((turn) => ({ rowNumber: turn.rowNumber, turnNumber: turn.turnNumber })) : [{ rowNumber: scenario.sourceRowNumber }];
}
export function snapshotRetestHistory(workbook: Workbook, retestRunId: string, scenario: PgnTestScenario, retestedAt: Date): void {
  const history = ensureRetestHistoryWorksheet(workbook);
  const results = workbook.getWorksheet(scenario.sheetName);
  if (!results) throw new Error(`Worksheet "${scenario.sheetName}" was not found`);
  const previousStatus = fieldCell(results, scenario.sourceRowNumber, "status").text.trim();
  for (const target of historyTargets(scenario)) {
    const key = historyKey(retestRunId, scenario, target.turnNumber);
    if (findHistoryRow(history, key)) continue;
    const row = appendSchemaRow(history, {
      runId: previousRunId(workbook, scenario, target.turnNumber), retestRunId,
      testCaseId: scenario.testCaseId, sheet: scenario.sheetName, excelRow: target.rowNumber,
      turn: target.turnNumber ?? null, previousStatus,
      previousBotResponse: fieldCell(results, target.rowNumber, "botResponse").value,
      previousResponseTime: optionalFieldCell(results, target.rowNumber, "responseTime")?.value ?? null,
      previousTestDate: optionalFieldCell(results, target.rowNumber, "testDate")?.value ?? null,
      retestedAt, historyKey: key,
    });
    const url = readEvidenceHyperlink(optionalFieldCell(results, target.rowNumber, "evidence"));
    if (url) writeEvidenceHyperlink(fieldCell(history, row.number, "previousEvidenceUrl"), url);
    fieldCell(history, row.number, "previousTestDate").numFmt = "yyyy-mm-dd hh:mm:ss";
    fieldCell(history, row.number, "retestedAt").numFmt = "yyyy-mm-dd hh:mm:ss";
    row.alignment = { vertical: "top", wrapText: true };
  }
}
function setHistoryLink(cell: Cell, url?: string): boolean {
  if (url) return writeEvidenceHyperlink(cell, url);
  if (!cell.value) return false;
  cell.value = null;
  return true;
}
export function updateRetestHistoryEvidence(workbook: Workbook, retestRunId: string, testCaseId: string, turnNumber: number | undefined, url?: string): boolean {
  const sheet = workbook.getWorksheet(RETEST_HISTORY_SHEET_NAME);
  if (!sheet) return false;
  const row = findHistoryRow(sheet, `${retestRunId}|${testCaseId}|${turnNumber ?? 0}`);
  return row ? setHistoryLink(fieldCell(sheet, row, "newEvidenceUrl"), url) : false;
}
export function updateRetestHistory(workbook: Workbook, retestRunId: string, scenario: PgnTestScenario, executions: ExecutedTurn[]): void {
  const history = ensureRetestHistoryWorksheet(workbook);
  const results = workbook.getWorksheet(scenario.sheetName);
  if (!results) throw new Error(`Worksheet "${scenario.sheetName}" was not found`);
  for (const target of historyTargets(scenario)) {
    const row = findHistoryRow(history, historyKey(retestRunId, scenario, target.turnNumber));
    if (!row) throw new Error(`Retest History snapshot is missing for ${scenario.testCaseId}`);
    const relevant = target.turnNumber ? executions.filter((execution) => execution.turn.turnNumber === target.turnNumber) : executions;
    if (!relevant.length) continue;
    fieldCell(history, row, "newTechnicalStatus").value = relevant.map((execution) => `${target.turnNumber ? "" : `Turn ${execution.turn.turnNumber}: `}${execution.technicalStatus}`).join("\n");
    fieldCell(history, row, "newBotResponse").value = fieldCell(results, target.rowNumber, "botResponse").value;
    fieldCell(history, row, "newResponseTime").value = optionalFieldCell(results, target.rowNumber, "responseTime")?.value ?? null;
    fieldCell(history, row, "newTestDate").value = optionalFieldCell(results, target.rowNumber, "testDate")?.value ?? null;
    fieldCell(history, row, "newTestDate").numFmt = "yyyy-mm-dd hh:mm:ss";
    setHistoryLink(fieldCell(history, row, "newEvidenceUrl"), readEvidenceHyperlink(optionalFieldCell(results, target.rowNumber, "evidence")));
  }
}
