import ExcelJS, { type Cell, type Worksheet } from "exceljs";
import {
  readEvidenceHyperlink,
  writeEvidenceHyperlink,
} from "./evidence-workbook";
import { cellText } from "./pgn-workbook-loader";
import {
  PGN_TEST_STATUSES,
  type PgnTestStatus,
} from "./pgn-test-status";
import {
  TRANSCRIPT_SHEET_NAME,
  type ExecutedTurn,
  type PgnTestScenario,
} from "./pgn-types";

export const RETEST_HISTORY_SHEET_NAME = "Retest History";
export const RETEST_METADATA_SHEET_NAME = "Retest Metadata";

const RETEST_HISTORY_HEADERS = [
  "Run ID",
  "Retest Run ID",
  "Test Case ID",
  "Sheet",
  "Excel Row",
  "Turn",
  "Previous Status",
  "Previous Bot Response",
  "Previous Response Time",
  "Previous Test Date",
  "Previous Evidence URL",
  "Retested At",
  "New Technical Status",
  "New Bot Response",
  "New Response Time",
  "New Test Date",
  "New Evidence URL",
  "History Key",
];

const RETEST_METADATA_HEADERS = [
  "Retest Run ID",
  "Started At",
  "State",
  "Selected Test IDs",
  "Successfully Completed Test IDs",
  "Last Updated",
  "Evidence Drive Folder ID",
  "Evidence Drive Folder URL",
];

export type RetestRunState = "IN_PROGRESS" | "COMPLETE";

export interface RetestRunMetadata {
  runId: string;
  startedAt: Date;
  state: RetestRunState;
  selectedIds: string[];
  finishedIds: string[];
  updatedAt: Date;
  folderId?: string;
  folderUrl?: string;
}

function styleHeader(cell: Cell): void {
  cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
  cell.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FF7030A0" },
  };
  cell.alignment = { vertical: "middle", wrapText: true };
}

function ensureSheet(
  workbook: ExcelJS.Workbook,
  name: string,
  headers: string[],
  widths: number[],
): { worksheet: Worksheet; changed: boolean } {
  let worksheet = workbook.getWorksheet(name);
  let changed = false;
  if (!worksheet) {
    worksheet = workbook.addWorksheet(name);
    worksheet.views = [{ state: "frozen", ySplit: 1 }];
    changed = true;
  }
  headers.forEach((header, index) => {
    const cell = worksheet!.getCell(1, index + 1);
    if (!cell.text) {
      cell.value = header;
      styleHeader(cell);
      changed = true;
    } else if (cell.text !== header) {
      throw new Error(`${name}!${cell.address} must be "${header}"`);
    }
    if ((worksheet!.getColumn(index + 1).width ?? 0) < widths[index]) {
      worksheet!.getColumn(index + 1).width = widths[index];
      changed = true;
    }
  });
  return { worksheet, changed };
}

function ensureRetestHistoryWorksheet(workbook: ExcelJS.Workbook): Worksheet {
  const { worksheet } = ensureSheet(
    workbook,
    RETEST_HISTORY_SHEET_NAME,
    RETEST_HISTORY_HEADERS,
    [24, 25, 18, 28, 12, 8, 22, 60, 20, 22, 45, 22, 24, 60, 20, 22, 45, 45],
  );
  worksheet.getColumn(18).hidden = true;
  return worksheet;
}

function ensureRetestMetadataWorksheet(workbook: ExcelJS.Workbook): Worksheet {
  return ensureSheet(
    workbook,
    RETEST_METADATA_SHEET_NAME,
    RETEST_METADATA_HEADERS,
    [25, 22, 16, 80, 80, 22, 30, 45],
  ).worksheet;
}

export function ensureRetestWorkbookSchema(workbook: ExcelJS.Workbook): boolean {
  const historyExists = Boolean(workbook.getWorksheet(RETEST_HISTORY_SHEET_NAME));
  const metadataExists = Boolean(workbook.getWorksheet(RETEST_METADATA_SHEET_NAME));
  ensureRetestHistoryWorksheet(workbook);
  ensureRetestMetadataWorksheet(workbook);
  return !historyExists || !metadataExists;
}

function parseIdList(value: string): string[] {
  if (!value.trim()) {
    return [];
  }
  try {
    const parsed = JSON.parse(value) as unknown;
    if (
      Array.isArray(parsed) &&
      parsed.every((item) => typeof item === "string")
    ) {
      return parsed;
    }
  } catch {
    // Invalid metadata is rejected below.
  }
  throw new Error("Retest Metadata contains an invalid Test Case ID list");
}

function readDate(cell: Cell, field: string): Date {
  const value = cell.value;
  const parsed = value instanceof Date ? value : new Date(cell.text);
  if (Number.isNaN(parsed.getTime())) {
    throw new Error(`Retest Metadata contains an invalid ${field}`);
  }
  return parsed;
}

export function getRetestRunMetadata(
  workbook: ExcelJS.Workbook,
  runId: string,
): RetestRunMetadata | undefined {
  const worksheet = workbook.getWorksheet(RETEST_METADATA_SHEET_NAME);
  if (!worksheet) {
    return undefined;
  }
  for (let rowNumber = 2; rowNumber <= worksheet.rowCount; rowNumber += 1) {
    const row = worksheet.getRow(rowNumber);
    if (row.getCell(1).text !== runId) {
      continue;
    }
    const state = row.getCell(3).text;
    if (state !== "IN_PROGRESS" && state !== "COMPLETE") {
      throw new Error(`Retest Run ${runId} has invalid state "${state}"`);
    }
    return {
      runId,
      startedAt: readDate(row.getCell(2), "Started At"),
      state,
      selectedIds: parseIdList(row.getCell(4).text),
      finishedIds: parseIdList(row.getCell(5).text),
      updatedAt: readDate(row.getCell(6), "Last Updated"),
      folderId: row.getCell(7).text || undefined,
      folderUrl: readEvidenceHyperlink(row.getCell(8)),
    };
  }
  return undefined;
}

export function upsertRetestRunMetadata(
  workbook: ExcelJS.Workbook,
  metadata: RetestRunMetadata,
): void {
  const worksheet = ensureRetestMetadataWorksheet(workbook);
  let rowNumber = worksheet.rowCount + 1;
  for (let candidate = 2; candidate <= worksheet.rowCount; candidate += 1) {
    if (worksheet.getCell(candidate, 1).text === metadata.runId) {
      rowNumber = candidate;
      break;
    }
  }
  const row = worksheet.getRow(rowNumber);
  row.getCell(1).value = metadata.runId;
  row.getCell(2).value = metadata.startedAt;
  row.getCell(2).numFmt = "yyyy-mm-dd hh:mm:ss";
  row.getCell(3).value = metadata.state;
  row.getCell(4).value = JSON.stringify(metadata.selectedIds);
  row.getCell(5).value = JSON.stringify(metadata.finishedIds);
  row.getCell(6).value = metadata.updatedAt;
  row.getCell(6).numFmt = "yyyy-mm-dd hh:mm:ss";
  row.getCell(7).value = metadata.folderId ?? "";
  if (metadata.folderUrl) {
    writeEvidenceHyperlink(row.getCell(8), metadata.folderUrl);
  } else {
    row.getCell(8).value = null;
  }
}

export function statusColumnForScenario(scenario: PgnTestScenario): number {
  return scenario.sheetKind === "kb" ? 12 : 11;
}

export function setScenarioStatus(
  workbook: ExcelJS.Workbook,
  scenario: PgnTestScenario,
  status: PgnTestStatus,
): void {
  const worksheet = workbook.getWorksheet(scenario.sheetName);
  if (!worksheet) {
    throw new Error(`Worksheet "${scenario.sheetName}" was not found`);
  }
  worksheet.getCell(
    scenario.sourceRowNumber,
    statusColumnForScenario(scenario),
  ).value = status;
}

export function applyRetestStatusTransition(
  workbook: ExcelJS.Workbook,
  scenario: PgnTestScenario,
  executions: ExecutedTurn[],
): boolean {
  const allTurnsCaptured =
    executions.length === scenario.turns.length &&
    executions.every((execution) => execution.technicalStatus === "CAPTURED");
  if (allTurnsCaptured) {
    setScenarioStatus(
      workbook,
      scenario,
      PGN_TEST_STATUSES.PendingEvaluation,
    );
  }
  return allTurnsCaptured;
}

function resultColumns(scenario: PgnTestScenario): {
  response: number;
  responseTime: number;
  testDate: number;
  evidence: number;
} {
  return scenario.sheetKind === "kb"
    ? { response: 9, responseTime: 10, testDate: 11, evidence: 14 }
    : { response: 8, responseTime: 9, testDate: 10, evidence: 14 };
}

function historyKey(
  runId: string,
  scenario: PgnTestScenario,
  turnNumber?: number,
): string {
  return `${runId}|${scenario.testCaseId}|${turnNumber ?? 0}`;
}

function findHistoryRow(worksheet: Worksheet, key: string): number | undefined {
  for (let rowNumber = 2; rowNumber <= worksheet.rowCount; rowNumber += 1) {
    if (worksheet.getCell(rowNumber, 18).text === key) {
      return rowNumber;
    }
  }
  return undefined;
}

function previousRunId(
  workbook: ExcelJS.Workbook,
  scenario: PgnTestScenario,
  turnNumber?: number,
): string {
  const transcript = workbook.getWorksheet(TRANSCRIPT_SHEET_NAME);
  if (!transcript) {
    return "";
  }
  for (let rowNumber = transcript.rowCount; rowNumber >= 2; rowNumber -= 1) {
    const row = transcript.getRow(rowNumber);
    if (
      row.getCell(2).text === scenario.testCaseId &&
      (!turnNumber || Number(row.getCell(5).value) === turnNumber)
    ) {
      return row.getCell(1).text;
    }
  }
  return "";
}

function historyTargets(
  scenario: PgnTestScenario,
): Array<{ rowNumber: number; turnNumber?: number }> {
  return scenario.sheetKind === "kb"
    ? scenario.turns.map((turn) => ({
        rowNumber: turn.rowNumber,
        turnNumber: turn.turnNumber,
      }))
    : [{ rowNumber: scenario.sourceRowNumber }];
}

export function snapshotRetestHistory(
  workbook: ExcelJS.Workbook,
  retestRunId: string,
  scenario: PgnTestScenario,
  retestedAt: Date,
): void {
  const history = ensureRetestHistoryWorksheet(workbook);
  const results = workbook.getWorksheet(scenario.sheetName);
  if (!results) {
    throw new Error(`Worksheet "${scenario.sheetName}" was not found`);
  }
  const columns = resultColumns(scenario);
  const previousStatus = cellText(
    results.getCell(
      scenario.sourceRowNumber,
      statusColumnForScenario(scenario),
    ),
  ).trim();
  for (const target of historyTargets(scenario)) {
    const key = historyKey(retestRunId, scenario, target.turnNumber);
    if (findHistoryRow(history, key)) {
      continue;
    }
    const sourceRow = results.getRow(target.rowNumber);
    const row = history.addRow([
      previousRunId(workbook, scenario, target.turnNumber),
      retestRunId,
      scenario.testCaseId,
      scenario.sheetName,
      target.rowNumber,
      target.turnNumber ?? null,
      previousStatus,
      sourceRow.getCell(columns.response).value,
      sourceRow.getCell(columns.responseTime).value,
      sourceRow.getCell(columns.testDate).value,
      "",
      retestedAt,
      "",
      "",
      "",
      "",
      "",
      key,
    ]);
    const previousEvidence = readEvidenceHyperlink(
      sourceRow.getCell(columns.evidence),
    );
    if (previousEvidence) {
      writeEvidenceHyperlink(row.getCell(11), previousEvidence);
    }
    row.getCell(10).numFmt = "yyyy-mm-dd hh:mm:ss";
    row.getCell(12).numFmt = "yyyy-mm-dd hh:mm:ss";
    row.alignment = { vertical: "top", wrapText: true };
  }
}

function setHistoryLink(cell: Cell, url?: string): boolean {
  if (url) {
    return writeEvidenceHyperlink(cell, url);
  }
  if (cell.value) {
    cell.value = null;
    return true;
  }
  return false;
}

export function updateRetestHistoryEvidence(
  workbook: ExcelJS.Workbook,
  retestRunId: string,
  testCaseId: string,
  turnNumber: number | undefined,
  url?: string,
): boolean {
  const history = workbook.getWorksheet(RETEST_HISTORY_SHEET_NAME);
  if (!history) {
    return false;
  }
  const rowNumber = findHistoryRow(
    history,
    `${retestRunId}|${testCaseId}|${turnNumber ?? 0}`,
  );
  return rowNumber ? setHistoryLink(history.getCell(rowNumber, 17), url) : false;
}

export function updateRetestHistory(
  workbook: ExcelJS.Workbook,
  retestRunId: string,
  scenario: PgnTestScenario,
  executions: ExecutedTurn[],
): void {
  const history = ensureRetestHistoryWorksheet(workbook);
  const results = workbook.getWorksheet(scenario.sheetName);
  if (!results) {
    throw new Error(`Worksheet "${scenario.sheetName}" was not found`);
  }
  const columns = resultColumns(scenario);
  for (const target of historyTargets(scenario)) {
    const rowNumber = findHistoryRow(
      history,
      historyKey(retestRunId, scenario, target.turnNumber),
    );
    if (!rowNumber) {
      throw new Error(
        `Retest History snapshot is missing for ${scenario.testCaseId}`,
      );
    }
    const historyRow = history.getRow(rowNumber);
    const sourceRow = results.getRow(target.rowNumber);
    const relevantExecutions = target.turnNumber
      ? executions.filter(
          (execution) => execution.turn.turnNumber === target.turnNumber,
        )
      : executions;
    if (!relevantExecutions.length) {
      continue;
    }
    historyRow.getCell(13).value = relevantExecutions
      .map(
        (execution) =>
          `${target.turnNumber ? "" : `Turn ${execution.turn.turnNumber}: `}${execution.technicalStatus}`,
      )
      .join("\n");
    historyRow.getCell(14).value = sourceRow.getCell(columns.response).value;
    historyRow.getCell(15).value = sourceRow.getCell(columns.responseTime).value;
    historyRow.getCell(16).value = sourceRow.getCell(columns.testDate).value;
    historyRow.getCell(16).numFmt = "yyyy-mm-dd hh:mm:ss";
    setHistoryLink(
      historyRow.getCell(17),
      readEvidenceHyperlink(sourceRow.getCell(columns.evidence)),
    );
  }
}
