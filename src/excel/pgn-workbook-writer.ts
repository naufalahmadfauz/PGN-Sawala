import {
  constants as fsConstants,
  access,
  copyFile,
  mkdir,
  rename,
  rm,
} from "node:fs/promises";
import path from "node:path";
import ExcelJS, { type Cell, type Worksheet } from "exceljs";
import {
  cellText,
  parsePgnWorkbook,
} from "./pgn-workbook-loader";
import {
  TRANSCRIPT_SHEET_NAME,
  type ExecutedTurn,
  type PgnTestScenario,
  type PgnWorkbookDocument,
} from "./pgn-types";
import type { BotSessionResetAttempt } from "../types";

const TRANSCRIPT_HEADERS = [
  "Run ID",
  "Test Case ID",
  "Sheet",
  "Excel Row",
  "Turn",
  "Role",
  "Message",
  "Timestamp",
  "First Response (ms)",
  "Total Response (ms)",
  "Status",
  "Error",
  "Evidence Path",
];

function ensureTranscriptWorksheet(workbook: ExcelJS.Workbook): Worksheet {
  const existing = workbook.getWorksheet(TRANSCRIPT_SHEET_NAME);
  if (existing) {
    existing.getColumn(6).width = Math.max(
      existing.getColumn(6).width ?? 0,
      16,
    );
    return existing;
  }

  const worksheet = workbook.addWorksheet(TRANSCRIPT_SHEET_NAME);
  const header = worksheet.addRow(TRANSCRIPT_HEADERS);
  header.font = { bold: true, color: { argb: "FFFFFFFF" } };
  header.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FF1F4E78" },
  };
  header.alignment = { vertical: "middle", wrapText: true };
  worksheet.views = [{ state: "frozen", ySplit: 1 }];
  const widths = [
    24, 18, 28, 12, 9, 16, 70, 24, 20, 20, 18, 45, 55,
  ];
  widths.forEach((width, index) => {
    worksheet.getColumn(index + 1).width = width;
  });
  return worksheet;
}

function appendTechnicalNote(cell: Cell, note: string): void {
  const existing = cellText(cell).trim();
  cell.value = existing ? `${existing}\n${note}` : note;
}

function writeExecutionDate(cell: Cell, value: Date): void {
  cell.value = value;
  cell.numFmt = "yyyy-mm-dd hh:mm:ss";
}

function secondsFromMilliseconds(milliseconds: number): number {
  return Math.round((milliseconds / 1_000) * 100) / 100;
}

function appendTranscriptRows(
  worksheet: Worksheet,
  runId: string,
  scenario: PgnTestScenario,
  execution: ExecutedTurn,
): void {
  const common = [
    runId,
    scenario.testCaseId,
    scenario.sheetName,
    execution.turn.rowNumber,
    execution.turn.turnNumber,
  ];
  const userRow = worksheet.addRow([
    ...common,
    "USER",
    execution.turn.userInput,
    execution.sentAt ?? execution.completedAt,
    null,
    null,
    execution.technicalStatus,
    execution.error ?? "",
    execution.evidencePath ?? "",
  ]);
  userRow.alignment = { vertical: "top", wrapText: true };
  userRow.getCell(8).numFmt = "yyyy-mm-dd hh:mm:ss";

  if (execution.botMessages.length > 0) {
    for (const botMessage of execution.botMessages) {
      const botRow = worksheet.addRow([
        ...common,
        "BOT",
        botMessage.message,
        botMessage.timestamp,
        execution.firstResponseMs ?? null,
        execution.totalResponseMs ?? null,
        execution.technicalStatus,
        execution.error ?? "",
        execution.evidencePath ?? "",
      ]);
      botRow.alignment = { vertical: "top", wrapText: true };
      botRow.getCell(8).numFmt = "yyyy-mm-dd hh:mm:ss";
    }
  } else if (execution.error) {
    const errorRow = worksheet.addRow([
      ...common,
      "SYSTEM",
      execution.error,
      execution.completedAt,
      execution.firstResponseMs ?? null,
      execution.totalResponseMs ?? null,
      execution.technicalStatus,
      execution.error,
      execution.evidencePath ?? "",
    ]);
    errorRow.alignment = { vertical: "top", wrapText: true };
    errorRow.getCell(8).numFmt = "yyyy-mm-dd hh:mm:ss";
  }
}

function applyKnowledgeBaseExecution(
  worksheet: Worksheet,
  executions: ExecutedTurn[],
): void {
  for (const execution of executions) {
    const row = execution.turn.rowNumber;
    if (execution.combinedResponse) {
      worksheet.getCell(row, 9).value = execution.combinedResponse;
    }
    if (execution.totalResponseMs !== undefined) {
      const responseTime = worksheet.getCell(row, 10);
      responseTime.value = secondsFromMilliseconds(execution.totalResponseMs);
      responseTime.numFmt = '0.00" s"';
    }
    writeExecutionDate(worksheet.getCell(row, 11), execution.completedAt);
    if (execution.technicalStatus !== "CAPTURED") {
      appendTechnicalNote(
        worksheet.getCell(row, 13),
        `[Technical execution ${execution.completedAt.toISOString()}] Turn ${execution.turn.turnNumber}: ${execution.technicalStatus}${execution.error ? ` - ${execution.error}` : ""}`,
      );
    }
  }
}

function applyNegativeExecution(
  worksheet: Worksheet,
  scenario: PgnTestScenario,
  executions: ExecutedTurn[],
): void {
  const row = scenario.sourceRowNumber;
  const isMultiTurn = scenario.turns.length > 1;
  const allTurnsExecuted = executions.length === scenario.turns.length;
  const allTurnsHaveResponses =
    allTurnsExecuted &&
    executions.every((execution) => Boolean(execution.combinedResponse));
  if (allTurnsHaveResponses) {
    worksheet.getCell(row, 8).value = isMultiTurn
      ? executions
          .map(
            (execution) =>
              `Turn ${execution.turn.turnNumber}:\n${execution.combinedResponse}`,
          )
          .join("\n\n")
      : executions[0].combinedResponse;
  }

  const allTurnsHaveTiming =
    allTurnsExecuted &&
    executions.every((execution) => execution.totalResponseMs !== undefined);
  if (allTurnsHaveTiming) {
    const responseTime = worksheet.getCell(row, 9);
    if (isMultiTurn) {
      responseTime.value = executions
        .map(
          (execution) =>
            `Turn ${execution.turn.turnNumber}: ${execution.totalResponseMs} ms`,
        )
        .join("\n");
      responseTime.numFmt = "@";
    } else {
      responseTime.value = secondsFromMilliseconds(
        executions[0].totalResponseMs!,
      );
      responseTime.numFmt = "0.00";
    }
  }

  const completedAt = executions.at(-1)?.completedAt ?? new Date();
  writeExecutionDate(worksheet.getCell(row, 10), completedAt);
  for (const execution of executions) {
    if (execution.technicalStatus !== "CAPTURED") {
      appendTechnicalNote(
        worksheet.getCell(row, 12),
        `[Technical execution ${execution.completedAt.toISOString()}] Turn ${execution.turn.turnNumber}: ${execution.technicalStatus}${execution.error ? ` - ${execution.error}` : ""}`,
      );
    }
  }
}

export async function openExecutedPgnWorkbook(
  sourcePath: string,
  outputPath: string,
): Promise<PgnWorkbookDocument & { resumed: boolean }> {
  await mkdir(path.dirname(outputPath), { recursive: true });
  const resumed = await access(outputPath)
    .then(() => true)
    .catch(() => false);
  if (!resumed) {
    await copyFile(sourcePath, outputPath, fsConstants.COPYFILE_EXCL);
  }

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(outputPath);
  const hadTranscript = Boolean(
    workbook.getWorksheet(TRANSCRIPT_SHEET_NAME),
  );
  ensureTranscriptWorksheet(workbook);
  if (!hadTranscript) {
    await saveExecutedPgnWorkbook(workbook, outputPath);
  }
  return { workbook, parsed: parsePgnWorkbook(workbook), resumed };
}

export function applyScenarioExecution(
  workbook: ExcelJS.Workbook,
  runId: string,
  scenario: PgnTestScenario,
  executions: ExecutedTurn[],
): void {
  const worksheet = workbook.getWorksheet(scenario.sheetName);
  if (!worksheet) {
    throw new Error(`Worksheet "${scenario.sheetName}" was not found`);
  }
  if (scenario.sheetKind === "kb") {
    applyKnowledgeBaseExecution(worksheet, executions);
  } else {
    applyNegativeExecution(worksheet, scenario, executions);
  }

  const transcript = ensureTranscriptWorksheet(workbook);
  for (const execution of executions) {
    appendTranscriptRows(transcript, runId, scenario, execution);
  }
}

export function appendSessionResetTranscript(
  workbook: ExcelJS.Workbook,
  runId: string,
  scenario: PgnTestScenario,
  attempt: BotSessionResetAttempt,
): void {
  const worksheet = ensureTranscriptWorksheet(workbook);
  const common = [
    runId,
    scenario.testCaseId,
    scenario.sheetName,
    scenario.sourceRowNumber,
    null,
  ];
  const appendRow = (
    role: "CONTROL_USER" | "CONTROL_BOT" | "CONTROL_SYSTEM",
    message: string,
    timestamp: Date,
    firstResponseMs?: number,
    totalResponseMs?: number,
  ): void => {
    const row = worksheet.addRow([
      ...common,
      role,
      message,
      timestamp,
      firstResponseMs ?? null,
      totalResponseMs ?? null,
      attempt.status,
      attempt.error ?? "",
      attempt.evidencePath ?? "",
    ]);
    row.alignment = { vertical: "top", wrapText: true };
    row.getCell(8).numFmt = "yyyy-mm-dd hh:mm:ss";
  };

  if (attempt.sentAt) {
    appendRow("CONTROL_USER", attempt.command, attempt.sentAt);
  }
  for (const response of attempt.responseMessages) {
    appendRow(
      "CONTROL_BOT",
      response.text,
      response.observedAt,
      attempt.firstResponseMs,
      attempt.totalResponseMs,
    );
  }
  if (attempt.status === "RESET_FAILED") {
    appendRow(
      "CONTROL_SYSTEM",
      attempt.error ?? "PGN bot session reset failed",
      attempt.completedAt,
      attempt.firstResponseMs,
      attempt.totalResponseMs,
    );
  }
}

export async function saveExecutedPgnWorkbook(
  workbook: ExcelJS.Workbook,
  outputPath: string,
): Promise<void> {
  const temporaryPath = `${outputPath}.${process.pid}.tmp`;
  try {
    await workbook.xlsx.writeFile(temporaryPath);
    await rename(temporaryPath, outputPath);
  } catch (error) {
    await rm(temporaryPath, { force: true }).catch(() => undefined);
    throw error;
  }
}
