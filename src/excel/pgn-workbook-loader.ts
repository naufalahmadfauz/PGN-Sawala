import { access } from "node:fs/promises";
import ExcelJS, { type Cell, type Worksheet } from "exceljs";
import { parseTurnsFromCell } from "./multi-turn-parser";
import { normalizeTestStatus } from "./pgn-test-status";
import { attachWorkbookMappings } from "./workbook-mapping";
import {
  KB_SCHEMA, NEGATIVE_SCHEMA, WORKBOOK_SCHEMAS, fieldCell, optionalFieldCell,
  resolveWorksheetSchema, mainSchemaFingerprint,
} from "./workbook-schema";
export { KB_HEADERS, NEGATIVE_HEADERS } from "./workbook-schema";
import {
  KB_SHEET_NAME,
  NEGATIVE_SHEET_NAME,
  type ParsedPgnWorkbook,
  type PgnSheetKind,
  type PgnSheetSummary,
  type PgnTestScenario,
  type PgnValidationIssue,
  type PgnWorkbookDocument,
} from "./pgn-types";

export function cellText(cell: Cell): string {
  return String(cell.text ?? "").replace(/\r\n/g, "\n");
}

function parseTurnNumber(
  value: string,
  sheetName: string,
  rowNumber: number,
  issues: PgnValidationIssue[],
): number | undefined {
  const normalized = value.trim();
  if (!normalized) {
    return undefined;
  }
  const turnNumber = Number(normalized);
  if (!Number.isInteger(turnNumber) || turnNumber < 1) {
    issues.push({
      code: "INVALID_TURN",
      severity: "ERROR",
      sheetName,
      rowNumber,
      message: `Turn must be a positive integer; found "${normalized}".`,
    });
    return undefined;
  }
  return turnNumber;
}

function readScenarioStatus(
  worksheet: Worksheet,
  rowNumber: number,
  issues: PgnValidationIssue[],
): { rawStatus: string; status: PgnTestScenario["status"] } {
  const rawStatus = cellText(fieldCell(worksheet, rowNumber, "status")).trim();
  const status = normalizeTestStatus(rawStatus);
  if (rawStatus && !status) {
    issues.push({
      code: "UNKNOWN_STATUS",
      severity: "WARNING",
      sheetName: worksheet.name,
      rowNumber,
      message: `Unknown Status "${rawStatus}".`,
    });
  }
  return { rawStatus, status };
}

function parseKnowledgeBaseSheet(
  worksheet: Worksheet,
  issues: PgnValidationIssue[],
  testIdRows: Map<string, { sheetName: string; rowNumber: number }>,
): PgnTestScenario[] {
  const scenarios: PgnTestScenario[] = [];
  let currentScenario: PgnTestScenario | undefined;

  for (let rowNumber = 2; rowNumber <= worksheet.rowCount; rowNumber += 1) {
    const testCaseId = cellText(fieldCell(worksheet, rowNumber, "testCaseId")).trim();
    const rawTurn = cellText(fieldCell(worksheet, rowNumber, "turn"));
    const userInput = cellText(fieldCell(worksheet, rowNumber, "userInput"));
    const hasTurn = Boolean(rawTurn.trim());
    const hasInput = Boolean(userInput.trim());
    if (!testCaseId && !hasTurn && !hasInput) {
      continue;
    }

    if (testCaseId) {
      const previous = testIdRows.get(testCaseId);
      if (previous) {
        issues.push({
          code: "DUPLICATE_TEST_ID",
          severity: "ERROR",
          sheetName: worksheet.name,
          rowNumber,
          message: `Test Case ID "${testCaseId}" duplicates ${previous.sheetName} row ${previous.rowNumber}.`,
        });
      } else {
        testIdRows.set(testCaseId, { sheetName: worksheet.name, rowNumber });
      }

      const turnNumber =
        parseTurnNumber(rawTurn, worksheet.name, rowNumber, issues) ?? 1;
      const scenarioStatus = readScenarioStatus(
        worksheet,
        rowNumber,
        issues,
      );
      currentScenario = {
        testCaseId,
        sheetKind: "kb",
        sheetName: worksheet.name,
        sourceRowNumber: rowNumber,
        category: optionalFieldCell(worksheet, rowNumber, "knowledgeBaseArticle")?.text.trim() ?? "",
        schemaFingerprint: mainSchemaFingerprint(worksheet),
        ...scenarioStatus,
        turns: [],
      };
      scenarios.push(currentScenario);
      if (!hasInput) {
        issues.push({
          code: "MISSING_USER_INPUT",
          severity: "ERROR",
          sheetName: worksheet.name,
          rowNumber,
          message: "User Input is empty.",
        });
      } else {
        currentScenario.turns.push({
          sheetName: worksheet.name,
          rowNumber,
          turnNumber,
          userInput,
        });
      }
      continue;
    }

    if (!hasTurn) {
      issues.push({
        code: "INVALID_TURN_ROW",
        severity: "ERROR",
        sheetName: worksheet.name,
        rowNumber,
        message: "A row without Test Case ID must provide Turn.",
      });
      continue;
    }
    if (!currentScenario) {
      issues.push({
        code: "INVALID_TURN_ROW",
        severity: "ERROR",
        sheetName: worksheet.name,
        rowNumber,
        message: "Continuation turn has no preceding Test Case ID.",
      });
      continue;
    }

    const turnNumber = parseTurnNumber(
      rawTurn,
      worksheet.name,
      rowNumber,
      issues,
    );
    if (!hasInput) {
      issues.push({
        code: "MISSING_USER_INPUT",
        severity: "ERROR",
        sheetName: worksheet.name,
        rowNumber,
        message: "User Input is empty.",
      });
    } else if (turnNumber !== undefined) {
      currentScenario.turns.push({
        sheetName: worksheet.name,
        rowNumber,
        turnNumber,
        userInput,
      });
    }
  }

  for (const scenario of scenarios) {
    scenario.turns.forEach((turn, index) => {
      const expected = index + 1;
      if (turn.turnNumber !== expected) {
        issues.push({
          code: "INVALID_TURN",
          severity: "ERROR",
          sheetName: scenario.sheetName,
          rowNumber: turn.rowNumber,
          message: `${scenario.testCaseId} turns must be sequential from Turn 1; expected ${expected}, found ${turn.turnNumber}.`,
        });
      }
    });
  }

  return scenarios;
}

function parseNegativeSheet(
  worksheet: Worksheet,
  issues: PgnValidationIssue[],
  testIdRows: Map<string, { sheetName: string; rowNumber: number }>,
): PgnTestScenario[] {
  const scenarios: PgnTestScenario[] = [];

  for (let rowNumber = 2; rowNumber <= worksheet.rowCount; rowNumber += 1) {
    const testCaseId = cellText(fieldCell(worksheet, rowNumber, "testCaseId")).trim();
    const userInput = cellText(fieldCell(worksheet, rowNumber, "userInput"));
    if (!testCaseId && !userInput.trim()) {
      continue;
    }
    if (!testCaseId) {
      issues.push({
        code: "MISSING_TEST_CASE_ID",
        severity: "ERROR",
        sheetName: worksheet.name,
        rowNumber,
        message: "Test Case ID is empty.",
      });
      continue;
    }

    const previous = testIdRows.get(testCaseId);
    if (previous) {
      issues.push({
        code: "DUPLICATE_TEST_ID",
        severity: "ERROR",
        sheetName: worksheet.name,
        rowNumber,
        message: `Test Case ID "${testCaseId}" duplicates ${previous.sheetName} row ${previous.rowNumber}.`,
      });
    } else {
      testIdRows.set(testCaseId, { sheetName: worksheet.name, rowNumber });
    }

    if (!userInput.trim()) {
      issues.push({
        code: "MISSING_USER_INPUT",
        severity: "ERROR",
        sheetName: worksheet.name,
        rowNumber,
        message: "User Input / Test Steps is empty.",
      });
    }
    const parsedTurns = parseTurnsFromCell(
      userInput,
      worksheet.name,
      rowNumber,
    );
    issues.push(...parsedTurns.issues);
    const scenarioStatus = readScenarioStatus(
      worksheet,
      rowNumber,
      issues,
    );
    scenarios.push({
      testCaseId,
      sheetKind: "negative",
      sheetName: worksheet.name,
      sourceRowNumber: rowNumber,
      category: optionalFieldCell(worksheet, rowNumber, "category")?.text.trim() ?? "",
      schemaFingerprint: mainSchemaFingerprint(worksheet),
      ...scenarioStatus,
      turns: parsedTurns.turns.map((turn) => ({
        sheetName: worksheet.name,
        rowNumber,
        turnNumber: turn.turnNumber,
        userInput: turn.userInput,
      })),
    });
  }

  return scenarios;
}

function emptySummary(): PgnSheetSummary {
  return {
    scenarios: 0,
    runnableTurns: 0,
    missingUserInput: 0,
    multiTurnScenarios: 0,
    completedScenarios: 0,
  };
}

export function isScenarioComplete(
  workbook: ExcelJS.Workbook,
  scenario: PgnTestScenario,
): boolean {
  const worksheet = workbook.getWorksheet(scenario.sheetName);
  if (!worksheet || scenario.turns.length === 0) {
    return false;
  }
  if (scenario.sheetKind === "negative") {
    return Boolean(
      cellText(fieldCell(worksheet, scenario.sourceRowNumber, "botResponse")).trim(),
    );
  }
  return scenario.turns.every((turn) =>
    Boolean(cellText(fieldCell(worksheet, turn.rowNumber, "botResponse")).trim()),
  );
}

export function isScenarioPartiallyComplete(
  workbook: ExcelJS.Workbook,
  scenario: PgnTestScenario,
): boolean {
  if (scenario.sheetKind === "negative" || scenario.turns.length < 2) {
    return false;
  }
  const worksheet = workbook.getWorksheet(scenario.sheetName);
  if (!worksheet) {
    return false;
  }
  const completedTurns = scenario.turns.filter((turn) =>
    Boolean(cellText(fieldCell(worksheet, turn.rowNumber, "botResponse")).trim()),
  ).length;
  return completedTurns > 0 && completedTurns < scenario.turns.length;
}

export function parsePgnWorkbook(workbook: ExcelJS.Workbook): ParsedPgnWorkbook {
  const issues: PgnValidationIssue[] = [];
  const testIdRows = new Map<
    string,
    { sheetName: string; rowNumber: number }
  >();
  const scenarios: PgnTestScenario[] = [];
  const kbSheet = workbook.getWorksheet(KB_SHEET_NAME);
  const negativeSheet = workbook.getWorksheet(NEGATIVE_SHEET_NAME);
  const schemas = WORKBOOK_SCHEMAS.flatMap((definition) => {
    const worksheet = workbook.getWorksheet(definition.sheetName);
    return worksheet ? [resolveWorksheetSchema(worksheet, definition)] : [];
  });
  for (const mapping of schemas) {
    for (const issue of mapping.issues) issues.push({ code: "INVALID_HEADER", severity: issue.severity, sheetName: mapping.sheetName, rowNumber: mapping.headerRow, message: issue.message });
  }

  if (!kbSheet) {
    issues.push({
      code: "MISSING_SHEET",
      severity: "ERROR",
      sheetName: KB_SHEET_NAME,
      message: `Worksheet "${KB_SHEET_NAME}" was not found.`,
    });
  } else if (resolveWorksheetSchema(kbSheet, KB_SCHEMA).valid) {
    scenarios.push(...parseKnowledgeBaseSheet(kbSheet, issues, testIdRows));
  }
  if (!negativeSheet) {
    issues.push({
      code: "MISSING_SHEET",
      severity: "ERROR",
      sheetName: NEGATIVE_SHEET_NAME,
      message: `Worksheet "${NEGATIVE_SHEET_NAME}" was not found.`,
    });
  } else if (resolveWorksheetSchema(negativeSheet, NEGATIVE_SCHEMA).valid) {
    scenarios.push(...parseNegativeSheet(negativeSheet, issues, testIdRows));
  }

  const summaries: Record<PgnSheetKind, PgnSheetSummary> = {
    kb: emptySummary(),
    negative: emptySummary(),
  };
  for (const kind of ["kb", "negative"] as const) {
    const kindScenarios = scenarios.filter((scenario) => scenario.sheetKind === kind);
    summaries[kind] = {
      scenarios: kindScenarios.length,
      runnableTurns: kindScenarios.reduce(
        (count, scenario) => count + scenario.turns.length,
        0,
      ),
      missingUserInput: issues.filter(
        (issue) =>
          issue.code === "MISSING_USER_INPUT" &&
          issue.sheetName ===
            (kind === "kb" ? KB_SHEET_NAME : NEGATIVE_SHEET_NAME),
      ).length,
      multiTurnScenarios: kindScenarios.filter(
        (scenario) => scenario.turns.length > 1,
      ).length,
      completedScenarios: kindScenarios.filter((scenario) =>
        isScenarioComplete(workbook, scenario),
      ).length,
    };
  }

  return {
    schemas,
    scenarios,
    issues,
    summaries,
    duplicateTestCaseIds: issues.filter(
      (issue) => issue.code === "DUPLICATE_TEST_ID",
    ).length,
    invalidTurnRows: issues.filter(
      (issue) =>
        issue.code === "INVALID_TURN" ||
        issue.code === "INVALID_TURN_ROW",
    ).length,
  };
}

export async function loadPgnWorkbook(
  filePath: string,
): Promise<PgnWorkbookDocument> {
  await access(filePath).catch(() => {
    throw new Error(`PGN source workbook was not found at ${filePath}`);
  });
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);
  await attachWorkbookMappings(workbook, filePath);
  return { workbook, parsed: parsePgnWorkbook(workbook) };
}
