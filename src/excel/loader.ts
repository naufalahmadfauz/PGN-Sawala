import { access } from "node:fs/promises";
import ExcelJS, { type Cell, type Worksheet } from "exceljs";
import type { TestCase } from "../types";

const requiredColumns = {
  testId: ["testid", "id"],
  category: ["category"],
  userInput: ["userinput", "input", "message"],
  expectedBehaviour: [
    "expectedbehaviour",
    "expectedbehavior",
    "expected",
  ],
} as const;

function normalizeHeader(value: string): string {
  return value.toLowerCase().replace(/[^a-z0-9]/g, "");
}

function cellText(cell: Cell, trim = true): string {
  const value = cell.text.replace(/\r\n/g, "\n");
  return trim ? value.trim() : value;
}

function findColumn(
  headers: Map<string, number>,
  aliases: readonly string[],
): number | undefined {
  return aliases.map((alias) => headers.get(alias)).find(Boolean);
}

function readHeaders(worksheet: Worksheet): Map<string, number> {
  const headers = new Map<string, number>();
  worksheet.getRow(1).eachCell({ includeEmpty: false }, (cell, column) => {
    const header = normalizeHeader(cellText(cell));
    if (header) {
      headers.set(header, column);
    }
  });
  return headers;
}

export async function loadTestCases(filePath: string): Promise<TestCase[]> {
  await access(filePath).catch(() => {
    throw new Error(
      `Test data file was not found at ${filePath}. Run npm run data:template to create a starter workbook.`,
    );
  });

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);
  const worksheet = workbook.getWorksheet("Test Cases");
  if (!worksheet) {
    throw new Error(
      `Worksheet "Test Cases" was not found in ${filePath}. Refusing to treat a report workbook as test input.`,
    );
  }

  const headers = readHeaders(worksheet);
  const columns = {
    testId: findColumn(headers, requiredColumns.testId),
    category: findColumn(headers, requiredColumns.category),
    userInput: findColumn(headers, requiredColumns.userInput),
    expectedBehaviour: findColumn(
      headers,
      requiredColumns.expectedBehaviour,
    ),
    scenarioId: findColumn(headers, ["scenarioid", "scenario", "group"]),
  };
  const missingColumns = Object.entries(columns)
    .filter(([name, column]) => name !== "scenarioId" && column === undefined)
    .map(([name]) => name);
  if (missingColumns.length > 0) {
    throw new Error(
      `Missing required columns in ${filePath}: ${missingColumns.join(", ")}`,
    );
  }

  const testCases: TestCase[] = [];
  const validationErrors: string[] = [];
  const testIdRows = new Map<string, number>();
  for (let rowNumber = 2; rowNumber <= worksheet.rowCount; rowNumber += 1) {
    const row = worksheet.getRow(rowNumber);
    const testId = cellText(row.getCell(columns.testId!));
    const userInput = cellText(row.getCell(columns.userInput!), false);
    const category = cellText(row.getCell(columns.category!));
    const expectedBehaviour = cellText(
      row.getCell(columns.expectedBehaviour!),
    );
    const scenarioId = columns.scenarioId
      ? cellText(row.getCell(columns.scenarioId))
      : undefined;

    if (!testId && !userInput.trim() && !category && !expectedBehaviour) {
      continue;
    }
    if (!testId) {
      validationErrors.push(`Row ${rowNumber}: Test ID is required`);
    }
    if (!userInput.trim()) {
      validationErrors.push(`Row ${rowNumber}: User Input is required`);
    }
    const previousRow = testIdRows.get(testId);
    if (testId && previousRow !== undefined) {
      validationErrors.push(
        `Row ${rowNumber}: Test ID "${testId}" duplicates row ${previousRow}`,
      );
    } else if (testId) {
      testIdRows.set(testId, rowNumber);
    }
    if (testId && userInput.trim()) {
      testCases.push({
        testId,
        category,
        userInput,
        expectedBehaviour,
        scenarioId: scenarioId || undefined,
        sourceRow: rowNumber,
      });
    }
  }

  if (validationErrors.length > 0) {
    throw new Error(validationErrors.join("\n"));
  }
  if (testCases.length === 0) {
    throw new Error(`No test cases were found in ${filePath}`);
  }

  return testCases;
}
