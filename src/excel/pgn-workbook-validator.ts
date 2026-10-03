import type { Workbook } from "exceljs";
import {
  KB_SHEET_NAME,
  NEGATIVE_SHEET_NAME,
  type ParsedPgnWorkbook,
  type PgnSheetKind,
  type PgnValidationIssue,
} from "./pgn-types";
import { formatWorksheetSchema } from "./workbook-schema";
import {
  CONTINUOUS_SESSION_WARNING,
  readSessionMode,
  sessionModeLabel,
  type SessionMode,
} from "../session-mode";

export interface PgnSessionIsolationConfig {
  sessionMode?: SessionMode;
  command: string;
  confirmation: string;
  timeoutMs: number;
  responseIdleMs: number;
  responseTimeoutMs: number;
  postResetQuietMs: number;
}

function formatSheetSummary(
  name: string,
  summary: ParsedPgnWorkbook["summaries"]["kb"],
  present: boolean,
): string[] {
  if (!present) return [name, "-".repeat(name.length), "Not present (optional)"];
  return [
    name,
    "-".repeat(name.length),
    `Scenarios: ${summary.scenarios}`,
    `Runnable turns: ${summary.runnableTurns}`,
    `Missing User Input: ${summary.missingUserInput}`,
    `Multi-turn scenarios: ${summary.multiTurnScenarios}`,
    `Already completed scenarios: ${summary.completedScenarios}`,
  ];
}

function formatIssue(issue: PgnValidationIssue): string {
  const location = issue.rowNumber
    ? `${issue.sheetName} row ${issue.rowNumber}`
    : issue.sheetName;
  return `${issue.severity}:\n${location}:\n${issue.message}`;
}

export function formatPgnValidation(
  parsed: ParsedPgnWorkbook,
  isolation: PgnSessionIsolationConfig,
): string {
  const sessionMode = readSessionMode(isolation.sessionMode);
  const isolated = sessionMode === "isolated";
  const lines = [
    "PGN workbook validation",
    "Workbook schema",
    ...(parsed.schemas ?? []).map(formatWorksheetSchema),
    "",
    ...formatSheetSummary(KB_SHEET_NAME, parsed.summaries.kb, parsed.availableSheets.includes("kb")),
    "",
    ...formatSheetSummary(NEGATIVE_SHEET_NAME, parsed.summaries.negative, parsed.availableSheets.includes("negative")),
    "",
    `Duplicate Test Case IDs: ${parsed.duplicateTestCaseIds}`,
    `Invalid turn rows: ${parsed.invalidTurnRows}`,
    "",
    "Session isolation",
    "-----------------",
    "Transport: WhatsApp",
    `Session Mode: ${sessionModeLabel(sessionMode)}`,
    "Initial reset: Required",
    `Between-scenario reset: ${isolated ? "Enabled" : "Disabled"}`,
    `Final cleanup reset: ${isolated ? "Enabled" : "Disabled"}`,
    `Strategy: WhatsApp Conversation Builder debug command "${isolation.command}"`,
    `Expected confirmation: "${isolation.confirmation}"`,
    `Reset timeout: ${isolation.timeoutMs} ms`,
    `Post-reset quiet window: ${isolation.postResetQuietMs} ms`,
    `Status: ${isolated ? "ENABLED" : "DISABLED"}`,
    ...(!isolated ? ["", CONTINUOUS_SESSION_WARNING] : []),
    "",
    "Response completion",
    "-------------------",
    `Idle window: ${isolation.responseIdleMs} ms`,
    `Hard timeout: ${isolation.responseTimeoutMs} ms`,
    "Multiple bot bubbles: ENABLED",
  ];

  const errors = parsed.issues.filter((issue) => issue.severity === "ERROR");
  if (errors.length > 0) {
    lines.push("", ...parsed.issues.map(formatIssue), "", "NOT READY");
  } else if (parsed.issues.length > 0) {
    lines.push(
      "",
      ...parsed.issues.map(formatIssue),
      "",
      "READY WITH WARNINGS",
    );
  } else {
    lines.push("", "READY TO EXECUTE");
  }
  return lines.join("\n");
}

export function assertMatchingTestCaseSheets(source: Workbook, executed: Workbook): void {
  const sheets = [KB_SHEET_NAME, NEGATIVE_SHEET_NAME];
  const onlyInSource = sheets.filter((name) => source.getWorksheet(name) && !executed.getWorksheet(name));
  const onlyInExecuted = sheets.filter((name) => executed.getWorksheet(name) && !source.getWorksheet(name));
  if (onlyInSource.length || onlyInExecuted.length) {
    const describe = (names: string[]) => names.map((name) => `"${name}"`).join(", ") || "(none)";
    throw new Error(
      `Source and executed workbook test-case sheet sets differ. Only in source: ${describe(onlyInSource)}. Only in executed workbook: ${describe(onlyInExecuted)}. Use a new executed-workbook path (PGN_EXECUTED_WORKBOOK) to retain the existing results.`,
    );
  }
}

export function assertPgnWorkbookValid(parsed: ParsedPgnWorkbook, requestedSheet?: PgnSheetKind): void {
  const errors = parsed.issues.filter((issue) => issue.severity === "ERROR");
  if (errors.length > 0) {
    throw new Error(
      `PGN workbook validation failed with ${errors.length} error(s). ${errors.map((issue) => `${issue.sheetName}: ${issue.message}`).join(" ")} Run npm run test:pgn:validate or review column mapping in npm run pgn.`,
    );
  }
  if (requestedSheet && !parsed.availableSheets.includes(requestedSheet)) {
    const sheetName = requestedSheet === "kb" ? KB_SHEET_NAME : NEGATIVE_SHEET_NAME;
    const available = parsed.availableSheets.map((kind) => `"${kind === "kb" ? KB_SHEET_NAME : NEGATIVE_SHEET_NAME}"`).join(", ");
    throw new Error(
      `Requested ${requestedSheet === "kb" ? "positive" : "negative"} category (--sheet ${requestedSheet}), worksheet "${sheetName}", is absent. Available test-case sheets: ${available}.`,
    );
  }
}
