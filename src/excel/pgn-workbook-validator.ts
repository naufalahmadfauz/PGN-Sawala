import {
  KB_SHEET_NAME,
  NEGATIVE_SHEET_NAME,
  type ParsedPgnWorkbook,
  type PgnValidationIssue,
} from "./pgn-types";

export interface PgnSessionIsolationConfig {
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
): string[] {
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
  const lines = [
    "PGN workbook validation",
    "",
    ...formatSheetSummary(KB_SHEET_NAME, parsed.summaries.kb),
    "",
    ...formatSheetSummary(NEGATIVE_SHEET_NAME, parsed.summaries.negative),
    "",
    `Duplicate Test Case IDs: ${parsed.duplicateTestCaseIds}`,
    `Invalid turn rows: ${parsed.invalidTurnRows}`,
    "",
    "Session isolation",
    "-----------------",
    `Strategy: WhatsApp Conversation Builder debug command "${isolation.command}"`,
    `Expected confirmation: "${isolation.confirmation}"`,
    `Reset timeout: ${isolation.timeoutMs} ms`,
    `Post-reset quiet window: ${isolation.postResetQuietMs} ms`,
    "Status: ENABLED",
    "",
    "Response completion",
    "-------------------",
    `Idle window: ${isolation.responseIdleMs} ms`,
    `Hard timeout: ${isolation.responseTimeoutMs} ms`,
    "Multiple bot bubbles: ENABLED",
  ];

  if (parsed.issues.length > 0) {
    lines.push("", ...parsed.issues.map(formatIssue), "", "NOT READY");
  } else {
    lines.push("", "READY TO EXECUTE");
  }
  return lines.join("\n");
}

export function assertPgnWorkbookValid(parsed: ParsedPgnWorkbook): void {
  const errors = parsed.issues.filter((issue) => issue.severity === "ERROR");
  if (errors.length > 0) {
    throw new Error(
      `PGN workbook validation failed with ${errors.length} error(s). Run npm run test:pgn:validate.`,
    );
  }
}
