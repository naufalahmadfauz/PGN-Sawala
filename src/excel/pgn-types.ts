import type ExcelJS from "exceljs";

export const KB_SHEET_NAME = "Test Case Knowledge Base";
export const NEGATIVE_SHEET_NAME = "Negative Case";
export const TRANSCRIPT_SHEET_NAME = "Execution Transcript";

export type EvidenceStatus =
  | "EVIDENCE_PENDING"
  | "EVIDENCE_SYNCED"
  | "EVIDENCE_ALREADY_SYNCED"
  | "EVIDENCE_LOCAL_ONLY"
  | "EVIDENCE_CAPTURE_ERROR"
  | "EVIDENCE_UPLOAD_ERROR"
  | "EVIDENCE_MISSING"
  | "EVIDENCE_REQUIRES_RERUN";

export type PgnSheetKind = "kb" | "negative";
export type TechnicalStatus =
  | "CAPTURED"
  | "TIMEOUT"
  | "SEND_ERROR"
  | "CHAT_ERROR";

export interface PgnTestTurn {
  sheetName: string;
  rowNumber: number;
  turnNumber: number;
  userInput: string;
}

export interface PgnTestScenario {
  testCaseId: string;
  sheetKind: PgnSheetKind;
  sheetName: string;
  sourceRowNumber: number;
  category: string;
  turns: PgnTestTurn[];
}

export type ValidationIssueCode =
  | "DUPLICATE_TEST_ID"
  | "INVALID_HEADER"
  | "INVALID_TURN"
  | "INVALID_TURN_ROW"
  | "MISSING_SHEET"
  | "MISSING_TEST_CASE_ID"
  | "MISSING_USER_INPUT";

export interface PgnValidationIssue {
  code: ValidationIssueCode;
  severity: "ERROR" | "WARNING";
  sheetName: string;
  rowNumber?: number;
  message: string;
}

export interface PgnSheetSummary {
  scenarios: number;
  runnableTurns: number;
  missingUserInput: number;
  multiTurnScenarios: number;
  completedScenarios: number;
}

export interface ParsedPgnWorkbook {
  scenarios: PgnTestScenario[];
  issues: PgnValidationIssue[];
  summaries: Record<PgnSheetKind, PgnSheetSummary>;
  duplicateTestCaseIds: number;
  invalidTurnRows: number;
}

export interface PgnWorkbookDocument {
  workbook: ExcelJS.Workbook;
  parsed: ParsedPgnWorkbook;
}

export interface ExecutedBotMessage {
  sequence: number;
  message: string;
  timestamp: Date;
}

export interface ExecutedTurn {
  turn: PgnTestTurn;
  technicalStatus: TechnicalStatus;
  sentAt?: Date;
  completedAt: Date;
  botMessages: ExecutedBotMessage[];
  combinedResponse: string;
  firstResponseMs?: number;
  totalResponseMs?: number;
  error?: string;
  evidencePath?: string;
  evidenceUrl?: string;
  evidenceStatus?: EvidenceStatus;
  evidenceDriveFileId?: string;
  evidenceDriveFileName?: string;
}
