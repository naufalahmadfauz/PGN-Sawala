import type { Workbook } from "exceljs";
import { readExecutionTransport, readSessionMode, type RunExecutionContext } from "../session-mode";
import { RUN_CONFIGURATION_SCHEMA, ensureOwnedWorksheet, fieldCell, optionalFieldCell } from "./workbook-schema";

export interface RunConfiguration extends RunExecutionContext {
  runId: string;
  sessionResetAttempts: number;
  restartedFromRunId?: string;
}

export function getRunConfiguration(workbook: Workbook, runId: string): RunConfiguration | undefined {
  const sheet = workbook.getWorksheet(RUN_CONFIGURATION_SCHEMA.sheetName);
  if (!sheet) return undefined;
  for (let row = 2; row <= sheet.rowCount; row += 1) {
    if (fieldCell(sheet, row, "runId").text !== runId) continue;
    const sessionResetAttempts = Number(fieldCell(sheet, row, "sessionResetAttempts").value);
    if (!Number.isInteger(sessionResetAttempts) || sessionResetAttempts < 0) throw new Error("Run Configuration has an invalid session reset count");
    return {
      runId,
      sessionMode: readSessionMode(fieldCell(sheet, row, "sessionMode").text),
      transport: readExecutionTransport(fieldCell(sheet, row, "transport").text),
      sessionResetAttempts,
      restartedFromRunId: optionalFieldCell(sheet, row, "restartedFromRunId")?.text || undefined,
    };
  }
  return undefined;
}

export function upsertRunConfiguration(workbook: Workbook, configuration: RunConfiguration): void {
  readSessionMode(configuration.sessionMode);
  readExecutionTransport(configuration.transport);
  if (!Number.isInteger(configuration.sessionResetAttempts) || configuration.sessionResetAttempts < 0) {
    throw new Error("Run Configuration has an invalid session reset count");
  }
  const previous = getRunConfiguration(workbook, configuration.runId);
  if (previous && (previous.sessionMode !== configuration.sessionMode || previous.transport !== configuration.transport)) {
    throw new Error("Run Configuration session mode and transport are immutable");
  }
  const sheet = ensureOwnedWorksheet(workbook, RUN_CONFIGURATION_SCHEMA);
  let row = sheet.rowCount + 1;
  for (let candidate = 2; candidate <= sheet.rowCount; candidate += 1) {
    if (fieldCell(sheet, candidate, "runId").text === configuration.runId) { row = candidate; break; }
  }
  fieldCell(sheet, row, "runId").value = configuration.runId;
  fieldCell(sheet, row, "transport").value = configuration.transport;
  fieldCell(sheet, row, "sessionMode").value = configuration.sessionMode;
  fieldCell(sheet, row, "sessionResetAttempts").value = configuration.sessionResetAttempts;
  const parent = optionalFieldCell(sheet, row, "restartedFromRunId");
  if (parent) parent.value = configuration.restartedFromRunId ?? "";
}

export function runExecutionContext(workbook: Workbook, runId: string): RunExecutionContext {
  const stored = getRunConfiguration(workbook, runId);
  return { transport: stored?.transport ?? "whatsapp", sessionMode: stored?.sessionMode ?? "isolated" };
}
