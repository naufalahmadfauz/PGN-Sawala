import { createHash } from "node:crypto";
import type { Cell, CellValue, Row, Workbook, Worksheet } from "exceljs";

export interface WorkbookFieldDefinition {
  field: WorkbookField;
  header: string;
  aliases?: readonly string[];
  required: boolean;
  writable: boolean;
  type?: "text" | "number" | "date" | "url";
}

export type WorkbookField =
  | "number" | "knowledgeBaseArticle" | "category" | "testCaseId" | "role"
  | "scenario" | "expectedResponse" | "turn" | "userInput" | "botResponse"
  | "negativeCondition" | "reference" | "responseTime" | "testDate" | "status"
  | "notes" | "evidence" | "runId" | "sheet" | "excelRow" | "message"
  | "timestamp" | "firstResponseMs" | "totalResponseMs" | "error"
  | "evidencePath" | "evidenceUrl" | "evidenceStatus" | "folderId" | "folderUrl"
  | "migrationVersion" | "mode" | "evidenceKey" | "driveFileId" | "driveFileName"
  | "localCleanPath" | "retestRunId" | "previousStatus" | "previousBotResponse"
  | "previousResponseTime" | "previousTestDate" | "previousEvidenceUrl"
  | "retestedAt" | "newTechnicalStatus" | "newBotResponse" | "newResponseTime"
  | "newTestDate" | "newEvidenceUrl" | "historyKey" | "startedAt" | "state"
  | "selectedIds" | "finishedIds" | "updatedAt"
  | "transport" | "sessionMode" | "sessionResetAttempts" | "restartedFromRunId"
  | "conversationId" | "dialogId" | "restResponseIdleMs" | "restResponseTimeoutMs" | "restPollIntervalMs";

export interface WorksheetSchemaDefinition {
  id: string;
  sheetName: string;
  fields: readonly WorkbookFieldDefinition[];
  interactive?: boolean;
  metadataGroup?: "run" | "file";
}

const field = (
  name: WorkbookField, header: string, required = true, writable = false,
  aliases?: readonly string[],
): WorkbookFieldDefinition => ({ field: name, header, required, writable, aliases });

export const KB_SCHEMA: WorksheetSchemaDefinition = {
  id: "kb", sheetName: "Test Case Knowledge Base", interactive: true,
  fields: [
    field("number", "No.", false), field("knowledgeBaseArticle", "Knowledge Base Article", false),
    field("testCaseId", "Test Case ID"), field("role", "Role Pengujian", false),
    field("scenario", "Scenario / Test Objective", false), field("expectedResponse", "Expected Bot Response", false),
    field("turn", "Turn"),
    field("userInput", "User Input", true, false, ["User Question", "Test Input", "User Message"]),
    field("botResponse", "Bot Response", true, true, ["Actual Bot Response", "Actual Response", "Bot Answer"]),
    field("responseTime", "Response Time", false, true), field("testDate", "Test Date", false, true),
    field("status", "Status", true, true), field("notes", "Notes", false, true),
    field("evidence", "Evidence", false, true, ["Evidence URL"]),
  ],
};
export const NEGATIVE_SCHEMA: WorksheetSchemaDefinition = {
  id: "negative", sheetName: "Negative Case", interactive: true,
  fields: [
    field("number", "No.", false), field("category", "Category", false), field("testCaseId", "Test Case ID"),
    field("scenario", "Scenario / Test Objective", false),
    field("userInput", "User Input / Test Steps", true, false, ["User Input", "Test Steps", "User Question", "Test Input", "User Message"]),
    field("negativeCondition", "Negative Condition", false), field("expectedResponse", "Expected Handling", false),
    field("botResponse", "Bot Response", true, true, ["Actual Bot Response", "Actual Response", "Bot Answer"]),
    field("responseTime", "Response Time", false, true), field("testDate", "Test Date", false, true),
    field("status", "Status", true, true), field("notes", "Notes", false, true), field("reference", "Reference", false),
    field("evidence", "Evidence", false, true, ["Evidence URL"]),
  ],
};
export const TRANSCRIPT_SCHEMA: WorksheetSchemaDefinition = {
  id: "transcript", sheetName: "Execution Transcript",
  fields: [
    field("runId", "Run ID"), field("testCaseId", "Test Case ID"), field("sheet", "Sheet"),
    field("excelRow", "Excel Row"), field("turn", "Turn"), field("role", "Role"),
    field("message", "Message"), field("timestamp", "Timestamp"),
    field("firstResponseMs", "First Response (ms)"), field("totalResponseMs", "Total Response (ms)"),
    field("status", "Status"), field("error", "Error"), field("evidencePath", "Evidence Path"),
    field("evidenceUrl", "Evidence URL", false, true), field("evidenceStatus", "Evidence Status", false, true),
    field("transport", "Transport", false, true), field("sessionMode", "Session Mode", false, true),
    field("conversationId", "Conversation ID", false, true), field("dialogId", "Dialog ID", false, true),
  ],
};
export const EVIDENCE_RUN_SCHEMA: WorksheetSchemaDefinition = {
  id: "evidenceRun", sheetName: "Execution Metadata", metadataGroup: "run",
  fields: [
    field("runId", "Run ID"), field("folderId", "Evidence Drive Folder ID"),
    field("folderUrl", "Evidence Drive Folder URL"), field("migrationVersion", "Evidence Migration Version"),
    field("timestamp", "Migration Timestamp"), field("mode", "Mode"),
  ],
};
export const EVIDENCE_FILE_SCHEMA: WorksheetSchemaDefinition = {
  id: "evidenceFile", sheetName: "Execution Metadata", metadataGroup: "file",
  fields: [
    field("evidenceKey", "Evidence Key"), field("runId", "Run ID"), field("testCaseId", "Test Case ID"),
    field("turn", "Turn"), field("driveFileId", "Drive File ID"), field("driveFileName", "Drive File Name"),
    field("evidenceUrl", "Evidence URL"), field("localCleanPath", "Local Clean Path"), field("evidenceStatus", "Evidence Status"),
  ],
};
export const RETEST_HISTORY_SCHEMA: WorksheetSchemaDefinition = {
  id: "retestHistory", sheetName: "Retest History",
  fields: [
    field("runId", "Run ID"), field("retestRunId", "Retest Run ID"), field("testCaseId", "Test Case ID"),
    field("sheet", "Sheet"), field("excelRow", "Excel Row"), field("turn", "Turn"),
    field("previousStatus", "Previous Status"), field("previousBotResponse", "Previous Bot Response"),
    field("previousResponseTime", "Previous Response Time"), field("previousTestDate", "Previous Test Date"),
    field("previousEvidenceUrl", "Previous Evidence URL"), field("retestedAt", "Retested At"),
    field("newTechnicalStatus", "New Technical Status"), field("newBotResponse", "New Bot Response"),
    field("newResponseTime", "New Response Time"), field("newTestDate", "New Test Date"),
    field("newEvidenceUrl", "New Evidence URL"), field("historyKey", "History Key"),
  ],
};
export const RETEST_METADATA_SCHEMA: WorksheetSchemaDefinition = {
  id: "retestMetadata", sheetName: "Retest Metadata",
  fields: [
    field("runId", "Retest Run ID"), field("startedAt", "Started At"), field("state", "State"),
    field("selectedIds", "Selected Test IDs"), field("finishedIds", "Successfully Completed Test IDs"),
    field("updatedAt", "Last Updated"), field("folderId", "Evidence Drive Folder ID"), field("folderUrl", "Evidence Drive Folder URL"),
  ],
};
export const RUN_CONFIGURATION_SCHEMA: WorksheetSchemaDefinition = {
  id: "runConfiguration", sheetName: "Run Configuration",
  fields: [
    field("runId", "Run ID"), field("transport", "Transport"), field("sessionMode", "Session Mode"),
    field("sessionResetAttempts", "Session Reset Attempts"), field("restartedFromRunId", "Restarted From Run ID", false),
    field("restResponseIdleMs", "REST Response Idle (ms)", false),
    field("restResponseTimeoutMs", "REST Response Timeout (ms)", false),
    field("restPollIntervalMs", "REST Poll Interval (ms)", false),
  ],
};
export const WORKBOOK_SCHEMAS = [KB_SCHEMA, NEGATIVE_SCHEMA, TRANSCRIPT_SCHEMA, EVIDENCE_RUN_SCHEMA, EVIDENCE_FILE_SCHEMA, RETEST_HISTORY_SCHEMA, RETEST_METADATA_SCHEMA, RUN_CONFIGURATION_SCHEMA];
for (const schema of WORKBOOK_SCHEMAS.filter((schema) => !schema.interactive)) {
  schema.fields = schema.fields.map((field) => ({ ...field, writable: true }));
}
export const KB_HEADERS = KB_SCHEMA.fields.filter((item) => item.field !== "evidence").map((item) => item.header);
export const NEGATIVE_HEADERS = NEGATIVE_SCHEMA.fields.filter((item) => item.field !== "evidence").map((item) => item.header);

export interface WorkbookMappingOverride { sheet: string; field: WorkbookField; header: string }
export interface ResolvedWorkbookField {
  header: string; columnIndex: number; columnLetter: string;
  match: "canonical" | "normalized" | "alias" | "override";
}
export interface WorksheetSchemaIssue {
  field: WorkbookField; severity: "ERROR" | "WARNING"; message: string;
  kind: "missing" | "ambiguous" | "collision";
}
export interface ResolvedWorksheetSchema {
  sheetName: string; headerRow: number; definition: WorksheetSchemaDefinition;
  fields: Partial<Record<WorkbookField, ResolvedWorkbookField>>;
  candidates: Partial<Record<WorkbookField, ResolvedWorkbookField[]>>;
  headers: Array<{ header: string; columnIndex: number; columnLetter: string }>;
  issues: WorksheetSchemaIssue[]; valid: boolean; fingerprint: string;
}
const overridesByWorkbook = new WeakMap<Workbook, readonly WorkbookMappingOverride[]>();
const resolvedCache = new WeakMap<Worksheet, Map<string, { signature: string; mapping: ResolvedWorksheetSchema }>>();
export function setWorkbookSchemaOverrides(workbook: Workbook, overrides: readonly WorkbookMappingOverride[]): void {
  overridesByWorkbook.set(workbook, overrides);
}
export function normalizeWorkbookHeader(value: string): string {
  return value.replace(/\s+/g, " ").trim().toLowerCase();
}
export function schemaForWorksheet(worksheet: Worksheet): WorksheetSchemaDefinition {
  const schema = WORKBOOK_SCHEMAS.find((item) => item.sheetName === worksheet.name);
  if (!schema || schema.metadataGroup) throw new Error(`An explicit schema is required for worksheet "${worksheet.name}"`);
  return schema;
}

export function resolveWorksheetSchema(
  worksheet: Worksheet,
  definition = schemaForWorksheet(worksheet),
  overrides = overridesByWorkbook.get(worksheet.workbook) ?? [],
): ResolvedWorksheetSchema {
  const headers: ResolvedWorksheetSchema["headers"] = [];
  worksheet.getRow(1).eachCell((cell, columnIndex) => {
    if (cell.text.trim()) headers.push({ header: cell.text, columnIndex, columnLetter: worksheet.getColumn(columnIndex).letter });
  });
  const signature = JSON.stringify([headers, overrides, definition]);
  const cache = resolvedCache.get(worksheet) ?? new Map<string, { signature: string; mapping: ResolvedWorksheetSchema }>();
  const previous = cache.get(definition.id);
  if (previous?.signature === signature) return previous.mapping;
  let scopedHeaders = headers;
  if (definition.metadataGroup) {
    // Shipped metadata has two independent Run ID fields, separated by Evidence Key.
    const keys = headers.filter((item) => normalizeWorkbookHeader(item.header) === "evidence key");
    scopedHeaders = keys.length === 1
      ? headers.filter((item) => definition.metadataGroup === "run" ? item.columnIndex < keys[0].columnIndex : item.columnIndex >= keys[0].columnIndex)
      : [];
  }
  const fields: ResolvedWorksheetSchema["fields"] = {};
  const candidates: ResolvedWorksheetSchema["candidates"] = {};
  const issues: WorksheetSchemaIssue[] = [];
  for (const item of definition.fields) {
    const override = definition.interactive
      ? overrides.find((value) => value.sheet === worksheet.name && value.field === item.field)
      : undefined;
    const recognized = scopedHeaders.flatMap((header): ResolvedWorkbookField[] => {
      const normalized = normalizeWorkbookHeader(header.header);
      const match = header.header === item.header ? "canonical"
        : normalized === normalizeWorkbookHeader(item.header) ? "normalized"
        : item.aliases?.some((alias) => normalizeWorkbookHeader(alias) === normalized) ? "alias"
        : override && normalizeWorkbookHeader(override.header) === normalized ? "override" : undefined;
      return match ? [{ ...header, match }] : [];
    });
    candidates[item.field] = recognized;
    const selected = override ? recognized.filter((header) => normalizeWorkbookHeader(header.header) === normalizeWorkbookHeader(override.header)) : [];
    if (override && selected.length === 0) {
      issues.push({ field: item.field, kind: "missing", severity: "ERROR", message: `Saved header override "${override.header}" for ${item.header} is missing; review mapping rather than switching columns silently.` });
      continue;
    }
    const matches = selected.length ? selected : recognized;
    if (matches.length === 1) {
      fields[item.field] = { ...matches[0], match: selected.length ? "override" : matches[0].match };
    } else if (matches.length > 1) {
      issues.push({ field: item.field, kind: "ambiguous", severity: "ERROR", message: `Ambiguous mapping for "${item.header}": ${matches.map((match) => `${match.columnLetter} - ${match.header}`).join(", ")}. Review column mapping.` });
    } else {
      issues.push({ field: item.field, kind: "missing", severity: item.required ? "ERROR" : "WARNING", message: `${item.required ? "Required" : "Optional"} column "${item.header}" not found in header row 1.${item.required ? " Review column mapping before execution." : item.field === "evidence" ? " Evidence is appended outside tables only when preparing an executed workbook." : " This field will not be read or written."}` });
    }
  }
  const used = new Map<number, WorkbookField>();
  for (const [name, value] of Object.entries(fields)) {
    const other = used.get(value.columnIndex);
    if (other) issues.push({ field: name as WorkbookField, kind: "collision", severity: "ERROR", message: `Column ${value.columnLetter} is mapped to both ${other} and ${name}; review column mapping.` });
    used.set(value.columnIndex, name as WorkbookField);
  }
  const fingerprint = createHash("sha256").update(JSON.stringify({ sheet: worksheet.name, headerRow: 1, headers: headers.map((item) => [item.columnIndex, normalizeWorkbookHeader(item.header)]) })).digest("hex");
  const mapping = { sheetName: worksheet.name, headerRow: 1, definition, fields, candidates, headers, issues, fingerprint, valid: !issues.some((issue) => issue.severity === "ERROR") };
  cache.set(definition.id, { signature, mapping });
  resolvedCache.set(worksheet, cache);
  return mapping;
}

export function getWorksheetSchema(worksheet: Worksheet, definition?: WorksheetSchemaDefinition): ResolvedWorksheetSchema {
  const mapping = resolveWorksheetSchema(worksheet, definition);
  if (!mapping.valid) throw new Error(`Workbook schema validation failed: ${worksheet.name}. ${mapping.issues.filter((issue) => issue.severity === "ERROR").map((issue) => issue.message).join(" ")}`);
  return mapping;
}
export function fieldColumn(worksheet: Worksheet, field: WorkbookField, definition?: WorksheetSchemaDefinition): number {
  const mapping = getWorksheetSchema(worksheet, definition);
  const column = mapping.fields[field]?.columnIndex;
  if (column === undefined) throw new Error(`Column for ${field} is unavailable in ${worksheet.name}; refusing positional fallback`);
  return column;
}
export function fieldCell(worksheet: Worksheet, row: number, field: WorkbookField, definition?: WorksheetSchemaDefinition): Cell {
  return worksheet.getCell(row, fieldColumn(worksheet, field, definition));
}
export function optionalFieldCell(worksheet: Worksheet, row: number, field: WorkbookField, definition?: WorksheetSchemaDefinition): Cell | undefined {
  const column = getWorksheetSchema(worksheet, definition).fields[field]?.columnIndex;
  return column === undefined ? undefined : worksheet.getCell(row, column);
}
export function appendSchemaRow(worksheet: Worksheet, values: Partial<Record<WorkbookField, CellValue>>, definition?: WorksheetSchemaDefinition): Row {
  const mapping = getWorksheetSchema(worksheet, definition);
  const row = worksheet.addRow([]);
  for (const [name, value] of Object.entries(values)) {
    const column = mapping.fields[name as WorkbookField]?.columnIndex;
    if (column !== undefined) row.getCell(column).value = value ?? null;
  }
  return row;
}

export function ensureOwnedWorksheet(workbook: Workbook, definition: WorksheetSchemaDefinition): Worksheet {
  let worksheet = workbook.getWorksheet(definition.sheetName);
  if (!worksheet) {
    worksheet = workbook.addWorksheet(definition.sheetName);
    worksheet.addRow(definition.fields.map((item) => item.header));
    worksheet.getRow(1).font = { bold: true };
    worksheet.views = [{ state: "frozen", ySplit: 1 }];
  }
  getWorksheetSchema(worksheet, definition);
  return worksheet;
}

export function ensureEvidenceField(worksheet: Worksheet, name: "evidence" | "evidenceUrl" | "evidenceStatus"): boolean {
  return ensureOptionalSchemaField(worksheet, name);
}

export function ensureOptionalSchemaField(worksheet: Worksheet, name: WorkbookField): boolean {
  const mapping = getWorksheetSchema(worksheet);
  if (mapping.fields[name]) return false;
  const definition = mapping.definition.fields.find((item) => item.field === name);
  if (!definition || definition.required) throw new Error(`Unsupported optional schema extension in ${worksheet.name}`);
  let lastColumn = 0;
  worksheet.eachRow((row) => row.eachCell((cell, column) => {
    if (cell.value !== null && cell.value !== undefined) lastColumn = Math.max(lastColumn, column);
  }));
  // ExcelJS 4.x exposes Table instances here despite its tuple-array declaration.
  const tables = worksheet.getTables() as unknown as Array<{ ref?: string; table: { tableRef?: string; columns: unknown[] } }>;
  for (const table of tables) {
    const range = table.table.tableRef;
    const letters = (range?.split(":").at(-1) ?? table.ref)?.match(/\$?([A-Z]+)\$?\d+/i)?.[1];
    if (!letters) throw new Error("Cannot locate workbook table boundary safely");
    lastColumn = Math.max(lastColumn, worksheet.getColumn(letters).number + (range ? 0 : table.table.columns.length - 1));
  }
  if (lastColumn >= 16_384) throw new Error("No safe column is available for evidence");
  const header = worksheet.getCell(mapping.headerRow, lastColumn + 1);
  if (lastColumn) header.style = structuredClone(worksheet.getCell(mapping.headerRow, lastColumn).style);
  header.value = definition.header;
  worksheet.getColumn(lastColumn + 1).width = name === "evidenceStatus" ? 24 : 20;
  return true;
}

export function mainSchemaFingerprint(worksheet: Worksheet): string {
  const mapping = getWorksheetSchema(worksheet);
  // Evidence is the one explicitly auto-created main-sheet extension.
  const fields = mapping.definition.fields.filter((item) => item.field !== "evidence").map((item) => [item.field, mapping.fields[item.field]?.columnIndex ?? null, normalizeWorkbookHeader(mapping.fields[item.field]?.header ?? "")]);
  return createHash("sha256").update(JSON.stringify(fields)).digest("hex");
}

export function formatWorksheetSchema(mapping: ResolvedWorksheetSchema): string {
  return [
    `Sheet: ${mapping.sheetName} (header row ${mapping.headerRow})`,
    ...mapping.definition.fields.map((item) => {
      const resolved = mapping.fields[item.field];
      return resolved ? `OK ${item.header}: ${resolved.columnLetter} - ${resolved.header}${resolved.match === "alias" || resolved.match === "override" ? ` (${resolved.match})` : ""}` : `${item.required ? "ERROR" : "WARN"} ${item.header}: not resolved`;
    }),
    ...mapping.issues.filter((issue) => issue.severity === "ERROR").map((issue) => `${issue.severity} ${issue.message}`),
  ].join("\n");
}
