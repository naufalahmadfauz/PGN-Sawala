import {
  constants as fsConstants,
  access,
  copyFile,
  mkdir,
  readFile,
  rename,
  rm,
  writeFile,
} from "node:fs/promises";
import { createHash } from "node:crypto";
import path from "node:path";
import ExcelJS, { type Cell, type Worksheet } from "exceljs";
import JSZip from "jszip";
import type {
  BotSessionResetAttempt,
  PostResetDrainResult,
} from "../types";
import {
  cellText,
  loadPgnWorkbook,
  parsePgnWorkbook,
} from "./pgn-workbook-loader";
import {
  ensureEvidenceWorkbookSchema,
  writeEvidenceHyperlink,
} from "./evidence-workbook";
import {
  TRANSCRIPT_SHEET_NAME,
  type ExecutedTurn,
  type PgnTestScenario,
  type PgnWorkbookDocument,
} from "./pgn-types";
import { attachWorkbookMappings } from "./workbook-mapping";
import { assertPgnWorkbookValid } from "./pgn-workbook-validator";
import {
  KB_SCHEMA, NEGATIVE_SCHEMA, TRANSCRIPT_SCHEMA, appendSchemaRow,
  fieldCell, fieldColumn, optionalFieldCell, getWorksheetSchema,
  ensureOptionalSchemaField,
} from "./workbook-schema";
import { runExecutionContext } from "./run-configuration";

interface PreservedTablePart {
  partPath: string;
  identity: string;
  structure: string;
  contents: Buffer;
}

type PreservedTableParts = Map<string, PreservedTablePart>;

const preservedTableParts = new WeakMap<ExcelJS.Workbook, PreservedTableParts>();
const expectedOutputHashes = new WeakMap<ExcelJS.Workbook, string>();

const TABLE_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml";

function ensureTranscriptWorksheet(workbook: ExcelJS.Workbook): Worksheet {
  const existing = workbook.getWorksheet(TRANSCRIPT_SHEET_NAME);
  if (existing) {
    for (const field of ["evidenceUrl", "evidenceStatus", "transport", "sessionMode", "conversationId", "dialogId"] as const) {
      ensureOptionalSchemaField(existing, field);
    }
    const role = fieldColumn(existing, "role");
    existing.getColumn(role).width = Math.max(
      existing.getColumn(role).width ?? 0,
      16,
    );
    return existing;
  }

  const worksheet = workbook.addWorksheet(TRANSCRIPT_SHEET_NAME);
  const header = worksheet.addRow(TRANSCRIPT_SCHEMA.fields.map((item) => item.header));
  header.font = { bold: true, color: { argb: "FFFFFFFF" } };
  header.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FF1F4E78" },
  };
  header.alignment = { vertical: "middle", wrapText: true };
  worksheet.views = [{ state: "frozen", ySplit: 1 }];
  const widths = [
    24, 18, 28, 12, 9, 16, 70, 24, 20, 20, 18, 45, 55, 20, 24, 16, 18, 38, 38,
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

function setNumberFormat(cell: Cell, numberFormat: string): void {
  cell.style = { ...structuredClone(cell.style), numFmt: numberFormat };
}

function writeExecutionDate(cell: Cell, value: Date): void {
  cell.value = value;
  setNumberFormat(cell, "yyyy-mm-dd hh:mm:ss");
}

function secondsFromMilliseconds(milliseconds: number): number {
  return Math.round((milliseconds / 1_000) * 100) / 100;
}

async function fileHash(filePath: string): Promise<string> {
  return createHash("sha256").update(await readFile(filePath)).digest("hex");
}

function bufferHash(contents: Uint8Array): string {
  return createHash("sha256").update(contents).digest("hex");
}

function decodeXmlAttribute(value: string): string {
  return value
    .replace(/&quot;/gi, '"')
    .replace(/&apos;/gi, "'")
    .replace(/&lt;/gi, "<")
    .replace(/&gt;/gi, ">")
    .replace(/&amp;/gi, "&")
    .replace(/&#x([0-9a-f]+);/gi, (_, digits: string) =>
      String.fromCodePoint(Number.parseInt(digits, 16)),
    )
    .replace(/&#([0-9]+);/g, (_, digits: string) =>
      String.fromCodePoint(Number.parseInt(digits, 10)),
    );
}

function xmlAttribute(tag: string, name: string): string | undefined {
  const match = tag.match(
    new RegExp(`\\b${name}\\s*=\\s*(?:"([^"]*)"|'([^']*)')`, "i"),
  );
  const value = match?.[1] ?? match?.[2];
  return value === undefined ? undefined : decodeXmlAttribute(value);
}

function describeTablePart(partPath: string, contents: Buffer): PreservedTablePart {
  const xml = contents.toString("utf8");
  const tableTag = xml.match(/<table(?:\s|>)[^>]*>/i)?.[0];
  if (!tableTag) {
    throw new Error(`XLSX table part is invalid: ${partPath}`);
  }
  const identity = xmlAttribute(tableTag, "displayName") ?? xmlAttribute(tableTag, "name");
  const ref = xmlAttribute(tableTag, "ref");
  if (!identity || !ref) {
    throw new Error(`XLSX table part lacks displayName/name or ref: ${partPath}`);
  }
  const columnNames = [...xml.matchAll(/<tableColumn(?:\s|>)[^>]*>/gi)].map(
    ([tag]) => xmlAttribute(tag, "name") ?? "",
  );
  return {
    partPath,
    identity,
    structure: JSON.stringify({ identity, ref, columnNames }),
    contents,
  };
}

async function validateTablePackageTopology(
  archive: JSZip,
  tableParts: PreservedTableParts,
): Promise<void> {
  const partPaths = new Set(
    [...tableParts.values()].map((tablePart) => tablePart.partPath),
  );
  const referencedPaths = new Set<string>();
  for (const [relationshipPath, entry] of Object.entries(archive.files)) {
    if (!/^xl\/worksheets\/_rels\/[^/]+\.rels$/i.test(relationshipPath) || entry.dir) {
      continue;
    }
    const relationshipsXml = await entry.async("string");
    for (const [tag] of relationshipsXml.matchAll(/<Relationship(?:\s|>)[^>]*\/?\s*>/gi)) {
      const type = xmlAttribute(tag, "Type");
      if (!type?.endsWith("/table")) {
        continue;
      }
      const target = xmlAttribute(tag, "Target");
      if (!target) {
        throw new Error(`XLSX table relationship lacks a target: ${relationshipPath}`);
      }
      const worksheetDirectory = path.posix.dirname(
        path.posix.dirname(relationshipPath),
      );
      referencedPaths.add(
        target.startsWith("/")
          ? target.slice(1)
          : path.posix.normalize(path.posix.join(worksheetDirectory, target)),
      );
    }
  }

  const contentTypesEntry = archive.file("[Content_Types].xml");
  if (!contentTypesEntry) {
    throw new Error("XLSX package lacks [Content_Types].xml");
  }
  const contentTypesXml = await contentTypesEntry.async("string");
  const tableContentTypes = new Map<string, string>();
  for (const [tag] of contentTypesXml.matchAll(/<Override(?:\s|>)[^>]*\/?\s*>/gi)) {
    const partName = xmlAttribute(tag, "PartName");
    const contentType = xmlAttribute(tag, "ContentType");
    if (partName && contentType) {
      tableContentTypes.set(partName.replace(/^\//, ""), contentType);
    }
  }

  for (const partPath of partPaths) {
    if (!referencedPaths.has(partPath)) {
      throw new Error(`XLSX table part is not referenced by a worksheet: ${partPath}`);
    }
    if (tableContentTypes.get(partPath) !== TABLE_CONTENT_TYPE) {
      throw new Error(`XLSX table part has no valid content type: ${partPath}`);
    }
  }
  for (const referencedPath of referencedPaths) {
    if (!partPaths.has(referencedPath)) {
      throw new Error(`XLSX worksheet references a missing table part: ${referencedPath}`);
    }
  }
}

async function tablePartsFromArchive(archive: JSZip): Promise<PreservedTableParts> {
  const tableParts: PreservedTableParts = new Map();
  for (const [partPath, entry] of Object.entries(archive.files)) {
    if (/^xl\/tables\/[^/]+\.xml$/i.test(partPath) && !entry.dir) {
      const tablePart = describeTablePart(
        partPath,
        await entry.async("nodebuffer"),
      );
      if (tableParts.has(tablePart.identity)) {
        throw new Error(
          `XLSX contains duplicate table identity "${tablePart.identity}"`,
        );
      }
      tableParts.set(tablePart.identity, tablePart);
    }
  }
  await validateTablePackageTopology(archive, tableParts);
  return tableParts;
}

async function readTableParts(filePath: string): Promise<PreservedTableParts> {
  return tablePartsFromArchive(await JSZip.loadAsync(await readFile(filePath)));
}

function assertCompatibleTableParts(
  expected: PreservedTableParts,
  actual: PreservedTableParts,
): void {
  if (expected.size !== actual.size) {
    throw new Error(
      `Executed workbook table count (${actual.size}) does not match source (${expected.size})`,
    );
  }
  for (const [identity, expectedPart] of expected) {
    const actualPart = actual.get(identity);
    if (!actualPart || actualPart.structure !== expectedPart.structure) {
      throw new Error(
        `Executed workbook table "${identity}" does not match the source structure`,
      );
    }
  }
}

async function tablePartsMatch(
  filePath: string,
  expected: PreservedTableParts,
): Promise<boolean> {
  const actual = await readTableParts(filePath);
  assertCompatibleTableParts(expected, actual);
  return [...expected].every(([identity, expectedPart]) =>
    actual.get(identity)?.contents.equals(expectedPart.contents),
  );
}

async function restoreTableParts(
  filePath: string,
  tableParts: PreservedTableParts,
): Promise<void> {
  const archive = await JSZip.loadAsync(await readFile(filePath));
  const generatedTableParts = await tablePartsFromArchive(archive);
  assertCompatibleTableParts(tableParts, generatedTableParts);
  for (const [identity, generatedPart] of generatedTableParts) {
    archive.file(generatedPart.partPath, tableParts.get(identity)!.contents);
  }
  const repaired = await archive.generateAsync({
    type: "nodebuffer",
    compression: "DEFLATE",
  });
  await writeFile(filePath, repaired);
  if (!(await tablePartsMatch(filePath, tableParts))) {
    throw new Error("XLSX table preservation verification failed");
  }
}

function sourceOwnedCells(workbook: ExcelJS.Workbook): string {
  const definitions = [KB_SCHEMA, NEGATIVE_SCHEMA];
  return JSON.stringify(
    definitions.map(({ sheetName, fields }) => {
      const worksheet = workbook.getWorksheet(sheetName);
      if (!worksheet) {
        return { sheetName, missing: true };
      }
      const rows = Array.from({ length: worksheet.rowCount - 1 }, (_, rowIndex) =>
        fields.filter((field) => !field.writable).map((field) => optionalFieldCell(worksheet, rowIndex + 2, field.field)?.text ?? ""),
      );
      return { sheetName, rows };
    }),
  );
}

async function assertExecutedWorkbookMatchesSource(
  sourcePath: string,
  executedWorkbook: ExcelJS.Workbook,
): Promise<void> {
  const sourceWorkbook = new ExcelJS.Workbook();
  await sourceWorkbook.xlsx.readFile(sourcePath);
  await attachWorkbookMappings(sourceWorkbook, sourcePath);
  if (sourceOwnedCells(sourceWorkbook) !== sourceOwnedCells(executedWorkbook)) {
    throw new Error(
      "Executed workbook inputs do not match the source workbook; use a new output path",
    );
  }
}

function appendTranscriptRows(
  worksheet: Worksheet,
  runId: string,
  scenario: PgnTestScenario,
  execution: ExecutedTurn,
): void {
  const common = {
    ...runExecutionContext(worksheet.workbook, runId),
    conversationId: execution.conversationId ?? "", dialogId: execution.dialogId ?? "",
    runId, testCaseId: scenario.testCaseId, sheet: scenario.sheetName,
    excelRow: execution.turn.rowNumber, turn: execution.turn.turnNumber,
    status: execution.technicalStatus, error: execution.error ?? "",
    evidencePath: execution.evidencePath ?? "", evidenceStatus: execution.evidenceStatus ?? "",
  };
  const userRow = appendSchemaRow(worksheet, {
    ...common, role: "USER", message: execution.turn.userInput,
    timestamp: execution.sentAt ?? execution.completedAt,
  });
  if (execution.evidenceUrl) {
    writeEvidenceHyperlink(fieldCell(worksheet, userRow.number, "evidenceUrl"), execution.evidenceUrl);
  }
  userRow.alignment = { vertical: "top", wrapText: true };
  fieldCell(worksheet, userRow.number, "timestamp").numFmt = "yyyy-mm-dd hh:mm:ss";

  if (execution.botMessages.length > 0) {
    for (const botMessage of execution.botMessages) {
      const botRow = appendSchemaRow(worksheet, {
        ...common, role: "BOT", message: botMessage.message, timestamp: botMessage.timestamp,
        firstResponseMs: execution.firstResponseMs ?? null, totalResponseMs: execution.totalResponseMs ?? null,
      });
      if (execution.evidenceUrl) {
        writeEvidenceHyperlink(fieldCell(worksheet, botRow.number, "evidenceUrl"), execution.evidenceUrl);
      }
      botRow.alignment = { vertical: "top", wrapText: true };
      fieldCell(worksheet, botRow.number, "timestamp").numFmt = "yyyy-mm-dd hh:mm:ss";
    }
  } else if (execution.error) {
    const errorRow = appendSchemaRow(worksheet, {
      ...common, role: "SYSTEM", message: execution.error, timestamp: execution.completedAt,
      firstResponseMs: execution.firstResponseMs ?? null, totalResponseMs: execution.totalResponseMs ?? null,
    });
    if (execution.evidenceUrl) {
      writeEvidenceHyperlink(fieldCell(worksheet, errorRow.number, "evidenceUrl"), execution.evidenceUrl);
    }
    errorRow.alignment = { vertical: "top", wrapText: true };
    fieldCell(worksheet, errorRow.number, "timestamp").numFmt = "yyyy-mm-dd hh:mm:ss";
  }
}

function applyKnowledgeBaseExecution(
  worksheet: Worksheet,
  scenario: PgnTestScenario,
  executions: ExecutedTurn[],
): void {
  for (const turn of scenario.turns) {
    for (const field of ["botResponse", "responseTime", "testDate", "evidence"] as const) {
      const cell = optionalFieldCell(worksheet, turn.rowNumber, field);
      if (cell) cell.value = null;
    }
  }
  for (const execution of executions) {
    const row = execution.turn.rowNumber;
    if (execution.combinedResponse) {
      fieldCell(worksheet, row, "botResponse").value = execution.combinedResponse;
    }
    const responseTime = optionalFieldCell(worksheet, row, "responseTime");
    if (execution.totalResponseMs !== undefined && responseTime) {
      responseTime.value = secondsFromMilliseconds(execution.totalResponseMs);
      setNumberFormat(responseTime, '0.00" s"');
    }
    const date = optionalFieldCell(worksheet, row, "testDate");
    if (date) writeExecutionDate(date, execution.completedAt);
    const evidence = optionalFieldCell(worksheet, row, "evidence");
    if (execution.evidenceUrl && evidence) {
      writeEvidenceHyperlink(evidence, execution.evidenceUrl);
    }
    const notes = optionalFieldCell(worksheet, row, "notes");
    if (execution.technicalStatus !== "CAPTURED" && notes) {
      appendTechnicalNote(
        notes,
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
  for (const field of ["botResponse", "responseTime", "testDate", "evidence"] as const) {
    const cell = optionalFieldCell(worksheet, row, field);
    if (cell) cell.value = null;
  }
  const isMultiTurn = scenario.turns.length > 1;
  const allTurnsExecuted = executions.length === scenario.turns.length;
  const latest = executions.at(-1);
  if (
    !allTurnsExecuted &&
    (!latest || latest.technicalStatus === "CAPTURED")
  ) {
    return;
  }
  const allTurnsHaveResponses =
    allTurnsExecuted &&
    executions.every((execution) => Boolean(execution.combinedResponse));
  if (allTurnsHaveResponses) {
    fieldCell(worksheet, row, "botResponse").value = isMultiTurn
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
  const responseTime = optionalFieldCell(worksheet, row, "responseTime");
  if (allTurnsHaveTiming && responseTime) {
    if (isMultiTurn) {
      responseTime.value = executions
        .map(
          (execution) =>
            `Turn ${execution.turn.turnNumber}: ${execution.totalResponseMs} ms`,
        )
        .join("\n");
      setNumberFormat(responseTime, "@");
    } else {
      responseTime.value = secondsFromMilliseconds(
        executions[0].totalResponseMs!,
      );
      setNumberFormat(responseTime, "0.00");
    }
  }

  const completedAt = executions.at(-1)?.completedAt ?? new Date();
  const date = optionalFieldCell(worksheet, row, "testDate");
  if (date) writeExecutionDate(date, completedAt);
  const finalExecution = executions.at(-1);
  const expectedFinalTurn = scenario.turns.at(-1)?.turnNumber;
  const evidence = optionalFieldCell(worksheet, row, "evidence");
  if (
    finalExecution &&
    finalExecution.turn.turnNumber === expectedFinalTurn &&
    finalExecution.evidenceUrl && evidence
  ) {
    writeEvidenceHyperlink(
      evidence,
      finalExecution.evidenceUrl,
    );
  }
  for (const execution of executions) {
    const notes = optionalFieldCell(worksheet, row, "notes");
    if (execution.technicalStatus !== "CAPTURED" && notes) {
      appendTechnicalNote(
        notes,
        `[Technical execution ${execution.completedAt.toISOString()}] Turn ${execution.turn.turnNumber}: ${execution.technicalStatus}${execution.error ? ` - ${execution.error}` : ""}`,
      );
    }
  }
}

export async function openExecutedPgnWorkbook(
  sourcePath: string,
  outputPath: string,
): Promise<PgnWorkbookDocument & { resumed: boolean }> {
  if (path.resolve(sourcePath) === path.resolve(outputPath)) {
    throw new Error("Executed workbook path must differ from the immutable source");
  }
  assertPgnWorkbookValid((await loadPgnWorkbook(sourcePath)).parsed);
  await mkdir(path.dirname(outputPath), { recursive: true });
  const tableParts = await readTableParts(sourcePath);
  const resumed = await access(outputPath)
    .then(() => true)
    .catch(() => false);
  if (!resumed) {
    await copyFile(sourcePath, outputPath, fsConstants.COPYFILE_EXCL);
  } else {
    assertCompatibleTableParts(tableParts, await readTableParts(outputPath));
  }

  const workbook = new ExcelJS.Workbook();
  const outputContents = await readFile(outputPath);
  await workbook.xlsx.load(
    outputContents as unknown as Parameters<typeof workbook.xlsx.load>[0],
  );
  await attachWorkbookMappings(workbook, outputPath, sourcePath);
  assertPgnWorkbookValid(parsePgnWorkbook(workbook));
  expectedOutputHashes.set(workbook, bufferHash(outputContents));
  if (resumed) {
    await assertExecutedWorkbookMatchesSource(sourcePath, workbook);
  }
  preservedTableParts.set(workbook, tableParts);
  const hadTranscript = Boolean(
    workbook.getWorksheet(TRANSCRIPT_SHEET_NAME),
  );
  const previousTranscriptSchema = hadTranscript
    ? getWorksheetSchema(workbook.getWorksheet(TRANSCRIPT_SHEET_NAME)!).fingerprint
    : undefined;
  ensureTranscriptWorksheet(workbook);
  const evidenceSchemaChanged = ensureEvidenceWorkbookSchema(workbook);
  if (
    !hadTranscript ||
    previousTranscriptSchema !== getWorksheetSchema(workbook.getWorksheet(TRANSCRIPT_SHEET_NAME)!).fingerprint ||
    evidenceSchemaChanged ||
    !(await tablePartsMatch(outputPath, tableParts))
  ) {
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
  applyScenarioResults(workbook, scenario, executions);
  const transcript = ensureTranscriptWorksheet(workbook);
  for (const execution of executions) {
    appendTranscriptRows(transcript, runId, scenario, execution);
  }
}

export function applyScenarioResults(
  workbook: ExcelJS.Workbook,
  scenario: PgnTestScenario,
  executions: ExecutedTurn[],
): void {
  const worksheet = workbook.getWorksheet(scenario.sheetName);
  if (!worksheet) {
    throw new Error(`Worksheet "${scenario.sheetName}" was not found`);
  }
  getWorksheetSchema(worksheet);
  if (scenario.sheetKind === "kb") {
    applyKnowledgeBaseExecution(worksheet, scenario, executions);
  } else {
    applyNegativeExecution(worksheet, scenario, executions);
  }
}

export function appendLatestTurnExecution(
  workbook: ExcelJS.Workbook,
  runId: string,
  scenario: PgnTestScenario,
  executions: ExecutedTurn[],
): void {
  const latest = executions.at(-1);
  if (!workbook.getWorksheet(scenario.sheetName) || !latest) {
    throw new Error(`Cannot append latest execution for "${scenario.testCaseId}"`);
  }
  appendTranscriptRows(
    ensureTranscriptWorksheet(workbook),
    runId,
    scenario,
    latest,
  );
}

export function appendSessionResetTranscript(
  workbook: ExcelJS.Workbook,
  runId: string,
  scenario: PgnTestScenario,
  attempt: BotSessionResetAttempt,
): void {
  const worksheet = ensureTranscriptWorksheet(workbook);
  const common = { ...runExecutionContext(workbook, runId), runId, testCaseId: scenario.testCaseId, sheet: scenario.sheetName, excelRow: scenario.sourceRowNumber };
  const appendRow = (
    role:
      | "CONTROL_USER"
      | "CONTROL_BOT"
      | "CONTROL_SYSTEM"
      | "STALE_BOT",
    message: string,
    timestamp: Date,
    firstResponseMs?: number,
    totalResponseMs?: number,
  ): void => {
    const row = appendSchemaRow(worksheet, {
      ...common, role, message, timestamp,
      firstResponseMs: firstResponseMs ?? null, totalResponseMs: totalResponseMs ?? null,
      status: attempt.status, error: attempt.error ?? "", evidencePath: attempt.evidencePath ?? "",
    });
    row.alignment = { vertical: "top", wrapText: true };
    fieldCell(worksheet, row.number, "timestamp").numFmt = "yyyy-mm-dd hh:mm:ss";
  };

  if (attempt.sentAt) {
    appendRow("CONTROL_USER", attempt.command, attempt.sentAt);
  }
  const controlMessageIndexes = resetConfirmationMessageIndexes(attempt);
  for (const [index, response] of attempt.responseMessages.entries()) {
    appendRow(
      controlMessageIndexes.has(index) ? "CONTROL_BOT" : "STALE_BOT",
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

function normalizeTranscriptText(value: string): string {
  return value.replace(/\s+/g, " ").trim().toLowerCase();
}

function resetConfirmationMessageIndexes(
  attempt: BotSessionResetAttempt,
): Set<number> {
  const expected = normalizeTranscriptText(attempt.expectedConfirmation);
  for (let end = 0; end < attempt.responseMessages.length; end += 1) {
    for (let start = end; start >= 0; start -= 1) {
      const combined = attempt.responseMessages
        .slice(start, end + 1)
        .map((message) => message.text)
        .join("\n");
      if (normalizeTranscriptText(combined).includes(expected)) {
        return new Set(
          Array.from({ length: end - start + 1 }, (_, index) => start + index),
        );
      }
    }
  }
  return new Set();
}

export function appendPostResetDrainTranscript(
  workbook: ExcelJS.Workbook,
  runId: string,
  scenario: PgnTestScenario,
  drain: PostResetDrainResult,
): void {
  const worksheet = ensureTranscriptWorksheet(workbook);
  const common = { ...runExecutionContext(workbook, runId), runId, testCaseId: scenario.testCaseId, sheet: scenario.sheetName, excelRow: scenario.sourceRowNumber };
  for (const staleMessage of drain.staleMessages) {
    const row = appendSchemaRow(worksheet, { ...common, role: "STALE_BOT", message: staleMessage.text, timestamp: staleMessage.observedAt, status: "STALE_DRAINED" });
    row.alignment = { vertical: "top", wrapText: true };
    fieldCell(worksheet, row.number, "timestamp").numFmt = "yyyy-mm-dd hh:mm:ss";
  }

  const completionRow = appendSchemaRow(worksheet, {
    ...common, role: "CONTROL_SYSTEM",
    message: `Post-reset quiet period confirmed: ${drain.quietMs} ms${drain.staleMessages.length ? `; stale messages drained: ${drain.staleMessages.length}` : ""}`,
    timestamp: drain.completedAt, status: "QUIET_CONFIRMED",
  });
  completionRow.alignment = { vertical: "top", wrapText: true };
  fieldCell(worksheet, completionRow.number, "timestamp").numFmt = "yyyy-mm-dd hh:mm:ss";
}

export type RecoveryTranscriptEvent =
  | "RUN_PREPARED"
  | "RUN_RESUMED"
  | "RUN_INTERRUPTED"
  | "RUN_FAILED"
  | "RUN_COMPLETED"
  | "RUN_ABANDONED"
  | "RUN_RESTARTED"
  | "RECOVERY_RECONCILED"
  | "SCENARIO_ATTEMPT_STARTED"
  | "SCENARIO_ATTEMPT_COMPLETED"
  | "SCENARIO_ATTEMPT_FAILED"
  | "SCENARIO_SKIPPED_BY_OPERATOR";

export function appendRecoveryTranscriptEvent(
  workbook: ExcelJS.Workbook,
  options: {
    runId: string;
    event: RecoveryTranscriptEvent;
    message: string;
    timestamp?: Date;
    scenario?: PgnTestScenario;
  },
): void {
  const worksheet = ensureTranscriptWorksheet(workbook);
  const row = appendSchemaRow(worksheet, {
    ...runExecutionContext(workbook, options.runId),
    runId: options.runId, testCaseId: options.scenario?.testCaseId ?? "",
    sheet: options.scenario?.sheetName ?? "", excelRow: options.scenario?.sourceRowNumber ?? null,
    role: "RECOVERY_SYSTEM", message: options.message, timestamp: options.timestamp ?? new Date(), status: options.event,
  });
  row.alignment = { vertical: "top", wrapText: true };
  fieldCell(worksheet, row.number, "timestamp").numFmt = "yyyy-mm-dd hh:mm:ss";
}

export async function saveExecutedPgnWorkbook(
  workbook: ExcelJS.Workbook,
  outputPath: string,
): Promise<void> {
  const temporaryPath = `${outputPath}.${process.pid}.tmp`;
  try {
    const expectedHash = expectedOutputHashes.get(workbook);
    if (expectedHash && (await fileHash(outputPath)) !== expectedHash) {
      throw new Error(
        "Executed workbook changed outside this process; refusing to overwrite it",
      );
    }
    await workbook.xlsx.writeFile(temporaryPath);
    const tableParts = preservedTableParts.get(workbook);
    if (tableParts) {
      await restoreTableParts(temporaryPath, tableParts);
    }
    if (expectedHash && (await fileHash(outputPath)) !== expectedHash) {
      throw new Error(
        "Executed workbook changed while saving; refusing to overwrite it",
      );
    }
    await rename(temporaryPath, outputPath);
    expectedOutputHashes.set(workbook, await fileHash(outputPath));
  } catch (error) {
    await rm(temporaryPath, { force: true }).catch(() => undefined);
    throw error;
  }
}
