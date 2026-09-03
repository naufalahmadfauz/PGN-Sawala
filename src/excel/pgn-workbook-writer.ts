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
  parsePgnWorkbook,
} from "./pgn-workbook-loader";
import {
  ensureEvidenceWorkbookSchema,
  writeEvidenceHyperlink,
} from "./evidence-workbook";
import {
  KB_SHEET_NAME,
  NEGATIVE_SHEET_NAME,
  TRANSCRIPT_SHEET_NAME,
  type ExecutedTurn,
  type PgnTestScenario,
  type PgnWorkbookDocument,
} from "./pgn-types";

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
  "Evidence URL",
  "Evidence Status",
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
    24, 18, 28, 12, 9, 16, 70, 24, 20, 20, 18, 45, 55, 20, 24,
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
  const definitions = [
    { sheetName: KB_SHEET_NAME, columns: [1, 2, 3, 4, 5, 6, 7, 8] },
    { sheetName: NEGATIVE_SHEET_NAME, columns: [1, 2, 3, 4, 5, 6, 7, 13] },
  ];
  return JSON.stringify(
    definitions.map(({ sheetName, columns }) => {
      const worksheet = workbook.getWorksheet(sheetName);
      if (!worksheet) {
        return { sheetName, missing: true };
      }
      const rows = Array.from({ length: worksheet.rowCount }, (_, rowIndex) =>
        columns.map((column) => cellText(worksheet.getCell(rowIndex + 1, column))),
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
    execution.evidenceUrl ? "View Evidence" : "",
    execution.evidenceStatus ?? "",
  ]);
  if (execution.evidenceUrl) {
    writeEvidenceHyperlink(userRow.getCell(14), execution.evidenceUrl);
  }
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
        execution.evidenceUrl ? "View Evidence" : "",
        execution.evidenceStatus ?? "",
      ]);
      if (execution.evidenceUrl) {
        writeEvidenceHyperlink(botRow.getCell(14), execution.evidenceUrl);
      }
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
      execution.evidenceUrl ? "View Evidence" : "",
      execution.evidenceStatus ?? "",
    ]);
    if (execution.evidenceUrl) {
      writeEvidenceHyperlink(errorRow.getCell(14), execution.evidenceUrl);
    }
    errorRow.alignment = { vertical: "top", wrapText: true };
    errorRow.getCell(8).numFmt = "yyyy-mm-dd hh:mm:ss";
  }
}

function applyKnowledgeBaseExecution(
  worksheet: Worksheet,
  scenario: PgnTestScenario,
  executions: ExecutedTurn[],
): void {
  for (const turn of scenario.turns) {
    for (const column of [9, 10, 11, 14]) {
      worksheet.getCell(turn.rowNumber, column).value = null;
    }
  }
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
    if (execution.evidenceUrl) {
      writeEvidenceHyperlink(worksheet.getCell(row, 14), execution.evidenceUrl);
    }
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
  for (const column of [8, 9, 10, 14]) {
    worksheet.getCell(row, column).value = null;
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
  const finalExecution = executions.at(-1);
  const expectedFinalTurn = scenario.turns.at(-1)?.turnNumber;
  if (
    finalExecution &&
    finalExecution.turn.turnNumber === expectedFinalTurn &&
    finalExecution.evidenceUrl
  ) {
    writeEvidenceHyperlink(
      worksheet.getCell(row, 14),
      finalExecution.evidenceUrl,
    );
  }
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
  if (path.resolve(sourcePath) === path.resolve(outputPath)) {
    throw new Error("Executed workbook path must differ from the immutable source");
  }
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
  expectedOutputHashes.set(workbook, bufferHash(outputContents));
  if (resumed) {
    await assertExecutedWorkbookMatchesSource(sourcePath, workbook);
  }
  preservedTableParts.set(workbook, tableParts);
  const hadTranscript = Boolean(
    workbook.getWorksheet(TRANSCRIPT_SHEET_NAME),
  );
  ensureTranscriptWorksheet(workbook);
  const evidenceSchemaChanged = ensureEvidenceWorkbookSchema(workbook);
  if (
    !hadTranscript ||
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
  const common = [
    runId,
    scenario.testCaseId,
    scenario.sheetName,
    scenario.sourceRowNumber,
    null,
  ];
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
      "",
      "",
    ]);
    row.alignment = { vertical: "top", wrapText: true };
    row.getCell(8).numFmt = "yyyy-mm-dd hh:mm:ss";
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
  const common = [
    runId,
    scenario.testCaseId,
    scenario.sheetName,
    scenario.sourceRowNumber,
    null,
  ];
  for (const staleMessage of drain.staleMessages) {
    const row = worksheet.addRow([
      ...common,
      "STALE_BOT",
      staleMessage.text,
      staleMessage.observedAt,
      null,
      null,
      "STALE_DRAINED",
      "",
      "",
      "",
      "",
    ]);
    row.alignment = { vertical: "top", wrapText: true };
    row.getCell(8).numFmt = "yyyy-mm-dd hh:mm:ss";
  }

  const completionRow = worksheet.addRow([
    ...common,
    "CONTROL_SYSTEM",
    `Post-reset quiet period confirmed: ${drain.quietMs} ms${drain.staleMessages.length ? `; stale messages drained: ${drain.staleMessages.length}` : ""}`,
    drain.completedAt,
    null,
    null,
    "QUIET_CONFIRMED",
    "",
    "",
    "",
    "",
  ]);
  completionRow.alignment = { vertical: "top", wrapText: true };
  completionRow.getCell(8).numFmt = "yyyy-mm-dd hh:mm:ss";
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
