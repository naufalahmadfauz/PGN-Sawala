import assert from "node:assert/strict";
import {
  mkdir,
  mkdtemp,
  readFile,
  rm,
  utimes,
  writeFile,
} from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import test from "node:test";
import ExcelJS from "exceljs";
import JSZip from "jszip";
import {
  applyScenarioResults,
  appendPostResetDrainTranscript,
  appendSessionResetTranscript,
  openExecutedPgnWorkbook,
  saveExecutedPgnWorkbook,
} from "../src/excel/pgn-workbook-writer";
import { readEvidenceHyperlink } from "../src/excel/evidence-workbook";
import { acquireWorkbookLock } from "../src/excel/workbook-lock";
import type {
  ExecutedTurn,
  PgnTestScenario,
} from "../src/excel/pgn-types";
import type { BotSessionResetAttempt, WhatsAppMessage } from "../src/types";

const sourcePath = path.resolve(
  "data/PGN AI Assistant - Knowledge Base Testing Report - User Inputs.xlsx",
);

async function tableXml(filePath: string): Promise<Buffer> {
  const archive = await JSZip.loadAsync(await readFile(filePath));
  const table = archive.file("xl/tables/table1.xml");
  assert(table, "xl/tables/table1.xml must exist");
  return table.async("nodebuffer");
}

async function writeArchive(archive: JSZip, filePath: string): Promise<void> {
  await writeFile(
    filePath,
    await archive.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }),
  );
}

async function tableParts(filePath: string): Promise<Map<string, Buffer>> {
  const archive = await JSZip.loadAsync(await readFile(filePath));
  const parts = new Map<string, Buffer>();
  for (const [partPath, entry] of Object.entries(archive.files)) {
    if (/^xl\/tables\/[^/]+\.xml$/i.test(partPath) && !entry.dir) {
      parts.set(partPath, await entry.async("nodebuffer"));
    }
  }
  return parts;
}

async function createRenamedTableSource(
  originalPath: string,
  renamedPath: string,
): Promise<void> {
  const archive = await JSZip.loadAsync(await readFile(originalPath));
  const table = archive.file("xl/tables/table1.xml");
  assert(table);
  const contents = await table.async("nodebuffer");
  archive.remove("xl/tables/table1.xml");
  archive.file("xl/tables/table7.xml", contents);

  const relationship = archive.file("xl/worksheets/_rels/sheet4.xml.rels");
  assert(relationship);
  archive.file(
    "xl/worksheets/_rels/sheet4.xml.rels",
    (await relationship.async("string")).replace(
      "../tables/table1.xml",
      "../tables/table7.xml",
    ),
  );
  const contentTypes = archive.file("[Content_Types].xml");
  assert(contentTypes);
  archive.file(
    "[Content_Types].xml",
    (await contentTypes.async("string")).replace(
      "/xl/tables/table1.xml",
      "/xl/tables/table7.xml",
    ),
  );
  await writeArchive(archive, renamedPath);
}

test("executed workbook preserves source table XML and records stale drain traffic", async (context) => {
  const temporaryDirectory = await mkdtemp(
    path.join(tmpdir(), "pgn-workbook-writer-"),
  );
  context.after(() => rm(temporaryDirectory, { recursive: true, force: true }));
  const outputPath = path.join(temporaryDirectory, "executed.xlsx");
  const sourceBefore = await readFile(sourcePath);
  const opened = await openExecutedPgnWorkbook(sourcePath, outputPath);
  const scenario = opened.parsed.scenarios.find(
    (item) => item.testCaseId === "PGN-KB-075",
  );
  assert(scenario);
  const worksheet = opened.workbook.getWorksheet(scenario.sheetName);
  assert(worksheet);
  worksheet.getCell(scenario.sourceRowNumber, 9).value =
    "Message 1:\nFirst\n\nMessage 2:\nSecond";
  worksheet.getCell(scenario.sourceRowNumber, 10).value = 11;
  worksheet.getCell(scenario.sourceRowNumber, 11).value = new Date(0);

  const staleMessage: WhatsAppMessage = {
    id: "stale-test",
    direction: "incoming",
    text: "Delayed stale message",
    domIndex: 1,
    observedAt: new Date(4_000),
  };
  appendPostResetDrainTranscript(opened.workbook, "xlsx-test", scenario, {
    startedAt: new Date(0),
    completedAt: new Date(14_000),
    quietMs: 10_000,
    staleMessages: [staleMessage],
  });
  const resetConfirmation: WhatsAppMessage = {
    id: "reset-confirmation",
    direction: "incoming",
    text: "Session deleted",
    domIndex: 2,
    observedAt: new Date(1_000),
  };
  const samePollStale: WhatsAppMessage = {
    id: "same-poll-stale",
    direction: "incoming",
    text: "Late stale response",
    domIndex: 3,
    observedAt: new Date(1_000),
  };
  const resetAttempt: BotSessionResetAttempt = {
    command: "reset",
    expectedConfirmation: "Session deleted",
    status: "RESET_CONFIRMED",
    startedAt: new Date(0),
    sentAt: new Date(500),
    completedAt: new Date(1_000),
    responseMessages: [resetConfirmation, samePollStale],
  };
  appendSessionResetTranscript(
    opened.workbook,
    "xlsx-test",
    scenario,
    resetAttempt,
  );
  await saveExecutedPgnWorkbook(opened.workbook, outputPath);

  const [sourceTable, outputTable] = await Promise.all([
    tableXml(sourcePath),
    tableXml(outputPath),
  ]);
  assert(outputTable.equals(sourceTable));
  const outputTableText = outputTable.toString("utf8");
  assert(!outputTableText.includes('headerRowCount="0"'));
  assert(!outputTableText.includes("<autoFilter"));

  const reloaded = new ExcelJS.Workbook();
  await reloaded.xlsx.readFile(outputPath);
  const reloadedWorksheet = reloaded.getWorksheet(scenario.sheetName);
  assert.equal(
    reloadedWorksheet?.getCell(scenario.sourceRowNumber, 9).text,
    "Message 1:\nFirst\n\nMessage 2:\nSecond",
  );
  const transcript = reloaded.getWorksheet("Execution Transcript");
  assert(transcript);
  const roles = (transcript.getRows(2, transcript.rowCount - 1) ?? []).map(
    (row) => row.getCell(6).text,
  );
  assert(roles.includes("STALE_BOT"));
  assert(roles.includes("CONTROL_SYSTEM"));
  const roleByMessage = new Map(
    (transcript.getRows(2, transcript.rowCount - 1) ?? []).map((row) => [
      row.getCell(7).text,
      row.getCell(6).text,
    ]),
  );
  assert.equal(roleByMessage.get("Session deleted"), "CONTROL_BOT");
  assert.equal(roleByMessage.get("Late stale response"), "STALE_BOT");

  const corrupted = await JSZip.loadAsync(await readFile(outputPath));
  const corruptedTable = corrupted.file("xl/tables/table1.xml");
  assert(corruptedTable);
  const malformedXml = (await corruptedTable.async("string"))
    .replace(' ref="A1:M19"', ' ref="A1:M19" totalsRowShown="1" headerRowCount="0"')
    .replace(
      "<tableColumns",
      '<autoFilter><filterColumn colId="0" hiddenButton="1"/></autoFilter><tableColumns',
    );
  corrupted.file("xl/tables/table1.xml", malformedXml);
  await writeArchive(corrupted, outputPath);

  const resumed = await openExecutedPgnWorkbook(sourcePath, outputPath);
  assert.equal(resumed.resumed, true);
  assert((await tableXml(outputPath)).equals(sourceTable));
  assert.equal(
    resumed.workbook
      .getWorksheet(scenario.sheetName)
      ?.getCell(scenario.sourceRowNumber, 9).text,
    "Message 1:\nFirst\n\nMessage 2:\nSecond",
  );
  assert((await readFile(sourcePath)).equals(sourceBefore));
});

test("table restoration follows generated relationships instead of source filenames", async (context) => {
  const temporaryDirectory = await mkdtemp(
    path.join(tmpdir(), "pgn-table-topology-"),
  );
  context.after(() => rm(temporaryDirectory, { recursive: true, force: true }));
  const renamedSourcePath = path.join(temporaryDirectory, "source-table7.xlsx");
  const changedSourcePath = path.join(temporaryDirectory, "source-changed.xlsx");
  const outputPath = path.join(temporaryDirectory, "executed.xlsx");
  await createRenamedTableSource(sourcePath, renamedSourcePath);

  const opened = await openExecutedPgnWorkbook(renamedSourcePath, outputPath);
  await saveExecutedPgnWorkbook(opened.workbook, outputPath);
  const sourceParts = await tableParts(renamedSourcePath);
  const outputParts = await tableParts(outputPath);
  assert.deepEqual([...sourceParts.keys()], ["xl/tables/table7.xml"]);
  assert.deepEqual([...outputParts.keys()], ["xl/tables/table1.xml"]);
  assert(outputParts.get("xl/tables/table1.xml")?.equals([...sourceParts.values()][0]));

  const outputArchive = await JSZip.loadAsync(await readFile(outputPath));
  const relationship = outputArchive.file("xl/worksheets/_rels/sheet4.xml.rels");
  const contentTypes = outputArchive.file("[Content_Types].xml");
  assert(relationship);
  assert(contentTypes);
  assert.match(await relationship.async("string"), /Target="\.\.\/tables\/table1\.xml"/);
  assert.match(await contentTypes.async("string"), /PartName="\/xl\/tables\/table1\.xml"/);

  const changedArchive = await JSZip.loadAsync(await readFile(renamedSourcePath));
  const changedTable = changedArchive.file("xl/tables/table7.xml");
  assert(changedTable);
  changedArchive.file(
    "xl/tables/table7.xml",
    (await changedTable.async("string")).replace('ref="A1:M19"', 'ref="A1:M18"'),
  );
  await writeArchive(changedArchive, changedSourcePath);
  const outputBefore = await readFile(outputPath);
  await assert.rejects(
    openExecutedPgnWorkbook(changedSourcePath, outputPath),
    /does not match the source structure/,
  );
  assert((await readFile(outputPath)).equals(outputBefore));
});

function capturedExecution(
  scenario: PgnTestScenario,
  turnIndex: number,
  evidenceUrl?: string,
): ExecutedTurn {
  const turn = scenario.turns[turnIndex];
  return {
    turn,
    technicalStatus: "CAPTURED",
    sentAt: new Date("2026-09-02T01:00:00Z"),
    completedAt: new Date("2026-09-02T01:00:01Z"),
    botMessages: [],
    combinedResponse: `New response turn ${turn.turnNumber}`,
    firstResponseMs: 500,
    totalResponseMs: 1_000,
    evidenceUrl,
    evidenceStatus: evidenceUrl
      ? "EVIDENCE_SYNCED"
      : "EVIDENCE_UPLOAD_ERROR",
  };
}

test("partial and upload-failed reruns clear stale result evidence", async (context) => {
  const temporaryDirectory = await mkdtemp(
    path.join(tmpdir(), "pgn-stale-rerun-"),
  );
  context.after(() => rm(temporaryDirectory, { recursive: true, force: true }));
  const outputPath = path.join(temporaryDirectory, "executed.xlsx");
  const opened = await openExecutedPgnWorkbook(sourcePath, outputPath);
  const kbScenario = opened.parsed.scenarios.find(
    (scenario) => scenario.testCaseId === "PGN-KB-031",
  )!;
  const kb = opened.workbook.getWorksheet(kbScenario.sheetName)!;
  for (const turn of kbScenario.turns) {
    kb.getCell(turn.rowNumber, 9).value = "Old response";
    kb.getCell(turn.rowNumber, 10).value = 99;
    kb.getCell(turn.rowNumber, 11).value = new Date(0);
    kb.getCell(turn.rowNumber, 14).value = {
      text: "View Evidence",
      hyperlink: "https://drive.google.com/file/d/old/view",
    };
  }
  applyScenarioResults(opened.workbook, kbScenario, [
    capturedExecution(kbScenario, 0),
  ]);
  assert.equal(kb.getCell(kbScenario.turns[0].rowNumber, 9).text, "New response turn 1");
  assert.equal(readEvidenceHyperlink(kb.getCell(kbScenario.turns[0].rowNumber, 14)), undefined);
  assert.equal(kb.getCell(kbScenario.turns[1].rowNumber, 9).text, "");
  assert.equal(kb.getCell(kbScenario.turns[1].rowNumber, 10).text, "");
  assert.equal(readEvidenceHyperlink(kb.getCell(kbScenario.turns[1].rowNumber, 14)), undefined);

  const negativeScenario = opened.parsed.scenarios.find(
    (scenario) => scenario.testCaseId === "PGN-NEG-018",
  )!;
  const negative = opened.workbook.getWorksheet(negativeScenario.sheetName)!;
  const negativeRow = negativeScenario.sourceRowNumber;
  negative.getCell(negativeRow, 8).value = "Old multi-turn response";
  negative.getCell(negativeRow, 9).value = "Old timing";
  negative.getCell(negativeRow, 14).value = {
    text: "View Evidence",
    hyperlink: "https://drive.google.com/file/d/old-negative/view",
  };
  const firstTurn = capturedExecution(
    negativeScenario,
    0,
    "https://drive.google.com/file/d/new-first/view",
  );
  applyScenarioResults(opened.workbook, negativeScenario, [firstTurn]);
  assert.equal(negative.getCell(negativeRow, 8).text, "");
  assert.equal(negative.getCell(negativeRow, 9).text, "");
  assert.equal(readEvidenceHyperlink(negative.getCell(negativeRow, 14)), undefined);

  applyScenarioResults(opened.workbook, negativeScenario, [
    firstTurn,
    capturedExecution(negativeScenario, 1),
  ]);
  assert.match(negative.getCell(negativeRow, 8).text, /New response turn 2/);
  assert.equal(readEvidenceHyperlink(negative.getCell(negativeRow, 14)), undefined);
});

test("workbook saves reject external changes and cooperating writers are locked", async (context) => {
  const temporaryDirectory = await mkdtemp(
    path.join(tmpdir(), "pgn-workbook-conflict-"),
  );
  context.after(() => rm(temporaryDirectory, { recursive: true, force: true }));
  const outputPath = path.join(temporaryDirectory, "nested", "executed.xlsx");
  const opened = await openExecutedPgnWorkbook(sourcePath, outputPath);
  const release = await acquireWorkbookLock(outputPath, "first test writer");
  await assert.rejects(
    acquireWorkbookLock(outputPath, "second test writer"),
    /locked by first test writer/,
  );
  await release();
  const releaseAgain = await acquireWorkbookLock(outputPath, "replacement writer");
  await releaseAgain();

  const firstReleaseAfterReplacement = await acquireWorkbookLock(
    outputPath,
    "removed owner",
  );
  const lockPath = `${outputPath}.lock`;
  await rm(lockPath, { recursive: true, force: true });
  const replacementRelease = await acquireWorkbookLock(
    outputPath,
    "replacement owner",
  );
  await firstReleaseAfterReplacement();
  await assert.rejects(
    acquireWorkbookLock(outputPath, "third writer"),
    /locked by replacement owner/,
  );
  await replacementRelease();

  await mkdir(lockPath);
  await writeFile(
    path.join(lockPath, "owner-stale.json"),
    JSON.stringify({
      pid: 99_999_999,
      purpose: "stale writer",
      createdAt: new Date(0).toISOString(),
      token: "stale",
    }),
  );
  const recoveredRelease = await acquireWorkbookLock(
    outputPath,
    "recovered writer",
  );
  await recoveredRelease();

  await mkdir(lockPath);
  await utimes(lockPath, new Date(0), new Date(0));
  const emptyLockRecovery = await acquireWorkbookLock(
    outputPath,
    "empty-lock recovery",
  );
  await emptyLockRecovery();

  await mkdir(lockPath);
  const partialOwnerPath = path.join(lockPath, "owner-partial.json");
  await writeFile(partialOwnerPath, "{");
  await utimes(partialOwnerPath, new Date(0), new Date(0));
  const partialLockRecovery = await acquireWorkbookLock(
    outputPath,
    "partial-lock recovery",
  );
  await partialLockRecovery();

  await writeFile(
    lockPath,
    JSON.stringify({
      pid: 99_999_999,
      purpose: "legacy stale writer",
      createdAt: new Date(0).toISOString(),
    }),
  );
  await assert.rejects(
    acquireWorkbookLock(outputPath, "legacy replacement"),
    /stale legacy lock; remove it/,
  );
  await rm(lockPath, { force: true });

  const external = new ExcelJS.Workbook();
  await external.xlsx.readFile(outputPath);
  external.getWorksheet("Test Case Knowledge Base")!.getCell(4, 14).value =
    "External edit";
  await external.xlsx.writeFile(outputPath);
  await assert.rejects(
    saveExecutedPgnWorkbook(opened.workbook, outputPath),
    /changed outside this process/,
  );
  const reloaded = new ExcelJS.Workbook();
  await reloaded.xlsx.readFile(outputPath);
  assert.equal(
    reloaded.getWorksheet("Test Case Knowledge Base")!.getCell(4, 14).text,
    "External edit",
  );
});
