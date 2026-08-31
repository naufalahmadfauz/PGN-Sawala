import assert from "node:assert/strict";
import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import test from "node:test";
import ExcelJS from "exceljs";
import JSZip from "jszip";
import {
  appendPostResetDrainTranscript,
  appendSessionResetTranscript,
  openExecutedPgnWorkbook,
  saveExecutedPgnWorkbook,
} from "../src/excel/pgn-workbook-writer";
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
