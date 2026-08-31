import { mkdir, rename, rm } from "node:fs/promises";
import path from "node:path";
import ExcelJS, { type Worksheet } from "exceljs";
import type { TestResult, TranscriptEntry } from "../types";

function styleWorksheet(worksheet: Worksheet): void {
  const header = worksheet.getRow(1);
  header.font = { bold: true, color: { argb: "FFFFFFFF" } };
  header.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FF1F4E78" },
  };
  header.alignment = { vertical: "middle", wrapText: true };
  header.height = 28;
  worksheet.views = [{ state: "frozen", ySplit: 1 }];
  worksheet.autoFilter = {
    from: { row: 1, column: 1 },
    to: { row: 1, column: worksheet.columnCount },
  };
  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber > 1) {
      row.alignment = { vertical: "top", wrapText: true };
    }
  });
}

export async function writeResultsWorkbook(
  filePath: string,
  results: TestResult[],
  transcript: TranscriptEntry[],
): Promise<void> {
  await mkdir(path.dirname(filePath), { recursive: true });
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "PGN WhatsApp QA Harness";
  workbook.created = new Date();

  const resultsSheet = workbook.addWorksheet("Results");
  resultsSheet.columns = [
    { header: "Run ID", key: "runId", width: 27 },
    { header: "Test ID", key: "testId", width: 16 },
    { header: "Category", key: "category", width: 22 },
    { header: "User Input", key: "userInput", width: 42 },
    {
      header: "Expected Behaviour",
      key: "expectedBehaviour",
      width: 42,
    },
    { header: "Bot Response", key: "botResponse", width: 60 },
    { header: "First Response (ms)", key: "firstResponseMs", width: 20 },
    { header: "Total Response (ms)", key: "totalResponseMs", width: 20 },
    { header: "Status", key: "status", width: 14 },
    { header: "Started At", key: "startedAt", width: 24 },
    { header: "Completed At", key: "completedAt", width: 24 },
    { header: "Error", key: "error", width: 45 },
    {
      header: "Screenshot/Evidence path",
      key: "evidencePath",
      width: 48,
    },
  ];
  for (const result of results) {
    const row = resultsSheet.addRow({
      runId: result.runId,
      testId: result.testCase.testId,
      category: result.testCase.category,
      userInput: result.testCase.userInput,
      expectedBehaviour: result.testCase.expectedBehaviour,
      botResponse: result.botResponse,
      firstResponseMs: result.firstResponseMs ?? null,
      totalResponseMs: result.totalResponseMs ?? null,
      status: result.status,
      startedAt: result.startedAt,
      completedAt: result.completedAt,
      error: result.error ?? "",
      evidencePath: result.evidencePath ?? "",
    });
    const statusCell = row.getCell("status");
    statusCell.font = { bold: true };
    statusCell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: {
        argb:
          result.status === "CAPTURED"
            ? "FFC6EFCE"
            : result.status === "TIMEOUT"
              ? "FFFFEB9C"
              : "FFFFC7CE",
      },
    };
  }
  resultsSheet.getColumn("startedAt").numFmt = "yyyy-mm-dd hh:mm:ss";
  resultsSheet.getColumn("completedAt").numFmt = "yyyy-mm-dd hh:mm:ss";
  styleWorksheet(resultsSheet);

  const transcriptSheet = workbook.addWorksheet("Transcript");
  transcriptSheet.columns = [
    { header: "Test ID", key: "testId", width: 16 },
    { header: "Sequence", key: "sequence", width: 12 },
    { header: "Role", key: "role", width: 10 },
    { header: "Message", key: "message", width: 75 },
    { header: "Timestamp", key: "timestamp", width: 24 },
  ];
  transcriptSheet.addRows(transcript);
  transcriptSheet.getColumn("timestamp").numFmt = "yyyy-mm-dd hh:mm:ss";
  styleWorksheet(transcriptSheet);

  const temporaryPath = `${filePath}.${process.pid}.tmp`;
  try {
    await workbook.xlsx.writeFile(temporaryPath);
    await rename(temporaryPath, filePath);
  } catch (error) {
    await rm(temporaryPath, { force: true }).catch(() => undefined);
    throw error;
  }
}
