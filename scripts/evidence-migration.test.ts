import assert from "node:assert/strict";
import { mkdtemp, mkdir, readFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import test from "node:test";
import ExcelJS from "exceljs";
import JSZip from "jszip";
import sharp from "sharp";
import {
  loadConfig,
  normalizeGoogleDriveFolderId,
  type AppConfig,
} from "../src/config";
import {
  MAIN_EVIDENCE_COLUMN,
  TRANSCRIPT_EVIDENCE_STATUS_COLUMN,
  TRANSCRIPT_EVIDENCE_URL_COLUMN,
  getEvidenceRunMetadata,
  readEvidenceHyperlink,
} from "../src/excel/evidence-workbook";
import type {
  ExecutedTurn,
  PgnTestScenario,
} from "../src/excel/pgn-types";
import {
  applyScenarioExecution,
  openExecutedPgnWorkbook,
  saveExecutedPgnWorkbook,
} from "../src/excel/pgn-workbook-writer";
import {
  evidenceFileName,
  runEvidenceMigration,
} from "../src/evidence/evidence-migration";
import type {
  DriveEvidenceItem,
  EvidenceDrivePublisher,
} from "../src/evidence/google-drive";

const sourcePath = path.resolve(
  "data/PGN AI Assistant - Knowledge Base Testing Report - User Inputs.xlsx",
);
const runId = "20260901T120000000Z";

class MockDrivePublisher implements EvidenceDrivePublisher {
  readonly parentFolderId = "parent-folder-12345";
  readonly folderPrefix = "PGN-WhatsApp-Evidence";
  readonly serviceAccountEmail = "test@example.iam.gserviceaccount.com";
  readonly uploadedNames: string[] = [];
  readonly files = new Map<string, DriveEvidenceItem>();
  folderCreations = 0;
  private folder?: DriveEvidenceItem;

  async validateParentFolder(): Promise<DriveEvidenceItem> {
    return {
      id: this.parentFolderId,
      name: "Parent",
      webViewLink: "https://drive.google.com/drive/folders/parent-folder-12345",
      reused: true,
    };
  }

  async ensureRunFolder(
    requestedRunId: string,
    existingFolderId?: string,
  ): Promise<DriveEvidenceItem> {
    if (this.folder) {
      return { ...this.folder, reused: true };
    }
    this.folderCreations += 1;
    this.folder = {
      id: existingFolderId ?? `folder-${requestedRunId}`,
      name: `${this.folderPrefix}-${requestedRunId}`,
      webViewLink: `https://drive.google.com/drive/folders/folder-${requestedRunId}`,
      reused: Boolean(existingFolderId),
    };
    return this.folder;
  }

  async uploadPng(options: {
    folderId: string;
    localPath: string;
    fileName: string;
    existingFileId?: string;
  }): Promise<DriveEvidenceItem> {
    assert.equal(options.folderId, this.folder?.id);
    await readFile(options.localPath);
    if (options.fileName.startsWith("PGN-KB-075")) {
      throw new Error("Simulated Drive failure");
    }
    const existing = this.files.get(options.fileName);
    if (existing) {
      return { ...existing, reused: true };
    }
    this.uploadedNames.push(options.fileName);
    const created = {
      id: `file-${options.fileName}`,
      name: options.fileName,
      webViewLink: `https://drive.google.com/file/d/${options.fileName}/view`,
      reused: false,
    };
    this.files.set(options.fileName, created);
    return created;
  }
}

function makeConfig(temporaryDirectory: string): AppConfig {
  const base = loadConfig();
  const artifactsDir = path.join(temporaryDirectory, "artifacts");
  const evidenceDir = path.join(artifactsDir, "evidence");
  return {
    ...base,
    projectRoot: temporaryDirectory,
    artifactsDir,
    evidenceDir,
    evidenceCleanDir: path.join(evidenceDir, "clean"),
    reportArchiveDir: path.join(temporaryDirectory, "reports", "archive"),
    pgnSourceWorkbookPath: sourcePath,
    pgnExecutedWorkbookPath: path.join(
      temporaryDirectory,
      "reports",
      "executed.xlsx",
    ),
    googleDriveEvidenceEnabled: true,
    googleDriveEvidenceParentFolderId: "parent-folder-12345",
    googleDriveEvidenceFolderPrefix: "PGN-WhatsApp-Evidence",
    googleServiceAccountJson: undefined,
    legacyEvidenceCropLeft: undefined,
  };
}

async function createLegacyScreenshot(filePath: string): Promise<void> {
  await mkdir(path.dirname(filePath), { recursive: true });
  await sharp({
    create: {
      width: 1440,
      height: 1000,
      channels: 3,
      background: "#ffffff",
    },
  })
    .composite([
      {
        input: {
          create: {
            width: 1,
            height: 1000,
            channels: 3,
            background: "#dcd8d3",
          },
        },
        left: 496,
        top: 0,
      },
      {
        input: {
          create: {
            width: 943,
            height: 1000,
            channels: 3,
            background: "#f5f1eb",
          },
        },
        left: 497,
        top: 0,
      },
      {
        input: {
          create: {
            width: 600,
            height: 180,
            channels: 3,
            background: "#d9fdd3",
          },
        },
        left: 760,
        top: 250,
      },
      {
        input: {
          create: {
            width: 620,
            height: 240,
            channels: 3,
            background: "#ffffff",
          },
        },
        left: 540,
        top: 500,
      },
    ])
    .png()
    .toFile(filePath);
}

function execution(
  scenario: PgnTestScenario,
  turnIndex: number,
  evidencePath: string,
): ExecutedTurn {
  const turn = scenario.turns[turnIndex];
  return {
    turn,
    technicalStatus: "CAPTURED",
    sentAt: new Date("2026-09-01T12:00:00Z"),
    completedAt: new Date("2026-09-01T12:00:01Z"),
    botMessages: [
      {
        sequence: 1,
        message: `Response ${scenario.testCaseId} turn ${turn.turnNumber}`,
        timestamp: new Date("2026-09-01T12:00:01Z"),
      },
    ],
    combinedResponse: `Response ${scenario.testCaseId} turn ${turn.turnNumber}`,
    firstResponseMs: 500,
    totalResponseMs: 1_000,
    evidencePath,
    evidenceStatus: "EVIDENCE_LOCAL_ONLY",
  };
}

async function tableXml(filePath: string): Promise<Buffer> {
  const archive = await JSZip.loadAsync(await readFile(filePath));
  const table = archive.file("xl/tables/table1.xml");
  assert(table);
  return table.async("nodebuffer");
}

function resultSnapshot(workbook: ExcelJS.Workbook): string {
  const kb = workbook.getWorksheet("Test Case Knowledge Base")!;
  const negative = workbook.getWorksheet("Negative Case")!;
  return JSON.stringify([
    [4, 9, 10, 11],
    [5, 9, 10, 11],
    [32, 9, 10, 11],
    [33, 9, 10, 11],
    [77, 9, 10, 11],
  ].map(([row, ...columns]) => columns.map((column) => kb.getCell(row, column).text)).concat([
    [8, 9, 10].map((column) => negative.getCell(19, column).text),
  ]));
}

test("Drive folder configuration accepts IDs and folder URLs", () => {
  assert.equal(
    normalizeGoogleDriveFolderId("1AbCdEf_12345"),
    "1AbCdEf_12345",
  );
  assert.equal(
    normalizeGoogleDriveFolderId(
      "https://drive.google.com/drive/folders/1AbCdEf_12345?usp=sharing",
    ),
    "1AbCdEf_12345",
  );
  assert.throws(
    () => normalizeGoogleDriveFolderId("https://example.com/folders/1AbCdEf_12345"),
    /drive\.google\.com/,
  );
});

test("migration is progressive, idempotent, and keeps result cells and table XML intact", async (context) => {
  const temporaryDirectory = await mkdtemp(
    path.join(tmpdir(), "pgn-evidence-migration-"),
  );
  context.after(() => rm(temporaryDirectory, { recursive: true, force: true }));
  const config = makeConfig(temporaryDirectory);
  const opened = await openExecutedPgnWorkbook(
    config.pgnSourceWorkbookPath,
    config.pgnExecutedWorkbookPath,
  );
  const scenarioIds = [
    "PGN-KB-003",
    "PGN-KB-004",
    "PGN-KB-031",
    "PGN-KB-075",
    "PGN-NEG-018",
  ];
  const scenarios = new Map(
    opened.parsed.scenarios
      .filter((scenario) => scenarioIds.includes(scenario.testCaseId))
      .map((scenario) => [scenario.testCaseId, scenario]),
  );
  assert.equal(scenarios.size, scenarioIds.length);

  const originalHashes = new Map<string, Buffer>();
  for (const scenarioId of scenarioIds) {
    const scenario = scenarios.get(scenarioId)!;
    const executions: ExecutedTurn[] = [];
    for (let index = 0; index < scenario.turns.length; index += 1) {
      const turn = scenario.turns[index];
      const relativeEvidencePath = `artifacts/evidence/${runId}-${scenarioId}-turn-${turn.turnNumber}.png`;
      const absoluteEvidencePath = path.join(
        temporaryDirectory,
        relativeEvidencePath,
      );
      if (scenarioId !== "PGN-KB-004") {
        await createLegacyScreenshot(absoluteEvidencePath);
        originalHashes.set(absoluteEvidencePath, await readFile(absoluteEvidencePath));
      }
      executions.push(execution(scenario, index, relativeEvidencePath));
    }
    applyScenarioExecution(opened.workbook, runId, scenario, executions);
  }
  await saveExecutedPgnWorkbook(
    opened.workbook,
    config.pgnExecutedWorkbookPath,
  );
  const beforeWorkbook = new ExcelJS.Workbook();
  await beforeWorkbook.xlsx.readFile(config.pgnExecutedWorkbookPath);
  const beforeResults = resultSnapshot(beforeWorkbook);
  const sourceTable = await tableXml(sourcePath);

  const publisher = new MockDrivePublisher();
  const summary = await runEvidenceMigration({
    config,
    publisher,
    now: () => new Date("2026-09-01T13:00:00Z"),
    log: () => undefined,
  });
  assert(summary.totalEvidenceDiscovered >= 7);
  assert(summary.successfullyUploaded >= 5);
  assert.equal(summary.uploadErrors, 1);
  assert(summary.uploadErrorIds.includes("PGN-KB-075 turn 1"));
  assert(summary.missingIds.includes("PGN-KB-004 turn 1"));
  assert.equal(publisher.folderCreations, 1);

  for (const [filePath, original] of originalHashes) {
    assert((await readFile(filePath)).equals(original));
  }
  const cleaned = await sharp(
    path.join(
      config.evidenceCleanDir,
      runId,
      evidenceFileName("PGN-KB-003", 1),
    ),
  ).metadata();
  assert.equal(cleaned.width, 943);
  assert.equal(cleaned.height, 1000);

  const migrated = new ExcelJS.Workbook();
  await migrated.xlsx.readFile(config.pgnExecutedWorkbookPath);
  assert.equal(resultSnapshot(migrated), beforeResults);
  assert((await tableXml(config.pgnExecutedWorkbookPath)).equals(sourceTable));
  const kb = migrated.getWorksheet("Test Case Knowledge Base")!;
  assert.equal(
    readEvidenceHyperlink(kb.getCell(4, MAIN_EVIDENCE_COLUMN)),
    "https://drive.google.com/file/d/PGN-KB-003-turn-1.png/view",
  );
  assert(readEvidenceHyperlink(kb.getCell(32, MAIN_EVIDENCE_COLUMN)));
  assert(readEvidenceHyperlink(kb.getCell(33, MAIN_EVIDENCE_COLUMN)));
  const negative = migrated.getWorksheet("Negative Case")!;
  assert.equal(
    readEvidenceHyperlink(negative.getCell(19, MAIN_EVIDENCE_COLUMN)),
    "https://drive.google.com/file/d/PGN-NEG-018-turn-2.png/view",
  );
  const transcript = migrated.getWorksheet("Execution Transcript")!;
  const kb075Rows = transcript
    .getRows(2, transcript.rowCount - 1)!
    .filter(
      (row) =>
        row.getCell(2).text === "PGN-KB-075" && row.getCell(5).text === "1",
    );
  assert(
    kb075Rows.every(
      (row) =>
        row.getCell(TRANSCRIPT_EVIDENCE_STATUS_COLUMN).text ===
        "EVIDENCE_UPLOAD_ERROR",
    ),
  );
  assert(
    readEvidenceHyperlink(
      transcript
        .getRows(2, transcript.rowCount - 1)!
        .find(
          (row) =>
            row.getCell(2).text === "PGN-KB-003" &&
            row.getCell(5).text === "1",
        )!
        .getCell(TRANSCRIPT_EVIDENCE_URL_COLUMN),
    ),
  );
  assert(getEvidenceRunMetadata(migrated, runId));
  assert(await readFile(summary.backupPath));

  const uploadCount = publisher.uploadedNames.length;
  const second = await runEvidenceMigration({
    config,
    publisher,
    now: () => new Date("2026-09-01T14:00:00Z"),
    log: () => undefined,
  });
  assert.equal(publisher.folderCreations, 1);
  assert.equal(publisher.uploadedNames.length, uploadCount);
  assert(second.alreadySynced >= 5);
  assert.equal(second.uploadErrors, 1);
});
