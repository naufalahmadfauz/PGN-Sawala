import assert from "node:assert/strict";
import {
  access,
  copyFile,
  mkdtemp,
  mkdir,
  readFile,
  rm,
  writeFile,
} from "node:fs/promises";
import { createHash } from "node:crypto";
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
  EVIDENCE_MIGRATION_VERSION,
  TRANSCRIPT_EVIDENCE_STATUS_COLUMN,
  TRANSCRIPT_EVIDENCE_URL_COLUMN,
  getEvidenceFileMetadata,
  getEvidenceRunMetadata,
  readEvidenceHyperlink,
  upsertEvidenceRunMetadata,
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
  RETEST_HISTORY_SHEET_NAME,
  ensureRetestWorkbookSchema,
  snapshotRetestHistory,
} from "../src/excel/retest-workbook";
import {
  evidenceFileName,
  runEvidenceMigration,
} from "../src/evidence/evidence-migration";
import type {
  DriveEvidenceItem,
  EvidenceDrivePublisher,
} from "../src/evidence/google-drive";
import { GoogleDriveEvidencePublisher } from "../src/evidence/google-drive";

const sourcePath = path.resolve(
  "data/PGN AI Assistant - Knowledge Base Testing Report - User Inputs.xlsx",
);
const runId = "20260901T120000000Z";

class MockDrivePublisher implements EvidenceDrivePublisher {
  readonly parentFolderId = "parent-folder-12345";
  readonly folderPrefix = "PGN-WhatsApp-Evidence";
  readonly retestFolderPrefix = "PGN-WhatsApp-Retest";
  readonly serviceAccountEmail = "test@example.iam.gserviceaccount.com";
  readonly uploadedNames: string[] = [];
  readonly uploads: Array<{ folderId: string; fileName: string }> = [];
  readonly ensuredFolderNames = new Map<string, string>();
  readonly failedNamePrefixes = new Set(["PGN-KB-075"]);
  readonly files = new Map<string, DriveEvidenceItem>();
  folderCreations = 0;
  private readonly folders = new Map<string, DriveEvidenceItem>();

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
    expectedFolderName?: string,
  ): Promise<DriveEvidenceItem> {
    const existingFolder = this.folders.get(requestedRunId);
    if (existingFolder) {
      this.ensuredFolderNames.set(requestedRunId, existingFolder.name);
      return { ...existingFolder, reused: true };
    }
    this.folderCreations += 1;
    const folder = {
      id: existingFolderId || `folder-${requestedRunId}`,
      name: expectedFolderName ?? `${this.folderPrefix}-${requestedRunId}`,
      webViewLink: `https://drive.google.com/drive/folders/folder-${requestedRunId}`,
      reused: Boolean(existingFolderId),
    };
    this.folders.set(requestedRunId, folder);
    this.ensuredFolderNames.set(requestedRunId, folder.name);
    return folder;
  }

  async uploadPng(options: {
    folderId: string;
    localPath: string;
    fileName: string;
    existingFileId?: string;
  }): Promise<DriveEvidenceItem> {
    assert(
      [...this.folders.values()].some((folder) => folder.id === options.folderId),
    );
    await readFile(options.localPath);
    if (
      [...this.failedNamePrefixes].some((prefix) =>
        options.fileName.startsWith(prefix),
      )
    ) {
      throw new Error("Simulated Drive failure");
    }
    const fileKey = `${options.folderId}|${options.fileName}`;
    const existing = this.files.get(fileKey);
    if (existing) {
      return { ...existing, reused: true };
    }
    this.uploadedNames.push(options.fileName);
    this.uploads.push({
      folderId: options.folderId,
      fileName: options.fileName,
    });
    const created = {
      id: `file-${options.fileName}`,
      name: options.fileName,
      webViewLink: `https://drive.google.com/file/d/${options.fileName}/view`,
      reused: false,
    };
    this.files.set(fileKey, created);
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
  const safeName = evidenceFileName("../../PGN/unsafe", 1);
  assert(!safeName.includes("/") && !safeName.includes(".."));
});

test("Drive publisher rejects stale folder metadata and refreshes changed bytes", async (context) => {
  const temporaryDirectory = await mkdtemp(
    path.join(tmpdir(), "pgn-drive-publisher-"),
  );
  context.after(() => rm(temporaryDirectory, { recursive: true, force: true }));
  const localPath = path.join(temporaryDirectory, "evidence.png");
  await writeFile(localPath, "first image bytes");
  const localMd5 = createHash("md5")
    .update(await readFile(localPath))
    .digest("hex");
  let remoteMd5 = "stale-md5";
  let updateCalls = 0;
  let listCalls = 0;
  const publisher = new GoogleDriveEvidencePublisher({
    parentFolderId: "parent-folder-12345",
    folderPrefix: "PGN-WhatsApp-Evidence",
    retestFolderPrefix: "PGN-WhatsApp-Retest",
    credentials: {
      type: "service_account",
      client_email: "test@example.iam.gserviceaccount.com",
      private_key: "not-a-real-private-key",
    },
  });
  Reflect.set(publisher, "drive", {
    files: {
      get: async (options: { fileId: string }) => {
        if (options.fileId === "wrong-folder") {
          return {
            data: {
              id: "wrong-folder",
              name: "Different-Run",
              mimeType: "application/vnd.google-apps.folder",
              parents: ["parent-folder-12345"],
              webViewLink: "https://drive.google.com/drive/folders/wrong-folder",
            },
          };
        }
        return {
          data: {
            id: "existing-file",
            name: "PGN-KB-003-turn-1.png",
            mimeType: "image/png",
            parents: ["expected-folder"],
            md5Checksum: remoteMd5,
            webViewLink: "https://drive.google.com/file/d/existing-file/view",
          },
        };
      },
      list: async () => {
        listCalls += 1;
        return {
          data: {
            files: [
              {
                id: "expected-folder",
                name: "PGN-WhatsApp-Evidence-run-a",
                mimeType: "application/vnd.google-apps.folder",
                parents: ["parent-folder-12345"],
                webViewLink:
                  "https://drive.google.com/drive/folders/expected-folder",
              },
            ],
          },
        };
      },
      update: async (options: {
        media: { body: NodeJS.ReadableStream };
      }) => {
        for await (const _chunk of options.media.body) {
          // Consume the mocked upload stream.
        }
        updateCalls += 1;
        remoteMd5 = localMd5;
        return {
          data: {
            id: "existing-file",
            name: "PGN-KB-003-turn-1.png",
            webViewLink: "https://drive.google.com/file/d/existing-file/view",
          },
        };
      },
      create: async () => {
        throw new Error("A new Drive item should not be created");
      },
    },
  });

  const folder = await publisher.ensureRunFolder("run-a", "wrong-folder");
  assert.equal(folder.id, "expected-folder");
  assert.equal(listCalls, 1);
  const updated = await publisher.uploadPng({
    folderId: folder.id,
    localPath,
    fileName: "PGN-KB-003-turn-1.png",
    existingFileId: "existing-file",
  });
  assert.equal(updated.reused, false);
  assert.equal(updateCalls, 1);
  const reused = await publisher.uploadPng({
    folderId: folder.id,
    localPath,
    fileName: "PGN-KB-003-turn-1.png",
    existingFileId: "existing-file",
  });
  assert.equal(reused.reused, true);
  assert.equal(updateCalls, 1);
});

test("invalid Drive parent leaves the completed workbook untouched", async (context) => {
  const temporaryDirectory = await mkdtemp(
    path.join(tmpdir(), "pgn-evidence-invalid-parent-"),
  );
  context.after(() => rm(temporaryDirectory, { recursive: true, force: true }));
  const config = makeConfig(temporaryDirectory);
  await mkdir(path.dirname(config.pgnExecutedWorkbookPath), { recursive: true });
  await copyFile(sourcePath, config.pgnExecutedWorkbookPath);
  const original = await readFile(config.pgnExecutedWorkbookPath);
  const publisher = new MockDrivePublisher();
  publisher.validateParentFolder = async () => {
    throw new Error("Invalid Drive parent");
  };

  await assert.rejects(
    runEvidenceMigration({ config, publisher, log: () => undefined }),
    /Invalid Drive parent/,
  );
  assert((await readFile(config.pgnExecutedWorkbookPath)).equals(original));
  await assert.rejects(access(config.reportArchiveDir));
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
  const beforeMigrationFile = await readFile(config.pgnExecutedWorkbookPath);
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
  assert((await readFile(summary.backupPath)).equals(beforeMigrationFile));

  const uploadCount = publisher.uploadedNames.length;
  const beforeSecondMigration = await readFile(config.pgnExecutedWorkbookPath);
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
  assert.notEqual(second.backupPath, summary.backupPath);
  assert((await readFile(second.backupPath)).equals(beforeSecondMigration));
});

test("migration isolates run folders and lets the newest record clear stale evidence", async (context) => {
  const temporaryDirectory = await mkdtemp(
    path.join(tmpdir(), "pgn-evidence-multi-run-"),
  );
  context.after(() => rm(temporaryDirectory, { recursive: true, force: true }));
  const config = makeConfig(temporaryDirectory);
  const opened = await openExecutedPgnWorkbook(
    config.pgnSourceWorkbookPath,
    config.pgnExecutedWorkbookPath,
  );
  const scenarios = new Map(
    opened.parsed.scenarios.map((scenario) => [scenario.testCaseId, scenario]),
  );
  const firstRun = "20260901T120000000Z";
  const secondRun = "20260901T130000000Z";

  const addScenario = async (
    scenarioId: string,
    executionRunId: string,
    createScreenshots: boolean,
  ): Promise<void> => {
    const scenario = scenarios.get(scenarioId)!;
    const executions: ExecutedTurn[] = [];
    for (let index = 0; index < scenario.turns.length; index += 1) {
      const relativeEvidencePath = `artifacts/evidence/${executionRunId}-${scenarioId}-turn-${scenario.turns[index].turnNumber}.png`;
      if (createScreenshots) {
        await createLegacyScreenshot(
          path.join(temporaryDirectory, relativeEvidencePath),
        );
      }
      executions.push(
        execution(scenario, index, relativeEvidencePath),
      );
    }
    applyScenarioExecution(
      opened.workbook,
      executionRunId,
      scenario,
      executions,
    );
  };

  await addScenario("PGN-KB-003", firstRun, true);
  await addScenario("PGN-KB-031", firstRun, true);
  await addScenario("PGN-KB-003", secondRun, false);
  const kb004 = scenarios.get("PGN-KB-004")!;
  opened.workbook
    .getWorksheet(kb004.sheetName)!
    .getCell(kb004.sourceRowNumber, 9).value = "Completed without transcript";
  await saveExecutedPgnWorkbook(
    opened.workbook,
    config.pgnExecutedWorkbookPath,
  );

  const publisher = new MockDrivePublisher();
  const summary = await runEvidenceMigration({
    config,
    publisher,
    log: () => undefined,
  });
  assert.equal(summary.driveFolders.length, 2);
  assert.equal(publisher.folderCreations, 2);
  assert(
    publisher.uploads.some(
      (upload) =>
        upload.folderId === `folder-${firstRun}` &&
        upload.fileName.includes("PGN-KB-003"),
    ),
  );
  assert(
    !publisher.uploads.some(
      (upload) =>
        upload.folderId === `folder-${secondRun}` &&
        upload.fileName.includes("PGN-KB-003"),
    ),
  );

  const migrated = new ExcelJS.Workbook();
  await migrated.xlsx.readFile(config.pgnExecutedWorkbookPath);
  const kb = migrated.getWorksheet("Test Case Knowledge Base")!;
  assert.equal(kb.getCell(scenarios.get("PGN-KB-003")!.sourceRowNumber, 14).text, "");
  assert(
    getEvidenceFileMetadata(migrated, "MISSING|PGN-KB-004|1"),
  );
  const newestRows = migrated
    .getWorksheet("Execution Transcript")!
    .getRows(2, migrated.getWorksheet("Execution Transcript")!.rowCount - 1)!
    .filter(
      (row) =>
        row.getCell(1).text === secondRun &&
        row.getCell(2).text === "PGN-KB-003",
    );
  assert(
    newestRows.every(
      (row) =>
        !readEvidenceHyperlink(row.getCell(TRANSCRIPT_EVIDENCE_URL_COLUMN)) &&
        row.getCell(TRANSCRIPT_EVIDENCE_STATUS_COLUMN).text ===
          "EVIDENCE_MISSING",
    ),
  );

  const resumed = await openExecutedPgnWorkbook(
    config.pgnSourceWorkbookPath,
    config.pgnExecutedWorkbookPath,
  );
  const resolvedScenario = resumed.parsed.scenarios.find(
    (scenario) => scenario.testCaseId === "PGN-KB-004",
  )!;
  const resolvedPath = `artifacts/evidence/${secondRun}-PGN-KB-004-turn-1.png`;
  await createLegacyScreenshot(path.join(temporaryDirectory, resolvedPath));
  applyScenarioExecution(resumed.workbook, secondRun, resolvedScenario, [
    execution(resolvedScenario, 0, resolvedPath),
  ]);
  await saveExecutedPgnWorkbook(
    resumed.workbook,
    config.pgnExecutedWorkbookPath,
  );
  await runEvidenceMigration({ config, publisher, log: () => undefined });
  const reconciled = new ExcelJS.Workbook();
  await reconciled.xlsx.readFile(config.pgnExecutedWorkbookPath);
  assert.equal(
    getEvidenceFileMetadata(reconciled, "MISSING|PGN-KB-004|1"),
    undefined,
  );
});

test("migration retains future and retest conversation screenshots without legacy cropping", async (context) => {
  const temporaryDirectory = await mkdtemp(
    path.join(tmpdir(), "pgn-evidence-future-run-"),
  );
  context.after(() => rm(temporaryDirectory, { recursive: true, force: true }));
  const config = makeConfig(temporaryDirectory);
  const opened = await openExecutedPgnWorkbook(
    config.pgnSourceWorkbookPath,
    config.pgnExecutedWorkbookPath,
  );
  const scenario = opened.parsed.scenarios.find(
    (item) => item.testCaseId === "PGN-KB-003",
  )!;
  const futureRunId = "20260902T010000000Z";
  const relativeEvidencePath = `artifacts/evidence/${futureRunId}-PGN-KB-003-turn-1.png`;
  const absoluteEvidencePath = path.join(
    temporaryDirectory,
    relativeEvidencePath,
  );
  await mkdir(path.dirname(absoluteEvidencePath), { recursive: true });
  await sharp({
    create: {
      width: 700,
      height: 900,
      channels: 3,
      background: "#efeae2",
    },
  })
    .composite([
      {
        input: {
          create: {
            width: 420,
            height: 140,
            channels: 3,
            background: "#d9fdd3",
          },
        },
        left: 240,
        top: 180,
      },
      {
        input: {
          create: {
            width: 480,
            height: 220,
            channels: 3,
            background: "#ffffff",
          },
        },
        left: 30,
        top: 430,
      },
    ])
    .png()
    .toFile(absoluteEvidencePath);
  applyScenarioExecution(opened.workbook, futureRunId, scenario, [
    execution(scenario, 0, relativeEvidencePath),
  ]);
  const retestScenario = opened.parsed.scenarios.find(
    (item) => item.testCaseId === "PGN-KB-004",
  )!;
  const retestRunId = "RETEST-20260902T020000Z";
  const retestRelativePath = `artifacts/evidence/${retestRunId}-PGN-KB-004-turn-1.png`;
  await copyFile(
    absoluteEvidencePath,
    path.join(temporaryDirectory, retestRelativePath),
  );
  ensureRetestWorkbookSchema(opened.workbook);
  snapshotRetestHistory(
    opened.workbook,
    retestRunId,
    retestScenario,
    new Date("2026-09-02T02:00:00Z"),
  );
  applyScenarioExecution(opened.workbook, retestRunId, retestScenario, [
    execution(retestScenario, 0, retestRelativePath),
  ]);
  const originalTimestamp = new Date("2026-09-02T01:00:00Z");
  upsertEvidenceRunMetadata(opened.workbook, {
    runId: futureRunId,
    folderId: "",
    folderUrl: "",
    migrationVersion: EVIDENCE_MIGRATION_VERSION,
    timestamp: originalTimestamp,
    mode: "FUTURE",
  });
  upsertEvidenceRunMetadata(opened.workbook, {
    runId: retestRunId,
    folderId: "",
    folderUrl: "",
    migrationVersion: EVIDENCE_MIGRATION_VERSION,
    timestamp: new Date("2026-09-02T02:00:00Z"),
    mode: "RETEST",
  });
  await saveExecutedPgnWorkbook(
    opened.workbook,
    config.pgnExecutedWorkbookPath,
  );

  const publisher = new MockDrivePublisher();
  const summary = await runEvidenceMigration({
    config,
    publisher,
    log: () => undefined,
  });
  assert.equal(summary.successfullyCleaned, 0);
  assert.equal(summary.successfullyUploaded, 2);
  assert.equal(publisher.uploads[0].folderId, `folder-${futureRunId}`);
  assert.equal(publisher.uploads[0].fileName, "PGN-KB-003-turn-1.png");
  assert.equal(
    publisher.ensuredFolderNames.get(retestRunId),
    "PGN-WhatsApp-Retest-20260902T020000Z",
  );
  await assert.rejects(
    access(path.join(config.evidenceCleanDir, futureRunId)),
  );
  await assert.rejects(
    access(path.join(config.evidenceCleanDir, retestRunId)),
  );
  const migrated = new ExcelJS.Workbook();
  await migrated.xlsx.readFile(config.pgnExecutedWorkbookPath);
  const runMetadata = getEvidenceRunMetadata(migrated, futureRunId)!;
  assert.equal(runMetadata.mode, "FUTURE");
  assert.equal(runMetadata.timestamp.toISOString(), originalTimestamp.toISOString());
  assert.equal(getEvidenceRunMetadata(migrated, retestRunId)?.mode, "RETEST");
  assert(
    readEvidenceHyperlink(
      migrated
        .getWorksheet(scenario.sheetName)!
        .getCell(scenario.sourceRowNumber, MAIN_EVIDENCE_COLUMN),
    ),
  );
  const retestEvidenceUrl = readEvidenceHyperlink(
    migrated
      .getWorksheet(retestScenario.sheetName)!
      .getCell(retestScenario.sourceRowNumber, MAIN_EVIDENCE_COLUMN),
  );
  assert(retestEvidenceUrl);
  assert.equal(
    readEvidenceHyperlink(
      migrated.getWorksheet(RETEST_HISTORY_SHEET_NAME)!.getCell(2, 17),
    ),
    retestEvidenceUrl,
  );

  publisher.failedNamePrefixes.add("PGN-KB-004");
  const retry = await runEvidenceMigration({
    config,
    publisher,
    log: () => undefined,
  });
  assert.equal(retry.uploadErrors, 1);
  const afterFailedRetry = new ExcelJS.Workbook();
  await afterFailedRetry.xlsx.readFile(config.pgnExecutedWorkbookPath);
  assert.equal(
    readEvidenceHyperlink(
      afterFailedRetry
        .getWorksheet(retestScenario.sheetName)!
        .getCell(retestScenario.sourceRowNumber, MAIN_EVIDENCE_COLUMN),
    ),
    undefined,
  );
  assert.equal(
    readEvidenceHyperlink(
      afterFailedRetry
        .getWorksheet(RETEST_HISTORY_SHEET_NAME)!
        .getCell(2, 17),
    ),
    retestEvidenceUrl,
  );
});
