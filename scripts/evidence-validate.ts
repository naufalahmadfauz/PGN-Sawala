import { access, readdir } from "node:fs/promises";
import path from "node:path";
import ExcelJS from "exceljs";
import { loadConfig } from "../src/config";
import {
  createRepresentativeEvidencePreviews,
  discoverEvidenceInventory,
} from "../src/evidence/evidence-migration";
import { createGoogleDriveEvidencePublisher } from "../src/evidence/google-drive";

async function exists(filePath: string): Promise<boolean> {
  return access(filePath)
    .then(() => true)
    .catch(() => false);
}

async function main(): Promise<void> {
  const config = loadConfig();
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(config.pgnExecutedWorkbookPath);
  const inventory = discoverEvidenceInventory(workbook);
  const localFiles = (await readdir(config.evidenceDir, { withFileTypes: true }))
    .filter((entry) => entry.isFile() && entry.name.toLowerCase().endsWith(".png"))
    .length;
  const mappedExisting = await Promise.all(
    inventory.records.map(async (record) =>
      record.localEvidencePath
        ? exists(path.resolve(config.projectRoot, record.localEvidencePath))
        : false,
    ),
  );
  const mappedLocal = mappedExisting.filter(Boolean).length;
  const missingMapped = inventory.records
    .filter((_, index) => !mappedExisting[index])
    .map((record) => `${record.testCaseId} turn ${record.turnNumber}`);
  const existingUrls = inventory.records.filter((record) => record.evidenceUrl).length;
  const preview = await createRepresentativeEvidencePreviews(config, inventory);

  console.log("Evidence validation");
  console.log("-------------------");
  console.log(`Local evidence count: ${localFiles}`);
  console.log(`Mapped transcript evidence: ${inventory.records.length}`);
  console.log(`Mapped local screenshots: ${mappedLocal}`);
  console.log(
    `Missing screenshots: ${new Set([...missingMapped, ...inventory.missingCompletedTurns]).size}`,
  );
  console.log(`Existing Drive URLs: ${existingUrls}`);
  console.log(`Pending uploads: ${inventory.records.length - existingUrls}`);
  console.log(
    `Validated legacy crop: left=${preview.crop.left}, ${preview.crop.originalWidth}x${preview.crop.originalHeight} -> ${preview.crop.width}x${preview.crop.height}`,
  );
  for (const item of preview.previews) {
    console.log(
      `Preview: ${item.testCaseId} turn ${item.turnNumber} -> ${path.relative(config.projectRoot, item.outputPath)}`,
    );
  }

  console.log("");
  console.log("Drive");
  console.log("-----");
  if (!config.googleDriveEvidenceEnabled) {
    console.log("Authentication status: DISABLED");
    console.log("Parent status: NOT CONFIGURED");
  } else {
    try {
      const publisher = createGoogleDriveEvidencePublisher(config);
      const parent = await publisher.validateParentFolder();
      console.log("Authentication status: READY");
      console.log(`Service Account: ${publisher.serviceAccountEmail}`);
      console.log(`Parent status: READY (${parent.name}, ${parent.id})`);
    } catch (error) {
      console.log("Authentication status: ERROR");
      console.log(
        `Parent status: ${error instanceof Error ? error.message : String(error)}`,
      );
      process.exitCode = 1;
    }
  }

  const missing = [...new Set([...missingMapped, ...inventory.missingCompletedTurns])];
  if (missing.length) {
    console.log("");
    console.log("Evidence missing");
    console.log("----------------");
    missing.forEach((item) => console.log(item));
  }
  console.log("");
  console.log("WhatsApp messages sent: 0");
  console.log("Testcases rerun: 0");
}

main().catch((error) => {
  console.error(error instanceof Error ? error.message : error);
  process.exitCode = 1;
});
