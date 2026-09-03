import { access, readdir } from "node:fs/promises";
import path from "node:path";
import ExcelJS from "exceljs";
import { loadConfig } from "../src/config";
import {
  createRepresentativeEvidencePreviews,
  discoverEvidenceInventory,
} from "../src/evidence/evidence-migration";
import { createGoogleDriveEvidencePublisher } from "../src/evidence/google-drive";
import { safeGoogleCredentialError } from "../src/evidence/google-service-account";

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
  const missingCompletedLabels = inventory.missingCompletedTurns.map(
    (item) => `${item.testCaseId} turn ${item.turnNumber}`,
  );
  console.log(
    `Missing screenshots: ${new Set([...missingMapped, ...missingCompletedLabels]).size}`,
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
  console.log("Google Drive Evidence");
  console.log("---------------------");
  console.log(`Enabled: ${config.googleDriveEvidenceEnabled ? "YES" : "NO"}`);
  console.log(
    `Environment file: ${config.environmentFileLoaded ? `${path.relative(config.projectRoot, config.environmentFilePath) || ".env"} (LOADED)` : "NOT FOUND"}`,
  );
  console.log(`Configuration source: ${config.googleDriveConfigurationSource}`);
  console.log(
    `Parent folder: ${config.googleDriveEvidenceParentFolderId ? "CONFIGURED" : "NOT CONFIGURED"}`,
  );
  console.log(
    `Credential source: ${config.googleServiceAccount?.source ?? "NOT CONFIGURED"}`,
  );
  if (config.googleServiceAccount?.source === "GOOGLE_SERVICE_ACCOUNT_FILE") {
    console.log(
      `Credential file: ${config.googleServiceAccount.fileDisplayPath}`,
    );
  }
  if ((config.googleServiceAccount?.configuredSources.length ?? 0) > 1) {
    console.log(
      `Credential warning: Multiple sources configured; using ${config.googleServiceAccount!.source}`,
    );
  }
  if (!config.googleDriveEvidenceEnabled) {
    console.log("Authentication: NOT CHECKED");
    console.log("Parent folder access: NOT CHECKED");
    console.log("Write access: NOT CHECKED");
    console.log("DISABLED");
  } else if (!config.googleDriveEvidenceParentFolderId) {
    console.log("Authentication: NOT CHECKED");
    console.log("Parent folder access: ERROR");
    console.log("Write access: NOT CHECKED");
    console.log("NOT READY");
    process.exitCode = 1;
  } else {
    try {
      const publisher = createGoogleDriveEvidencePublisher(config);
      console.log(`Service account: ${publisher.serviceAccountEmail}`);
      await publisher.validateParentFolder();
      console.log("Authentication: OK");
      console.log("Parent folder access: OK");
      console.log("Write access: OK");
      console.log("READY");
    } catch (error) {
      console.log("Authentication: ERROR");
      console.log("Parent folder access: ERROR");
      console.log("Write access: ERROR");
      console.log(
        `Reason: ${safeGoogleCredentialError(error, config.googleServiceAccount?.value)}`,
      );
      console.log("NOT READY");
      process.exitCode = 1;
    }
  }

  const missing = [
    ...new Set([
      ...missingMapped,
      ...missingCompletedLabels,
    ]),
  ];
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
  console.error(safeGoogleCredentialError(error));
  process.exitCode = 1;
});
