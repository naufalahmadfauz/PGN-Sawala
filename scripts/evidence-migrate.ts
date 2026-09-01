import path from "node:path";
import { loadConfig } from "../src/config";
import { runEvidenceMigration } from "../src/evidence/evidence-migration";
import { createGoogleDriveEvidencePublisher } from "../src/evidence/google-drive";

async function main(): Promise<void> {
  const config = loadConfig();
  const publisher = createGoogleDriveEvidencePublisher(config);
  const summary = await runEvidenceMigration({ config, publisher });

  console.log("");
  console.log("Migration");
  console.log("---------");
  console.log(`Total evidence discovered: ${summary.totalEvidenceDiscovered}`);
  console.log(`Successfully cleaned: ${summary.successfullyCleaned}`);
  console.log(`Successfully uploaded: ${summary.successfullyUploaded}`);
  console.log(`Already synced: ${summary.alreadySynced}`);
  console.log(`Missing: ${summary.missing}`);
  console.log(`Requires rerun: ${summary.requiresRerun}`);
  console.log(`Upload errors: ${summary.uploadErrors}`);

  console.log("");
  console.log("Drive");
  console.log("-----");
  console.log(`Configured parent: ${summary.driveParentId}`);
  console.log(
    `${summary.driveFolderReused ? "Reused" : "Generated"} run folder: ${summary.driveFolderId}`,
  );
  console.log(`Run folder URL: ${summary.driveFolderUrl}`);

  console.log("");
  console.log("Excel");
  console.log("-----");
  console.log(
    `Workbook updated: ${path.relative(config.projectRoot, config.pgnExecutedWorkbookPath)}`,
  );
  console.log("Evidence column used: N");
  console.log("Transcript Evidence URL column: N");
  console.log(
    `Backup path: ${path.relative(config.projectRoot, summary.backupPath)}`,
  );

  if (summary.missingIds.length) {
    console.log("");
    console.log("Evidence missing:");
    summary.missingIds.forEach((item) => console.log(item));
  }
  if (summary.requiresRerunIds.length) {
    console.log("");
    console.log("Evidence requires rerun:");
    summary.requiresRerunIds.forEach((item) => console.log(item));
  }
  if (summary.uploadErrorIds.length) {
    console.log("");
    console.log("Evidence upload errors:");
    summary.uploadErrorIds.forEach((item) => console.log(item));
  }

  console.log("");
  console.log("Safety");
  console.log("------");
  console.log("WhatsApp messages sent during migration: 0");
  console.log("Testcases rerun during migration: 0");
}

main().catch((error) => {
  console.error(error instanceof Error ? error.message : error);
  process.exitCode = 1;
});
