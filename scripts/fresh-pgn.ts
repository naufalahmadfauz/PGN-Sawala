import path from "node:path";
import { loadConfig } from "../src/config";
import { createFreshPgnWorkbook } from "../src/excel/fresh-workbook";
import { acquireWorkbookLock } from "../src/excel/workbook-lock";

async function main(): Promise<void> {
  const config = loadConfig();
  const release = await acquireWorkbookLock(
    config.pgnExecutedWorkbookPath,
    "PGN fresh-run preparation",
  );
  try {
    const result = await createFreshPgnWorkbook(config);
    console.log("PGN fresh run prepared");
    console.log(
      `Workbook: ${path.relative(config.projectRoot, result.workbookPath)}`,
    );
    console.log(`Scenarios cleared: ${result.scenariosCleared}`);
    console.log(
      `Archived previous run: ${result.archivePath ? path.relative(config.projectRoot, result.archivePath) : "none"}`,
    );
    console.log("WhatsApp was not opened. Run validation and execution separately.");
  } finally {
    await release();
  }
}

main().catch((error) => {
  console.error(error instanceof Error ? error.message : error);
  process.exitCode = 1;
});
