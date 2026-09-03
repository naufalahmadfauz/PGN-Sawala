import path from "node:path";
import { loadConfig, type AppConfig } from "../src/config";
import { isEntrypoint, runCliMain } from "../src/cli-entrypoint";
import { createFreshPgnWorkbook } from "../src/excel/fresh-workbook";
import { acquireWorkbookLock } from "../src/excel/workbook-lock";

export async function prepareFreshPgnWorkbook(
  config: AppConfig = loadConfig(),
): Promise<void> {
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

if (isEntrypoint(import.meta.url)) {
  runCliMain(() => prepareFreshPgnWorkbook());
}
