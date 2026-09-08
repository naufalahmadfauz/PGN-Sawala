import path from "node:path";
import { loadConfig, type AppConfig } from "../src/config";
import { parseCliOptions } from "../src/pgn-cli";
import { isEntrypoint, runCliMain } from "../src/cli-entrypoint";
import { createFreshPgnWorkbook } from "../src/excel/fresh-workbook";
import { acquireWorkbookLock } from "../src/excel/workbook-lock";
import {
  acquireRunProcessLock,
  discoverRecoveryRun,
} from "../src/recovery/run-state";

export async function prepareFreshPgnWorkbook(
  config: AppConfig = loadConfig(),
): Promise<void> {
  const recovery = await discoverRecoveryRun(config.projectRoot);
  if (recovery.kind === "running") {
    throw new Error(
      `PGN run ${recovery.state.runId} is still active; fresh-run preparation is blocked`,
    );
  }
  if (recovery.kind === "recoverable") {
    throw new Error(
      `Recoverable PGN run ${recovery.state.runId} exists. Resume it or explicitly abandon it before preparing a fresh run.`,
    );
  }
  if (recovery.kind === "unreadable") {
    throw new Error(
      `Recovery state is unreadable${recovery.runId ? ` for ${recovery.runId}` : ""}; repair it before preparing a fresh run`,
    );
  }
  const runLock = await acquireRunProcessLock(
    config.projectRoot,
    "PGN fresh-run preparation",
  );
  let release: (() => Promise<void>) | undefined;
  try {
    release = await acquireWorkbookLock(
      config.pgnExecutedWorkbookPath,
      "PGN fresh-run preparation",
    );
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
    try {
      await release?.();
    } finally {
      await runLock.release();
    }
  }
}

if (isEntrypoint(import.meta.url)) {
  runCliMain(() => prepareFreshPgnWorkbook(loadConfig({ transport: parseCliOptions(process.argv.slice(2)).transport })));
}
