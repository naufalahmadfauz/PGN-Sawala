import {
  constants as fsConstants,
  access,
  copyFile,
  mkdir,
  rename,
  rm,
} from "node:fs/promises";
import path from "node:path";
import type { AppConfig } from "../config";
import { EXECUTION_METADATA_SHEET_NAME } from "./evidence-workbook";
import {
  RETEST_HISTORY_SHEET_NAME,
  RETEST_METADATA_SHEET_NAME,
} from "./retest-workbook";
import { TRANSCRIPT_SHEET_NAME } from "./pgn-types";
import {
  openExecutedPgnWorkbook,
  saveExecutedPgnWorkbook,
} from "./pgn-workbook-writer";

export interface FreshWorkbookResult {
  workbookPath: string;
  archivePath?: string;
  scenariosCleared: number;
}

async function exists(filePath: string): Promise<boolean> {
  return access(filePath)
    .then(() => true)
    .catch(() => false);
}

async function nextArchivePath(config: AppConfig, now: Date): Promise<string> {
  await mkdir(config.reportArchiveDir, { recursive: true });
  const extension = path.extname(config.pgnExecutedWorkbookPath);
  const baseName = path.basename(config.pgnExecutedWorkbookPath, extension);
  const timestamp = now.toISOString().replace(/[:.]/g, "-");
  let counter = 0;
  let archivePath: string;
  do {
    archivePath = path.join(
      config.reportArchiveDir,
      `${baseName}-before-fresh-${timestamp}${counter ? `-${counter}` : ""}${extension}`,
    );
    counter += 1;
  } while (await exists(archivePath));
  return archivePath;
}

export async function createFreshPgnWorkbook(
  config: AppConfig,
  now = new Date(),
): Promise<FreshWorkbookResult> {
  const temporaryPath = `${config.pgnExecutedWorkbookPath}.${process.pid}.fresh.xlsx`;
  await rm(temporaryPath, { force: true });
  try {
    const fresh = await openExecutedPgnWorkbook(
      config.pgnSourceWorkbookPath,
      temporaryPath,
    );
    for (const scenario of fresh.parsed.scenarios) {
      const worksheet = fresh.workbook.getWorksheet(scenario.sheetName)!;
      if (scenario.sheetKind === "kb") {
        for (const turn of scenario.turns) {
          for (const column of [9, 10, 11, 14]) {
            worksheet.getCell(turn.rowNumber, column).value = null;
          }
        }
      } else {
        for (const column of [8, 9, 10, 14]) {
          worksheet.getCell(scenario.sourceRowNumber, column).value = null;
        }
      }
    }
    for (const sheetName of [
      TRANSCRIPT_SHEET_NAME,
      EXECUTION_METADATA_SHEET_NAME,
      RETEST_HISTORY_SHEET_NAME,
      RETEST_METADATA_SHEET_NAME,
    ]) {
      const worksheet = fresh.workbook.getWorksheet(sheetName);
      if (worksheet) {
        fresh.workbook.removeWorksheet(worksheet.id);
      }
    }
    await saveExecutedPgnWorkbook(fresh.workbook, temporaryPath);

    let archivePath: string | undefined;
    if (await exists(config.pgnExecutedWorkbookPath)) {
      archivePath = await nextArchivePath(config, now);
      await copyFile(
        config.pgnExecutedWorkbookPath,
        archivePath,
        fsConstants.COPYFILE_EXCL,
      );
    }
    await rename(temporaryPath, config.pgnExecutedWorkbookPath);
    return {
      workbookPath: config.pgnExecutedWorkbookPath,
      archivePath,
      scenariosCleared: fresh.parsed.scenarios.length,
    };
  } catch (error) {
    await rm(temporaryPath, { force: true }).catch(() => undefined);
    throw error;
  }
}
