import { access } from "node:fs/promises";
import path from "node:path";
import { loadConfig, type AppConfig } from "../src/config";
import { isEntrypoint, runCliMain } from "../src/cli-entrypoint";
import {
  MAIN_EVIDENCE_COLUMN,
  readEvidenceHyperlink,
} from "../src/excel/evidence-workbook";
import { loadPgnWorkbook } from "../src/excel/pgn-workbook-loader";
import { assertPgnWorkbookValid } from "../src/excel/pgn-workbook-validator";
import { getRetestRunMetadata } from "../src/excel/retest-workbook";
import type { PgnTestScenario } from "../src/excel/pgn-types";
import { parseCliOptions } from "../src/pgn-cli";
import { createGoogleDriveEvidencePublisher } from "../src/evidence/google-drive";
import { safeGoogleCredentialError } from "../src/evidence/google-service-account";
import { needsFinalRetestCleanup } from "../src/retest/retest-run";
import { selectRetestScenarios } from "../src/retest/retest-selection";

async function exists(filePath: string): Promise<boolean> {
  return access(filePath)
    .then(() => true)
    .catch(() => false);
}

function hasExistingResponse(
  workbook: Awaited<ReturnType<typeof loadPgnWorkbook>>["workbook"],
  scenario: PgnTestScenario,
): boolean {
  const worksheet = workbook.getWorksheet(scenario.sheetName)!;
  return scenario.sheetKind === "kb"
    ? scenario.turns.some((turn) =>
        Boolean(worksheet.getCell(turn.rowNumber, 9).text.trim()),
      )
    : Boolean(worksheet.getCell(scenario.sourceRowNumber, 8).text.trim());
}

function hasExistingEvidence(
  workbook: Awaited<ReturnType<typeof loadPgnWorkbook>>["workbook"],
  scenario: PgnTestScenario,
): boolean {
  const worksheet = workbook.getWorksheet(scenario.sheetName)!;
  const rows =
    scenario.sheetKind === "kb"
      ? scenario.turns.map((turn) => turn.rowNumber)
      : [scenario.sourceRowNumber];
  return rows.some((rowNumber) =>
    Boolean(
      readEvidenceHyperlink(
        worksheet.getCell(rowNumber, MAIN_EVIDENCE_COLUMN),
      ),
    ),
  );
}

function printScenario(
  workbook: Awaited<ReturnType<typeof loadPgnWorkbook>>["workbook"],
  scenario: PgnTestScenario,
): void {
  const input = scenario.turns
    .map((turn) => turn.userInput.replace(/\s+/g, " ").trim())
    .join(" | ");
  console.log(
    `${scenario.testCaseId} | Row ${scenario.sourceRowNumber} | Status: ${scenario.rawStatus || "(blank)"}`,
  );
  console.log(`User Input: ${input}`);
  console.log(
    `Existing Bot Response: ${hasExistingResponse(workbook, scenario) ? "YES" : "NO"}`,
  );
  console.log(
    `Existing Evidence: ${hasExistingEvidence(workbook, scenario) ? "YES" : "NO"}`,
  );
}

export interface RetestValidationResult {
  selectedCount: number;
  readyCount: number;
  finalCleanupOnly: boolean;
  shouldExecute: boolean;
  readyToExecute: boolean;
}

export async function validateRetest(
  args = process.argv.slice(2),
  config: AppConfig = loadConfig(),
): Promise<RetestValidationResult> {
  const options = parseCliOptions(args);
  if (options.rerunAll || options.rerunIds.size) {
    throw new Error("--rerun is not used in retest mode; use --test instead");
  }
  if (
    options.resumeRunId &&
    (options.testIds.size > 0 || options.sheet !== undefined)
  ) {
    throw new Error("--resume cannot be combined with --test or --sheet");
  }
  const workbookPath = (await exists(config.pgnExecutedWorkbookPath))
    ? config.pgnExecutedWorkbookPath
    : config.pgnSourceWorkbookPath;
  const loaded = await loadPgnWorkbook(workbookPath);
  assertPgnWorkbookValid(loaded.parsed);
  const resumedRun = options.resumeRunId
    ? getRetestRunMetadata(loaded.workbook, options.resumeRunId)
    : undefined;
  if (options.resumeRunId && !resumedRun) {
    throw new Error(`Retest Run was not found: ${options.resumeRunId}`);
  }
  const selection = selectRetestScenarios(loaded.parsed.scenarios, {
    testIds: options.testIds,
    sheet: options.sheet,
    limit: options.limit,
    resumeSelectedIds: resumedRun?.selectedIds,
    completedIds: new Set(resumedRun?.finishedIds ?? []),
  });
  const finalCleanupOnly = needsFinalRetestCleanup(
    resumedRun,
    selection.selected.length,
  );

  console.log("PGN Retest Validation");
  console.log(`Workbook: ${path.relative(config.projectRoot, workbookPath)}`);
  console.log("");
  for (const [kind, title] of [
    ["kb", "Test Case Knowledge Base"],
    ["negative", "Negative Case"],
  ] as const) {
    console.log(title);
    console.log("-".repeat(title.length));
    console.log(`Ready for Re-test: ${selection.readyBySheet[kind].length}`);
    const selected = selection.selected.filter(
      (scenario) => scenario.sheetKind === kind,
    );
    if (selected.length) {
      console.log("Selected scenarios:");
      selected.forEach((scenario) => printScenario(loaded.workbook, scenario));
    }
    console.log("");
  }
  selection.warnings.forEach((warning) =>
    console.log(`WARNING: ${warning}`),
  );
  const unknownStatuses = loaded.parsed.issues.filter(
    (issue) => issue.code === "UNKNOWN_STATUS",
  );
  if (unknownStatuses.length) {
    console.log("Unknown statuses:");
    unknownStatuses.forEach((issue) =>
      console.log(`${issue.sheetName} row ${issue.rowNumber}: ${issue.message}`),
    );
    console.log("");
  }
  console.log(`Total selected scenarios: ${selection.selected.length}`);
  console.log(
    `Multi-turn scenarios: ${selection.selected.filter((scenario) => scenario.turns.length > 1).length}`,
  );
  if (finalCleanupOnly) {
    console.log("Final session cleanup: REQUIRED");
  }
  console.log("");

  let driveReady = false;
  if (!config.googleDriveEvidenceEnabled) {
    console.log("Google Drive Evidence: DISABLED");
  } else {
    try {
      const publisher = createGoogleDriveEvidencePublisher(config);
      const parent = await publisher.validateParentFolder();
      console.log(`Google Drive Evidence: READY (${parent.name}, ${parent.id})`);
      driveReady = true;
    } catch (error) {
      console.log(
        `Google Drive Evidence: ERROR (${safeGoogleCredentialError(error, config.googleServiceAccount?.value)})`,
      );
      driveReady = false;
    }
  }
  const profilePresent = await exists(config.profileDir);
  console.log(`WhatsApp profile: ${profilePresent ? "PRESENT" : "MISSING"}`);
  console.log("");
  const shouldExecute = selection.selected.length > 0 || finalCleanupOnly;
  const readyToExecute = !shouldExecute || (driveReady && profilePresent);
  if (!shouldExecute) {
    console.log("Nothing to execute.");
  } else if (driveReady && profilePresent) {
    console.log(
      finalCleanupOnly ? "READY TO COMPLETE RETEST CLEANUP" : "READY TO RETEST",
    );
  } else {
    console.log(
      finalCleanupOnly
        ? "NOT READY TO COMPLETE RETEST CLEANUP"
        : "NOT READY TO RETEST",
    );
  }
  return {
    selectedCount: selection.selected.length,
    readyCount:
      selection.readyBySheet.kb.length + selection.readyBySheet.negative.length,
    finalCleanupOnly,
    shouldExecute,
    readyToExecute,
  };
}

if (isEntrypoint(import.meta.url)) {
  runCliMain(async () => {
    const result = await validateRetest();
    if (result.shouldExecute && !result.readyToExecute) {
      process.exitCode = 1;
    }
  }, safeGoogleCredentialError);
}
