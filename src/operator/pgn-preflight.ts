import { access } from "node:fs/promises";
import type { AppConfig } from "../config";
import { getRetestRunMetadata } from "../excel/retest-workbook";
import { loadPgnWorkbook } from "../excel/pgn-workbook-loader";
import { assertPgnWorkbookValid } from "../excel/pgn-workbook-validator";
import { parseCliOptions } from "../pgn-cli";
import type { PgnExecutionMode } from "../pgn-runner";
import { selectScenarios } from "../pgn-selection";
import { needsFinalRetestCleanup } from "../retest/retest-run";
import { selectRetestScenarios } from "../retest/retest-selection";

export interface PgnExecutionPreflight {
  browserRequired: boolean;
  selectedCount: number;
  finalCleanupOnly: boolean;
}

async function exists(filePath: string): Promise<boolean> {
  return access(filePath)
    .then(() => true)
    .catch(() => false);
}

export async function inspectPgnExecution(
  args: readonly string[],
  mode: PgnExecutionMode,
  config: AppConfig,
): Promise<PgnExecutionPreflight> {
  const options = parseCliOptions([...args]);
  if (mode === "full" && options.resumeRunId) {
    throw new Error("--resume is only available in retest mode");
  }
  if (mode === "retest" && (options.rerunAll || options.rerunIds.size)) {
    throw new Error("--rerun is not used in retest mode; use --test instead");
  }
  if (
    mode === "retest" &&
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

  if (mode === "full") {
    const selection = selectScenarios(
      loaded.parsed.scenarios,
      options,
      loaded.workbook,
    );
    return {
      browserRequired: selection.runnable.length > 0,
      selectedCount: selection.runnable.length,
      finalCleanupOnly: false,
    };
  }

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
  return {
    browserRequired: selection.selected.length > 0 || finalCleanupOnly,
    selectedCount: selection.selected.length,
    finalCleanupOnly,
  };
}
