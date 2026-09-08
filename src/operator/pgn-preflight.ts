import { access } from "node:fs/promises";
import type { AppConfig } from "../config";
import { getRetestRunMetadata } from "../excel/retest-workbook";
import { loadPgnWorkbook } from "../excel/pgn-workbook-loader";
import { assertPgnWorkbookValid } from "../excel/pgn-workbook-validator";
import {
  assertResumeOptionsCompatible,
  parseCliOptions,
} from "../pgn-cli";
import type { PgnExecutionMode } from "../pgn-runner";
import { selectScenarios } from "../pgn-selection";
import { needsFinalRetestCleanup } from "../retest/retest-run";
import { selectRetestScenarios } from "../retest/retest-selection";
import { selectRecoveryScenarios } from "../recovery/recovery-service";
import { CONTINUOUS_RECOVERY_WARNING, readSessionMode, readExecutionTransport } from "../session-mode";
import { assertRestConfig } from "../rest/config";
import {
  assertRecoveryRunExecutable,
  readRecoveryRun,
  type RecoveryRunState,
} from "../recovery/run-state";

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
  assertResumeOptionsCompatible(options);
  if (mode === "retest" && (options.rerunAll || options.rerunIds.size)) {
    throw new Error("--rerun is not used in retest mode; use --test instead");
  }

  let resumedState: RecoveryRunState | undefined;
  const recoveryRunId = options.resumeRunId ?? options.restartRunId;
  if (recoveryRunId) {
    try {
      resumedState = (await readRecoveryRun(
        config.projectRoot,
        recoveryRunId,
      )).state;
    } catch (error) {
      if ((error as NodeJS.ErrnoException).code === "ENOENT") {
        assertRecoveryRunExecutable({ runId: recoveryRunId });
      }
      if (mode === "full" || options.restartRunId) throw error;
      if ((error as NodeJS.ErrnoException).code !== "ENOENT") throw error;
      // Retest runs created before filesystem checkpoints retain legacy resume support.
    }
    if (resumedState) {
      assertRecoveryRunExecutable(resumedState);
      const transport = readExecutionTransport(resumedState.transport);
      if (options.transportExplicit && options.transport !== transport) throw new Error("Transport conflicts with the stored run");
      options.transport = transport;
      const sessionMode = readSessionMode(resumedState.sessionMode);
      if (options.sessionModeExplicit && options.sessionMode !== sessionMode) {
        throw new Error("Session mode conflicts with the stored run; recovery cannot change isolation semantics");
      }
      if (options.resumeRunId && sessionMode === "continuous") throw new Error(CONTINUOUS_RECOVERY_WARNING);
      if (options.restartRunId && sessionMode !== "continuous") throw new Error("--restart-run is only for interrupted continuous runs");
      if (resumedState.mode !== mode) {
        throw new Error(
          `Run ${resumedState.runId} is a ${resumedState.mode} run, not a ${mode} run`,
        );
      }
    }
  }

  const workbookPath = (await exists(config.pgnExecutedWorkbookPath))
    ? config.pgnExecutedWorkbookPath
    : config.pgnSourceWorkbookPath;
  const loaded = await loadPgnWorkbook(workbookPath);
  assertPgnWorkbookValid(loaded.parsed);
  if (options.transport === "rest") assertRestConfig(config.livePersonRest);

  if (resumedState) {
    const selected = options.restartRunId
      ? resumedState.selectedScenarioIds.map((id) => {
          const scenario = loaded.parsed.scenarios.find((scenario) => scenario.testCaseId === id);
          if (!scenario) throw new Error(`Restart scenario was not found: ${id}`);
          return scenario;
        })
      : selectRecoveryScenarios(loaded.parsed.scenarios, resumedState);
    const finalCleanupOnly =
      selected.length === 0 && !resumedState.finalCleanupComplete;
    return {
      browserRequired: options.transport === "whatsapp" && (selected.length > 0 || finalCleanupOnly),
      selectedCount: selected.length,
      finalCleanupOnly,
    };
  }

  if (mode === "full") {
    const selection = selectScenarios(
      loaded.parsed.scenarios,
      options,
      loaded.workbook,
    );
    return {
      browserRequired: options.transport === "whatsapp" && selection.runnable.length > 0,
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
  if (resumedRun && readExecutionTransport(resumedRun.transport) !== options.transport) throw new Error("Transport conflicts with the stored retest");
  if (resumedRun && options.sessionModeExplicit && options.sessionMode !== readSessionMode(resumedRun.sessionMode)) {
    throw new Error("Session mode conflicts with the stored retest; recovery cannot change isolation semantics");
  }
  if (resumedRun && readSessionMode(resumedRun.sessionMode) === "continuous") throw new Error(CONTINUOUS_RECOVERY_WARNING);
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
    browserRequired: options.transport === "whatsapp" && (selection.selected.length > 0 || finalCleanupOnly),
    selectedCount: selection.selected.length,
    finalCleanupOnly,
  };
}
