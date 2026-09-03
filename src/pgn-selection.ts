import type { Workbook } from "exceljs";
import type { CliOptions } from "./pgn-cli";
import {
  isScenarioComplete,
  isScenarioPartiallyComplete,
} from "./excel/pgn-workbook-loader";
import type { PgnTestScenario } from "./excel/pgn-types";

export function selectScenarios(
  scenarios: PgnTestScenario[],
  options: CliOptions,
  workbook: Workbook,
): { runnable: PgnTestScenario[]; skipped: string[] } {
  const allIds = new Set(scenarios.map((scenario) => scenario.testCaseId));
  const requestedIds = new Set([...options.testIds, ...options.rerunIds]);
  for (const id of [...options.testIds, ...options.rerunIds]) {
    if (!allIds.has(id)) {
      throw new Error(`Test Case ID was not found: ${id}`);
    }
  }

  let selected = scenarios.filter(
    (scenario) => !options.sheet || scenario.sheetKind === options.sheet,
  );
  if (requestedIds.size > 0) {
    const selectedIds = new Set(
      selected.map((scenario) => scenario.testCaseId),
    );
    for (const id of requestedIds) {
      if (!selectedIds.has(id)) {
        throw new Error(
          `Test Case ID ${id} is not in the selected ${options.sheet} sheet`,
        );
      }
    }
    selected = selected.filter((scenario) =>
      requestedIds.has(scenario.testCaseId),
    );
  }

  const skipped: string[] = [];
  const runnable = selected.filter((scenario) => {
    const rerun = options.rerunAll || options.rerunIds.has(scenario.testCaseId);
    if (!rerun && isScenarioPartiallyComplete(workbook, scenario)) {
      skipped.push(
        `${scenario.testCaseId}: partially completed; use --rerun ${scenario.testCaseId} to preserve multi-turn context`,
      );
      return false;
    }
    if (!rerun && isScenarioComplete(workbook, scenario)) {
      skipped.push(`${scenario.testCaseId}: already completed`);
      return false;
    }
    return true;
  });

  return {
    runnable: options.limit ? runnable.slice(0, options.limit) : runnable,
    skipped,
  };
}
