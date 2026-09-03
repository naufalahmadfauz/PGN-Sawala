import { PGN_TEST_STATUSES } from "../excel/pgn-test-status";
import type {
  PgnSheetKind,
  PgnTestScenario,
} from "../excel/pgn-types";

export interface RetestSelectionOptions {
  testIds?: ReadonlySet<string>;
  sheet?: PgnSheetKind;
  limit?: number;
  resumeSelectedIds?: readonly string[];
  completedIds?: ReadonlySet<string>;
}

export interface RetestSelection {
  selected: PgnTestScenario[];
  readyBySheet: Record<PgnSheetKind, PgnTestScenario[]>;
  warnings: string[];
}

export function selectRetestScenarios(
  scenarios: PgnTestScenario[],
  options: RetestSelectionOptions = {},
): RetestSelection {
  const inSelectedSheet = (scenario: PgnTestScenario): boolean =>
    !options.sheet || scenario.sheetKind === options.sheet;
  const readyBySheet: Record<PgnSheetKind, PgnTestScenario[]> = {
    kb: scenarios.filter(
      (scenario) =>
        scenario.sheetKind === "kb" &&
        scenario.status === PGN_TEST_STATUSES.ReadyForRetest,
    ),
    negative: scenarios.filter(
      (scenario) =>
        scenario.sheetKind === "negative" &&
        scenario.status === PGN_TEST_STATUSES.ReadyForRetest,
    ),
  };

  const scenariosById = new Map(
    scenarios.map((scenario) => [scenario.testCaseId, scenario]),
  );
  const requestedIds = options.resumeSelectedIds
    ? [...options.resumeSelectedIds]
    : [...(options.testIds ?? [])];
  for (const testCaseId of requestedIds) {
    const scenario = scenariosById.get(testCaseId);
    if (!scenario) {
      throw new Error(`Test Case ID was not found: ${testCaseId}`);
    }
    if (!inSelectedSheet(scenario)) {
      throw new Error(
        `Test Case ID ${testCaseId} is not in the selected ${options.sheet} sheet`,
      );
    }
  }

  let candidates: PgnTestScenario[];
  if (options.resumeSelectedIds) {
    candidates = requestedIds.map((testCaseId) => scenariosById.get(testCaseId)!);
  } else if (options.testIds?.size) {
    candidates = scenarios.filter(
      (scenario) =>
        inSelectedSheet(scenario) && options.testIds!.has(scenario.testCaseId),
    );
  } else {
    candidates = scenarios.filter(
      (scenario) =>
        inSelectedSheet(scenario) &&
        scenario.status === PGN_TEST_STATUSES.ReadyForRetest,
    );
  }

  const warnings = options.resumeSelectedIds
    ? []
    : candidates
        .filter(
          (scenario) =>
            options.testIds?.has(scenario.testCaseId) &&
            scenario.status !== PGN_TEST_STATUSES.ReadyForRetest,
        )
        .map(
          (scenario) =>
            `${scenario.testCaseId} was explicitly selected with Status "${scenario.rawStatus || "(blank)"}", not "${PGN_TEST_STATUSES.ReadyForRetest}".`,
        );

  const completedIds = options.completedIds ?? new Set<string>();
  candidates = candidates.filter(
    (scenario) => !completedIds.has(scenario.testCaseId),
  );
  return {
    selected: options.limit ? candidates.slice(0, options.limit) : candidates,
    readyBySheet,
    warnings,
  };
}
