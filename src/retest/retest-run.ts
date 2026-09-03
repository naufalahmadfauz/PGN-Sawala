export function createRetestRunId(now = new Date()): string {
  return `RETEST-${now
    .toISOString()
    .replace(/[-:]/g, "")
    .replace(/\.\d{3}Z$/, "Z")}`;
}

export function retestDriveFolderName(
  prefix: string,
  runId: string,
): string {
  return `${prefix}-${runId.replace(/^RETEST-/, "")}`;
}

export function needsFinalRetestCleanup(
  run:
    | {
        state: "IN_PROGRESS" | "COMPLETE";
        selectedIds: readonly string[];
        finishedIds: readonly string[];
      }
    | undefined,
  remainingScenarioCount: number,
): boolean {
  return Boolean(
    run &&
      run.state === "IN_PROGRESS" &&
      remainingScenarioCount === 0 &&
      run.selectedIds.length > 0 &&
      run.selectedIds.every((testCaseId) =>
        run.finishedIds.includes(testCaseId),
      ),
  );
}
