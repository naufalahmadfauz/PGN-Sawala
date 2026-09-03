export const PGN_TEST_STATUSES = {
  Passed: "Passed",
  Failed: "Failed",
  Blocked: "Blocked",
  Review: "Review",
  ReadyForRetest: "Ready for Re-test",
  PendingEvaluation: "Pending Evaluation",
} as const;

export type PgnTestStatus =
  (typeof PGN_TEST_STATUSES)[keyof typeof PGN_TEST_STATUSES];

export const PGN_TEST_STATUS_VALUES = Object.values(PGN_TEST_STATUSES);

const NORMALIZED_STATUSES = new Map<string, PgnTestStatus>([
  ["passed", PGN_TEST_STATUSES.Passed],
  ["failed", PGN_TEST_STATUSES.Failed],
  ["blocked", PGN_TEST_STATUSES.Blocked],
  ["review", PGN_TEST_STATUSES.Review],
  ["ready for re-test", PGN_TEST_STATUSES.ReadyForRetest],
  ["ready for retest", PGN_TEST_STATUSES.ReadyForRetest],
  ["pending evaluation", PGN_TEST_STATUSES.PendingEvaluation],
]);

export function normalizeTestStatus(value: string): PgnTestStatus | undefined {
  const normalized = value.trim().replace(/\s+/g, " ").toLowerCase();
  return normalized ? NORMALIZED_STATUSES.get(normalized) : undefined;
}

export function isReadyForRetest(value: string): boolean {
  return normalizeTestStatus(value) === PGN_TEST_STATUSES.ReadyForRetest;
}
