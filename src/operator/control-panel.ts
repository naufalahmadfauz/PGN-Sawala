import { safeGoogleCredentialError } from "../evidence/google-service-account";
import {
  safeDiscordError,
  type DiscordValidationResult,
} from "../notifications/discord";
import type { SetupNextAction } from "./setup";
import type { OperatorUi } from "./ui";
import {
  formatRecoveryDiscovery,
  formatRecoveryValidation,
  type RecoveryMutationResult,
  type RecoveryRepairStrategy,
  type RecoveryValidation,
} from "../recovery/recovery-service";
import type { RecoveryDiscovery } from "../recovery/run-state";
import {
  CONTINUOUS_RECOVERY_WARNING,
  CONTINUOUS_SESSION_WARNING,
  readSessionMode,
  type SessionMode,
} from "../session-mode";

export interface RetestReadiness {
  selectedCount: number;
  finalCleanupOnly: boolean;
  shouldExecute: boolean;
  readyToExecute: boolean;
}

export interface OperatorActions {
  workbookSchema?(review: boolean, redetect?: boolean): Promise<void>;
  inspectRecovery?(): Promise<RecoveryDiscovery>;
  validateRecovery?(runId: string): Promise<RecoveryValidation>;
  resumeRecovery?(runId: string, acceptSourceDrift?: boolean): Promise<void>;
  restartRecovery?(runId: string, acceptSourceDrift?: boolean): Promise<void>;
  skipRecoveryScenario?(runId: string): Promise<RecoveryMutationResult>;
  repairRecovery?(
    runId: string,
    strategy: RecoveryRepairStrategy,
  ): Promise<RecoveryMutationResult>;
  abandonRecovery?(runId: string): Promise<RecoveryMutationResult>;
  validatePgn(): Promise<boolean>;
  prepareFresh(): Promise<void>;
  runPgn(args: string[]): Promise<void>;
  validateRetest(args: string[]): Promise<RetestReadiness>;
  runRetest(args: string[]): Promise<void>;
  validateEvidence(): Promise<{ ready: boolean }>;
  migrateEvidence(): Promise<void>;
  validateDiscord(sendTest: boolean): Promise<DiscordValidationResult>;
  configureNotifications(): Promise<void>;
  loginWhatsApp(): Promise<void>;
  verifyWhatsApp(): Promise<void>;
  recreateWhatsApp(): Promise<void>;
  setup(): Promise<SetupNextAction | void>;
  diagnostics(): Promise<void>;
  typecheck(): Promise<void>;
  regressionTests(): Promise<void>;
  createTemplate(): Promise<void>;
}

function recoveryActionsAvailable(actions: OperatorActions): actions is OperatorActions & {
  inspectRecovery: NonNullable<OperatorActions["inspectRecovery"]>;
  validateRecovery: NonNullable<OperatorActions["validateRecovery"]>;
  resumeRecovery: NonNullable<OperatorActions["resumeRecovery"]>;
  skipRecoveryScenario: NonNullable<OperatorActions["skipRecoveryScenario"]>;
  repairRecovery: NonNullable<OperatorActions["repairRecovery"]>;
  abandonRecovery: NonNullable<OperatorActions["abandonRecovery"]>;
} {
  return Boolean(
    actions.inspectRecovery &&
      actions.validateRecovery &&
      actions.resumeRecovery &&
      actions.skipRecoveryScenario &&
      actions.repairRecovery &&
      actions.abandonRecovery,
  );
}

interface ActionSuccess<Value> {
  ok: true;
  value: Value;
}

interface ActionFailure {
  ok: false;
}

async function attempt<Value>(
  ui: OperatorUi,
  label: string,
  action: () => Promise<Value>,
): Promise<ActionSuccess<Value> | ActionFailure> {
  ui.info(label);
  try {
    const value = await action();
    return { ok: true, value };
  } catch (error) {
    ui.error(
      safeGoogleCredentialError(new Error(safeDiscordError(error))),
    );
    return { ok: false };
  }
}

function discordValidationLines(result: DiscordValidationResult): string {
  const connectivity =
    result.connectivity === "ok"
      ? "OK"
      : result.connectivity === "failed"
        ? "FAILED"
        : "not tested";
  return [
    `Enabled ........ ${result.enabled ? "YES" : "NO"}`,
    `Webhook ........ ${result.configured && result.valid ? "configured" : "not configured"}`,
    `Connectivity ... ${connectivity}`,
  ].join("\n");
}

function parseIds(value: string): string[] {
  return value
    .split(",")
    .map((item) => item.trim())
    .filter(Boolean);
}

async function selectSessionMode(ui: OperatorUi): Promise<SessionMode | undefined> {
  const sessionMode = await ui.select<SessionMode>({
    message: "Session Mode for this run",
    options: [
      {
        value: "isolated",
        label: "Isolated (recommended)",
        hint: "Reset before each scenario and after the run",
      },
      {
        value: "continuous",
        label: "Continuous",
        hint: "One initial reset; all scenarios share conversation context",
      },
    ],
    initialValue: "isolated",
  });
  if (sessionMode === "continuous") ui.warn(CONTINUOUS_SESSION_WARNING);
  return sessionMode;
}

async function confirmExecution(
  ui: OperatorUi,
  scope: string,
  sessionMode: SessionMode = "isolated",
): Promise<boolean | undefined> {
  return ui.confirm({
    message: `${scope}${sessionMode === "continuous" ? " in Continuous Session Mode" : ""} will open WhatsApp, send reset and testcase messages, update the executed workbook, and may upload evidence. Continue?`,
    initialValue: false,
  });
}

function isContinuousRecovery(
  recovery: Pick<RecoveryValidation, "state" | "manifest">,
): boolean {
  return readSessionMode(recovery.state.sessionMode) === "continuous" ||
    readSessionMode(recovery.manifest.sessionMode) === "continuous";
}

async function repairRecoveryMismatch(
  ui: OperatorUi,
  actions: OperatorActions & {
    repairRecovery: NonNullable<OperatorActions["repairRecovery"]>;
  },
  validation: RecoveryValidation,
): Promise<boolean> {
  const mismatches = validation.reconciliation?.mismatchedScenarioIds ?? [];
  if (!mismatches.length) return true;
  ui.warn(
    `Recovery progress disagrees for: ${mismatches.join(", ")}. No scenario will be guessed as complete.`,
  );
  const strategy = await ui.select<RecoveryRepairStrategy | "back">({
    message: "How should recovery reconcile this progress?",
    options: [
      {
        value: "rerun",
        label: "Re-run mismatched scenarios",
        hint: "Safest; completed evidence is preserved",
      },
      { value: "artifacts", label: "Trust workbook and transcript" },
      { value: "checkpoint", label: "Trust checkpoint records" },
      { value: "back", label: "Back" },
    ],
    initialValue: "rerun",
  });
  if (strategy === undefined || strategy === "back") return false;
  const confirmed = await ui.confirm({
    message: `Apply the ${strategy} reconciliation strategy to Run ${validation.runId}?`,
    initialValue: false,
  });
  if (!confirmed) return false;
  const repaired = await attempt(ui, "Saving recovery reconciliation", () =>
    actions.repairRecovery(validation.runId, strategy),
  );
  if (!repaired.ok) return false;
  if (repaired.value.warning) ui.warn(repaired.value.warning);
  return true;
}

async function resumeRecoveryFromMenu(
  ui: OperatorUi,
  actions: OperatorActions & {
    validateRecovery: NonNullable<OperatorActions["validateRecovery"]>;
    resumeRecovery: NonNullable<OperatorActions["resumeRecovery"]>;
    repairRecovery: NonNullable<OperatorActions["repairRecovery"]>;
  },
  runId: string,
  discoveredIsDemo: boolean,
  discoveredIsContinuous = false,
): Promise<void> {
  if (discoveredIsContinuous) {
    ui.warn(CONTINUOUS_RECOVERY_WARNING);
    return;
  }
  let inspected = await attempt(ui, "Validating recovery without executing", () =>
    actions.validateRecovery(runId),
  );
  if (!inspected.ok) return;
  const isDemo =
    discoveredIsDemo ||
    inspected.value.state.isDemo === true ||
    inspected.value.manifest.isDemo === true;
  ui.note(
    formatRecoveryValidation(inspected.value),
    isDemo ? "DEMO recovery validation" : "Recovery validation",
  );
  if (isContinuousRecovery(inspected.value)) {
    ui.warn(CONTINUOUS_RECOVERY_WARNING);
    return;
  }
  if (isDemo) {
    const { state, manifest, reconciliation } = inspected.value;
    const finished = new Set([
      ...state.completedScenarioIds,
      ...state.skippedScenarioIds,
    ]);
    const remaining = state.selectedScenarioIds.filter((id) => !finished.has(id));
    const nextScenario = reconciliation?.nextScenarioId ?? remaining[0];
    const interruptedScenario =
      reconciliation?.interruptedScenarioId ?? state.activeScenarioId;
    const turnCount = manifest.scenarios.find(
      (scenario) => scenario.testCaseId === interruptedScenario,
    )?.turnCount;
    ui.note(
      [
        "DEMO: UI/testing only; no live execution. No WhatsApp, Playwright, Drive, or Discord actions.",
        `Checkpoint progress: ${state.completedScenarioIds.length} completed, ${state.skippedScenarioIds.length} skipped, ${remaining.length} remaining`,
        `Previous interruption: ${state.interruptionReason ?? "not recorded"}`,
        ...(interruptedScenario
          ? [`Restart interrupted scenario: ${interruptedScenario} from Turn 1${turnCount ? ` of ${turnCount}` : ""} (preview only)`]
          : []),
        `Next scenario: ${nextScenario ? `${nextScenario} from Turn 1 (preview only)` : "none; no remaining scenarios"}`,
        "This preview did not change recovery progress or artifacts.",
      ].join("\n"),
      "DEMO recovery preview",
    );
    if (!inspected.value.ready) {
      ui.warn("DEMO validation remains BLOCKED; no repair or execution was attempted.");
    }
    return;
  }
  if (!(await repairRecoveryMismatch(ui, actions, inspected.value))) return;
  if (inspected.value.reconciliation?.mismatchedScenarioIds.length) {
    inspected = await attempt(ui, "Revalidating reconciled recovery", () =>
      actions.validateRecovery(runId),
    );
    if (!inspected.ok) return;
    ui.note(formatRecoveryValidation(inspected.value), "Recovery validation");
    if (isContinuousRecovery(inspected.value)) {
      ui.warn(CONTINUOUS_RECOVERY_WARNING);
      return;
    }
    if (inspected.value.state.isDemo || inspected.value.manifest.isDemo) {
      ui.warn("DEMO recovery is preview only; no testcase was executed.");
      return;
    }
  }
  if (!inspected.value.ready) {
    ui.warn("Recovery prerequisites are not ready. No testcase was executed.");
    return;
  }
  let acceptSourceDrift = false;
  if (inspected.value.sourceDrift === "formatting-only") {
    const acceptDrift = await ui.confirm({
      message: "The source file hash changed, but testcase inputs are unchanged. Continue with this formatting-only drift?",
      initialValue: false,
    });
    if (!acceptDrift) return;
    acceptSourceDrift = true;
  }
  const nextScenario = inspected.value.reconciliation?.nextScenarioId;
  const confirmed = await ui.confirm({
    message: `Resume Run ${runId}${nextScenario ? ` at ${nextScenario} from Turn 1` : ""}? This will open WhatsApp, send messages, update the workbook, and reuse existing Drive artifacts.`,
    initialValue: false,
  });
  if (!confirmed) {
    ui.info("Recovery execution cancelled");
    return;
  }
  const resumed = await attempt(ui, `Resuming Run ${runId}`, () =>
    actions.resumeRecovery(runId, acceptSourceDrift),
  );
  if (resumed.ok) ui.success(`Run ${runId} recovery finished`);
}

async function restartRecoveryFromMenu(
  ui: OperatorUi,
  actions: OperatorActions & {
    validateRecovery: NonNullable<OperatorActions["validateRecovery"]>;
  },
  runId: string,
  discoveredIsDemo: boolean,
): Promise<void> {
  const inspected = await attempt(ui, "Validating full continuous restart without executing", () =>
    actions.validateRecovery(runId),
  );
  if (!inspected.ok) return;
  const validation = inspected.value;
  const isDemo = discoveredIsDemo || validation.state.isDemo === true ||
    validation.manifest.isDemo === true || runId.startsWith("DEMO-RECOVERY-");
  ui.note(formatRecoveryValidation(validation), isDemo ? "DEMO recovery validation" : "Recovery validation");
  if (isDemo) {
    ui.note([
      "DEMO: UI/testing only; no live execution. No WhatsApp, Playwright, Drive, or Discord actions.",
      `Full continuous restart: all ${validation.state.selectedScenarioIds.length} originally selected scenarios from the beginning (preview only).`,
      `Original order: ${validation.state.selectedScenarioIds.join(", ")}`,
      "A real restart creates a NEW Run ID and new evidence folder; the old run history is preserved.",
      "This preview did not change recovery progress or artifacts.",
    ].join("\n"), "DEMO continuous restart preview");
    return;
  }
  if (!isContinuousRecovery(validation) || validation.restartReady !== true ||
    !["unchanged", "formatting-only"].includes(validation.sourceDrift)) {
    ui.warn("Full continuous restart prerequisites are not ready. No testcase was executed.");
    return;
  }
  const restart = actions.restartRecovery;
  if (!restart) {
    ui.warn("Full continuous restart is unavailable. No testcase was executed.");
    return;
  }
  let acceptSourceDrift = false;
  if (validation.sourceDrift === "formatting-only") {
    const accepted = await ui.confirm({
      message: "The source file hash changed, but testcase inputs are unchanged. Accept this formatting-only drift for the full restart?",
      initialValue: false,
    });
    if (!accepted) return;
    acceptSourceDrift = true;
  }
  ui.warn(CONTINUOUS_SESSION_WARNING);
  const confirmed = await ui.confirm({
    message: `Restart ALL ${validation.state.selectedScenarioIds.length} originally selected scenarios in Run ${runId} from the beginning, including completed/skipped scenarios, in their original order? This opens WhatsApp and sends messages in Continuous Session Mode under a NEW Run ID and new evidence folder. The old run is marked ABANDONED only after the new checkpoint is saved; its history and artifacts are preserved. Continue with the FULL restart?`,
    initialValue: false,
  });
  if (!confirmed) {
    ui.info("Full continuous restart cancelled");
    return;
  }
  const restarted = await attempt(ui, `Restarting all scenarios from Run ${runId}`, () =>
    restart(runId, acceptSourceDrift),
  );
  if (restarted.ok) ui.success(`Continuous restart of Run ${runId} finished under a new Run ID`);
}

async function recoveryStartupMenu(
  ui: OperatorUi,
  actions: OperatorActions,
): Promise<boolean> {
  if (!recoveryActionsAvailable(actions)) return true;
  while (true) {
    const inspected = await attempt(ui, "Checking for interrupted PGN runs", () =>
      actions.inspectRecovery(),
    );
    if (!inspected.ok) return true;
    const recovery = inspected.value;
    if (recovery.kind === "none") return true;
    const isDemo =
      recovery.kind !== "unreadable" &&
      (recovery.state.isDemo === true || recovery.manifest.isDemo === true);
    ui.note(
      formatRecoveryDiscovery(recovery),
      recovery.kind === "recoverable" && isDemo
        ? "Recoverable run found [DEMO]"
        : "PGN run recovery",
    );
    if (recovery.kind === "running" || recovery.kind === "unreadable") {
      const choice = await ui.select({
        message:
          recovery.kind === "running"
            ? `A PGN run is already active${isDemo ? " [DEMO]" : ""}`
            : "Recovery state needs attention",
        options: [
          { value: "inspect", label: "Inspect again" },
          { value: "menu", label: "Continue to main menu" },
          { value: "exit", label: "Exit" },
        ],
      });
      if (choice === undefined || choice === "exit") return false;
      if (choice === "menu") return true;
      continue;
    }

    const runId = recovery.state.runId;
    const isContinuous = isContinuousRecovery(recovery);
    if (isContinuous) ui.warn(CONTINUOUS_RECOVERY_WARNING);
    const choice = await ui.select({
      message: `Recover interrupted Run ${runId}${isDemo ? " [DEMO]" : ""}`,
      options: isContinuous ? [
        { value: "inspect", label: "Inspect details", hint: "No messages sent" },
        {
          value: "restart",
          label: "Restart continuous run from beginning",
          hint: isDemo ? "Preview only; no changes" : "All original scenarios; new Run ID",
          disabled: !isDemo && !actions.restartRecovery,
        },
        { value: "abandon", label: "Abandon", hint: "Preserves artifacts" },
        { value: "menu", label: "Main menu" },
        { value: "exit", label: "Exit" },
      ] : [
        { value: "inspect", label: "Inspect recovery details", hint: "No messages sent" },
        { value: "resume", label: "Resume safely" },
        ...(isDemo
          ? [{ value: "restart", label: "Restart interrupted scenario", hint: "Preview Turn 1 only; no changes" }]
          : []),
        { value: "skip", label: "Skip current scenario and continue" },
        { value: "abandon", label: "Abandon recovery", hint: "Preserves artifacts" },
        { value: "menu", label: "Continue to main menu" },
        { value: "exit", label: "Exit" },
      ],
      ...(isContinuous ? { initialValue: "inspect" } : {}),
    });
    if (choice === undefined || choice === "exit") return false;
    if (choice === "menu") return true;
    if (choice === "inspect") {
      const validation = await attempt(ui, "Inspecting recovery details", () =>
        actions.validateRecovery(runId),
      );
      if (validation.ok) {
        ui.note(
          formatRecoveryValidation(validation.value),
          isDemo ? "DEMO recovery details" : "Recovery details",
        );
      }
      continue;
    }
    if (isContinuous) {
      if (choice === "restart") {
        await restartRecoveryFromMenu(ui, actions, runId, isDemo);
        continue;
      }
      if (choice !== "abandon") {
        ui.warn(CONTINUOUS_RECOVERY_WARNING);
        continue;
      }
    }
    if (choice === "resume" || (isDemo && choice === "restart")) {
      await resumeRecoveryFromMenu(ui, actions, runId, isDemo, isContinuous);
      continue;
    }
    if (choice === "skip") {
      const scenarioId =
        recovery.state.activeScenarioId ??
        recovery.state.selectedScenarioIds.find(
          (id) =>
            !recovery.state.completedScenarioIds.includes(id) &&
            !recovery.state.skippedScenarioIds.includes(id),
        );
      const confirmed = await ui.confirm({
        message: `Explicitly skip ${scenarioId ?? "the next incomplete scenario"} in Run ${runId}? Existing evidence will be preserved.`,
        initialValue: false,
      });
      if (!confirmed) continue;
      const skipped = await attempt(ui, "Saving operator skip", () =>
        actions.skipRecoveryScenario(runId),
      );
      if (!skipped.ok) continue;
      if (skipped.value.warning) ui.warn(skipped.value.warning);
      await resumeRecoveryFromMenu(ui, actions, runId, isDemo, isContinuous);
      continue;
    }
    if (choice !== "abandon") continue;
    const confirmed = await ui.confirm({
      message: `Abandon recovery for Run ${runId}? This keeps the workbook, transcript, evidence, and recovery history, but permits a fresh run.`,
      initialValue: false,
    });
    if (!confirmed) continue;
    const abandoned = await attempt(ui, `Abandoning Run ${runId}`, () =>
      actions.abandonRecovery(runId),
    );
    if (abandoned.ok) {
      if (abandoned.value.warning) ui.warn(abandoned.value.warning);
      ui.success(`Run ${runId} marked ABANDONED; artifacts were preserved`);
    }
  }
}

export async function runConfirmedFullTest(
  ui: OperatorUi,
  actions: OperatorActions,
): Promise<boolean> {
  const sessionMode = await selectSessionMode(ui);
  if (sessionMode === undefined) return false;
  const confirmed = await confirmExecution(ui, "A full test run", sessionMode);
  if (confirmed === undefined) return false;
  if (!confirmed) {
    ui.info("Test execution cancelled");
    return true;
  }
  const result = await attempt(ui, "Starting PGN execution", () =>
    actions.runPgn(sessionMode === "continuous" ? ["--session=continuous"] : []),
  );
  if (result.ok) ui.success("PGN execution finished");
  return true;
}

async function runTestsMenu(
  ui: OperatorUi,
  actions: OperatorActions,
): Promise<boolean> {
  while (true) {
    const choice = await ui.select({
      message: "Run tests",
      options: [
        { value: "validate", label: "Validate source workbook", hint: "No messages sent" },
        { value: "remaining", label: "Run remaining scenarios" },
        { value: "sheet", label: "Run one sheet" },
        { value: "ids", label: "Run specific testcase IDs" },
        { value: "fresh", label: "Prepare a fresh full run", hint: "Archives the current report" },
        { value: "back", label: "Back" },
      ],
    });
    if (choice === undefined) return false;
    if (choice === "back") return true;
    if (choice === "validate") {
      const result = await attempt(ui, "Validating source workbook", actions.validatePgn);
      if (result.ok) {
        result.value
          ? ui.success("Workbook validation passed")
          : ui.warn("Workbook validation reported errors");
      }
      continue;
    }
    if (choice === "fresh") {
      const confirmed = await ui.confirm({
        message: "Archive the current executed report and prepare a fresh workbook?",
        initialValue: false,
      });
      if (confirmed === undefined) return false;
      if (!confirmed) {
        ui.info("Fresh-run preparation cancelled");
        continue;
      }
      const result = await attempt(
        ui,
        "Preparing a fresh workbook",
        actions.prepareFresh,
      );
      if (result.ok) ui.success("Fresh workbook prepared");
      continue;
    }

    let args: string[] = [];
    let scope = "This test run";
    if (choice === "sheet") {
      const sheet = await ui.select({
        message: "Select a workbook sheet",
        options: [
          { value: "kb", label: "Test Case Knowledge Base" },
          { value: "negative", label: "Negative Case" },
          { value: "back", label: "Back" },
        ],
      });
      if (sheet === undefined) return false;
      if (sheet === "back") continue;
      args = ["--sheet", sheet];
      scope = `Running the ${sheet} sheet`;
    } else if (choice === "ids") {
      const input = await ui.text({
        message: "Testcase IDs",
        placeholder: "PGN-KB-003, PGN-NEG-018",
        validate: (value) =>
          parseIds(value).length ? undefined : "Enter at least one testcase ID",
      });
      if (input === undefined) return false;
      const ids = parseIds(input).join(",");
      const rerun = await ui.confirm({
        message: "Rerun these IDs even if results already exist?",
        initialValue: false,
      });
      if (rerun === undefined) return false;
      args = ["--test", ids, ...(rerun ? ["--rerun"] : [])];
      scope = `Running ${parseIds(input).length} selected testcase(s)`;
    }
    const sessionMode = await selectSessionMode(ui);
    if (sessionMode === undefined) return false;
    if (sessionMode === "continuous") args.push("--session=continuous");
    const confirmed = await confirmExecution(ui, scope, sessionMode);
    if (confirmed === undefined) return false;
    if (!confirmed) {
      ui.info("Test execution cancelled");
      continue;
    }
    const result = await attempt(ui, "Starting PGN execution", () =>
      actions.runPgn(args),
    );
    if (result.ok) ui.success("PGN execution finished");
  }
}

async function validateBeforeRetest(
  ui: OperatorUi,
  actions: OperatorActions,
  args: string[],
): Promise<RetestReadiness | undefined> {
  const validation = await attempt(ui, "Reviewing retest selection", () =>
    actions.validateRetest(args),
  );
  if (!validation.ok) return undefined;
  if (!validation.value.shouldExecute) {
    ui.info("No retest scenarios are selected. Nothing will be executed.");
    return undefined;
  }
  if (!validation.value.readyToExecute) {
    ui.warn("Retest prerequisites are not ready. Resolve validation errors first.");
    return undefined;
  }
  return validation.value;
}

async function retestMenu(
  ui: OperatorUi,
  actions: OperatorActions,
): Promise<boolean> {
  while (true) {
    const choice = await ui.select({
      message: "Retest fixed cases",
      options: [
        { value: "review", label: "Review approved candidates", hint: "No messages sent" },
        { value: "ready", label: "Run all Ready for Re-test cases" },
        { value: "ids", label: "Run selected testcase IDs" },
        { value: "resume", label: "Resume a retest run" },
        { value: "back", label: "Back" },
      ],
    });
    if (choice === undefined) return false;
    if (choice === "back") return true;
    if (choice === "review") {
      await attempt(ui, "Validating retest candidates", () =>
        actions.validateRetest([]),
      );
      continue;
    }

    let args: string[] = [];
    if (choice === "ids") {
      const input = await ui.text({
        message: "Retest testcase IDs",
        placeholder: "PGN-KB-031, PGN-KB-075",
        validate: (value) =>
          parseIds(value).length ? undefined : "Enter at least one testcase ID",
      });
      if (input === undefined) return false;
      args = ["--test", parseIds(input).join(",")];
    } else if (choice === "resume") {
      const runId = await ui.text({
        message: "Retest Run ID",
        placeholder: "RETEST-20260902T053000Z",
        validate: (value) =>
          value.trim() ? undefined : "Retest Run ID must not be empty",
      });
      if (runId === undefined) return false;
      args = ["--resume", runId.trim()];
    }
    let sessionMode: SessionMode = "isolated";
    if (choice !== "resume") {
      const selectedMode = await selectSessionMode(ui);
      if (selectedMode === undefined) return false;
      sessionMode = selectedMode;
      if (sessionMode === "continuous") args.push("--session=continuous");
    }
    const readiness = await validateBeforeRetest(ui, actions, args);
    if (!readiness) continue;
    const description = readiness.finalCleanupOnly
      ? "The final WhatsApp session cleanup"
      : `${readiness.selectedCount} retest scenario(s)`;
    const confirmed = await confirmExecution(ui, description, sessionMode);
    if (confirmed === undefined) return false;
    if (!confirmed) {
      ui.info("Retest execution cancelled");
      continue;
    }
    const result = await attempt(ui, "Starting retest execution", () =>
      actions.runRetest(args),
    );
    if (result.ok) ui.success("Retest execution finished");
  }
}

async function validationMenu(
  ui: OperatorUi,
  actions: OperatorActions,
): Promise<boolean> {
  while (true) {
    const choice = await ui.select({
      message: "Validation",
      options: [
        { value: "workbook", label: "Source workbook" },
        { value: "retest", label: "Retest readiness" },
        { value: "evidence", label: "Evidence and Drive" },
        { value: "back", label: "Back" },
      ],
    });
    if (choice === undefined) return false;
    if (choice === "back") return true;
    if (choice === "workbook") {
      await attempt(ui, "Validating source workbook", actions.validatePgn);
    } else if (choice === "retest") {
      await attempt(ui, "Validating retest readiness", () =>
        actions.validateRetest([]),
      );
    } else {
      await attempt(ui, "Validating evidence", actions.validateEvidence);
    }
  }
}

async function evidenceMenu(
  ui: OperatorUi,
  actions: OperatorActions,
): Promise<boolean> {
  while (true) {
    const choice = await ui.select({
      message: "Evidence",
      options: [
        { value: "validate", label: "Validate local evidence and Drive" },
        { value: "migrate", label: "Migrate evidence to Drive", hint: "Updates the executed workbook" },
        { value: "back", label: "Back" },
      ],
    });
    if (choice === undefined) return false;
    if (choice === "back") return true;
    if (choice === "validate") {
      await attempt(ui, "Validating evidence", actions.validateEvidence);
      continue;
    }
    const confirmed = await ui.confirm({
      message: "Back up the executed workbook, upload evidence, and write Drive links?",
      initialValue: false,
    });
    if (confirmed === undefined) return false;
    if (!confirmed) {
      ui.info("Evidence migration cancelled");
      continue;
    }
    const result = await attempt(ui, "Migrating evidence", actions.migrateEvidence);
    if (result.ok) ui.success("Evidence migration finished");
  }
}

async function whatsappMenu(
  ui: OperatorUi,
  actions: OperatorActions,
): Promise<boolean> {
  while (true) {
    const choice = await ui.select({
      message: "WhatsApp",
      options: [
        { value: "verify", label: "Verify saved session" },
        { value: "login", label: "Sign in or repair session" },
        { value: "recreate", label: "Recreate authentication", hint: "Clears the saved profile" },
        { value: "back", label: "Back" },
      ],
    });
    if (choice === undefined) return false;
    if (choice === "back") return true;
    if (choice === "verify") {
      await attempt(ui, "Verifying WhatsApp session", actions.verifyWhatsApp);
    } else if (choice === "login") {
      await attempt(ui, "Opening WhatsApp login", actions.loginWhatsApp);
    } else {
      const confirmed = await ui.confirm({
        message: "Delete the saved WhatsApp profile and require a new QR login?",
        initialValue: false,
      });
      if (confirmed === undefined) return false;
      if (!confirmed) {
        ui.info("Authentication recreation cancelled");
        continue;
      }
      await attempt(ui, "Recreating WhatsApp authentication", actions.recreateWhatsApp);
    }
  }
}

async function notificationsMenu(
  ui: OperatorUi,
  actions: OperatorActions,
): Promise<boolean> {
  while (true) {
    const choice = await ui.select({
      message: "Notifications",
      options: [
        {
          value: "status",
          label: "Discord status",
          hint: "Safely inspect the configured webhook; posts nothing",
        },
        { value: "test", label: "Test Discord notification" },
        { value: "configure", label: "Configure notifications" },
        { value: "back", label: "Back" },
      ],
    });
    if (choice === undefined) return false;
    if (choice === "back") return true;
    if (choice === "configure") {
      await attempt(
        ui,
        "Opening notification settings",
        actions.configureNotifications,
      );
      continue;
    }
    if (choice === "status") {
      const result = await attempt(ui, "Inspecting Discord webhook", () =>
        actions.validateDiscord(false),
      );
      if (result.ok) {
        ui.note(discordValidationLines(result.value), "Discord notifications");
        if (result.value.reason) {
          ui.warn(safeDiscordError(new Error(result.value.reason)));
        }
      }
      continue;
    }

    const confirmed = await ui.confirm({
      message: "Send one test notification to the configured Discord channel?",
      initialValue: false,
    });
    if (confirmed === undefined) return false;
    if (!confirmed) {
      ui.info("Discord test notification cancelled");
      continue;
    }
    const result = await attempt(ui, "Sending Discord test notification", () =>
      actions.validateDiscord(true),
    );
    if (result.ok) {
      if (result.value.testNotificationSent) {
        ui.success("Discord test notification sent");
      } else if (result.value.testNotificationDeliveryUncertain) {
        ui.warn(
          `Discord test notification delivery could not be confirmed; check the channel before retrying: ${safeDiscordError(new Error(result.value.reason ?? "request outcome is unknown"))}`,
        );
      } else {
        ui.warn(
          `Discord test notification was not sent: ${safeDiscordError(new Error(result.value.reason ?? "validation failed"))}`,
        );
      }
    }
  }
}

async function developerMenu(
  ui: OperatorUi,
  actions: OperatorActions,
): Promise<boolean> {
  while (true) {
    const choice = await ui.select({
      message: "Developer tools",
      options: [
        { value: "check", label: "TypeScript check" },
        {
          value: "tests",
          label: "Safe regression tests",
          hint: "No WhatsApp, Drive, or Discord calls",
        },
        { value: "template", label: "Create legacy workbook template" },
        { value: "back", label: "Back" },
      ],
    });
    if (choice === undefined) return false;
    if (choice === "back") return true;
    if (choice === "check") {
      await attempt(ui, "Running TypeScript check", actions.typecheck);
    } else if (choice === "tests") {
      await attempt(ui, "Running safe regression tests", actions.regressionTests);
    } else {
      await attempt(ui, "Creating workbook template", actions.createTemplate);
    }
  }
}

export async function runControlPanel(
  ui: OperatorUi,
  actions: OperatorActions,
): Promise<void> {
  ui.intro("PGN Sawala operator control panel");
  if (!(await recoveryStartupMenu(ui, actions))) {
    ui.cancel("Operator control panel closed");
    return;
  }
  while (true) {
    const choice = await ui.select({
      message: "Choose an operation",
      options: [
        { value: "run", label: "Run tests" },
        { value: "retest", label: "Retest fixed cases" },
        { value: "validate", label: "Validate" },
        ...(actions.workbookSchema ? [{ value: "workbook", label: "Workbook" }] : []),
        { value: "evidence", label: "Evidence" },
        { value: "notifications", label: "Notifications" },
        { value: "whatsapp", label: "WhatsApp" },
        { value: "setup", label: "Setup" },
        { value: "diagnostics", label: "Diagnostics" },
        { value: "developer", label: "Developer tools" },
        { value: "exit", label: "Exit" },
      ],
    });
    if (choice === undefined) {
      ui.cancel("Operator control panel closed");
      return;
    }
    if (choice === "exit") {
      ui.outro("Operator control panel closed");
      return;
    }
    let keepRunning = true;
    if (choice === "run") keepRunning = await runTestsMenu(ui, actions);
    if (choice === "retest") keepRunning = await retestMenu(ui, actions);
    if (choice === "validate") keepRunning = await validationMenu(ui, actions);
    if (choice === "workbook" && actions.workbookSchema) {
      while (true) {
        const operation = await ui.select({ message: "Workbook", options: [
          { value: "status", label: "Schema status" },
          { value: "review", label: "Review column mapping" },
          { value: "redetect", label: "Re-detect columns" },
          { value: "back", label: "Back" },
        ] });
        if (!operation) { keepRunning = false; break; }
        if (operation === "back") break;
        await attempt(ui, "Inspecting workbook schema without live execution", () => actions.workbookSchema!(operation === "review", operation === "redetect"));
      }
    }
    if (choice === "evidence") keepRunning = await evidenceMenu(ui, actions);
    if (choice === "notifications") {
      keepRunning = await notificationsMenu(ui, actions);
    }
    if (choice === "whatsapp") keepRunning = await whatsappMenu(ui, actions);
    if (choice === "developer") keepRunning = await developerMenu(ui, actions);
    if (choice === "setup") {
      const result = await attempt(ui, "Opening setup", actions.setup);
      keepRunning = result.ok;
      if (result.ok && result.value === "diagnostics") {
        const diagnostics = await attempt(
          ui,
          "Running diagnostics",
          actions.diagnostics,
        );
        keepRunning = diagnostics.ok;
      }
      if (result.ok && result.value === "full-test") {
        keepRunning = await runConfirmedFullTest(ui, actions);
      }
      if (result.ok && result.value === "exit") {
        ui.outro("Operator control panel closed");
        return;
      }
    }
    if (choice === "diagnostics") {
      const result = await attempt(ui, "Running diagnostics", actions.diagnostics);
      keepRunning = result.ok;
    }
    if (!keepRunning) {
      ui.cancel("Operator control panel closed");
      return;
    }
  }
}
