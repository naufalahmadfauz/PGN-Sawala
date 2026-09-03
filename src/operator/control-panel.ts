import { safeGoogleCredentialError } from "../evidence/google-service-account";
import type { OperatorUi } from "./ui";

export interface RetestReadiness {
  selectedCount: number;
  finalCleanupOnly: boolean;
  shouldExecute: boolean;
  readyToExecute: boolean;
}

export interface OperatorActions {
  validatePgn(): Promise<boolean>;
  prepareFresh(): Promise<void>;
  runPgn(args: string[]): Promise<void>;
  validateRetest(args: string[]): Promise<RetestReadiness>;
  runRetest(args: string[]): Promise<void>;
  validateEvidence(): Promise<{ ready: boolean }>;
  migrateEvidence(): Promise<void>;
  loginWhatsApp(): Promise<void>;
  verifyWhatsApp(): Promise<void>;
  recreateWhatsApp(): Promise<void>;
  setup(): Promise<void>;
  diagnostics(): Promise<void>;
  typecheck(): Promise<void>;
  regressionTests(): Promise<void>;
  createTemplate(): Promise<void>;
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
    ui.error(safeGoogleCredentialError(error));
    return { ok: false };
  }
}

function parseIds(value: string): string[] {
  return value
    .split(",")
    .map((item) => item.trim())
    .filter(Boolean);
}

async function confirmExecution(
  ui: OperatorUi,
  scope: string,
): Promise<boolean | undefined> {
  return ui.confirm({
    message: `${scope} will open WhatsApp, send reset and testcase messages, update the executed workbook, and may upload evidence. Continue?`,
    initialValue: false,
  });
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
    const confirmed = await confirmExecution(ui, scope);
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
    const readiness = await validateBeforeRetest(ui, actions, args);
    if (!readiness) continue;
    const description = readiness.finalCleanupOnly
      ? "The final WhatsApp session cleanup"
      : `${readiness.selectedCount} retest scenario(s)`;
    const confirmed = await confirmExecution(ui, description);
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

async function developerMenu(
  ui: OperatorUi,
  actions: OperatorActions,
): Promise<boolean> {
  while (true) {
    const choice = await ui.select({
      message: "Developer tools",
      options: [
        { value: "check", label: "TypeScript check" },
        { value: "tests", label: "Safe regression tests", hint: "No WhatsApp or Drive calls" },
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
  while (true) {
    const choice = await ui.select({
      message: "Choose an operation",
      options: [
        { value: "run", label: "Run tests" },
        { value: "retest", label: "Retest fixed cases" },
        { value: "validate", label: "Validate" },
        { value: "evidence", label: "Evidence" },
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
    if (choice === "evidence") keepRunning = await evidenceMenu(ui, actions);
    if (choice === "whatsapp") keepRunning = await whatsappMenu(ui, actions);
    if (choice === "developer") keepRunning = await developerMenu(ui, actions);
    if (choice === "setup") {
      const result = await attempt(ui, "Opening setup", actions.setup);
      keepRunning = result.ok;
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
