import assert from "node:assert/strict";
import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import test from "node:test";
import { parsePgnValidationArgs } from "./validate-pgn-workbook";
import { formatPgnValidation } from "../src/excel/pgn-workbook-validator";
import type { ParsedPgnWorkbook } from "../src/excel/pgn-types";
import {
  runControlPanel,
  type OperatorActions,
} from "../src/operator/control-panel";
import {
  detectBrowserRuntime,
  planBrowserRuntime,
  runBrowserAction,
} from "../src/operator/browser-runtime";
import {
  collectDiagnostics,
  formatDiagnosticReport,
  formatSetupChecklist,
  formatSetupInspection,
  type DiagnosticDependencies,
  type DiagnosticReport,
} from "../src/operator/diagnostics";
import {
  configureDiscordNotifications,
  runSetupWizard,
  updateEnvironmentText,
} from "../src/operator/setup";
import type {
  ConfirmPrompt,
  OperatorUi,
  SecretPrompt,
  SelectPrompt,
  TextPrompt,
} from "../src/operator/ui";
import type { RecoveryValidation } from "../src/recovery/recovery-service";
import {
  RECOVERY_SCHEMA_VERSION,
  type RecoveryDiscovery,
  type RecoveryRunManifest,
  type RecoveryRunState,
} from "../src/recovery/run-state";
import {
  CONTINUOUS_RECOVERY_WARNING,
  CONTINUOUS_SESSION_WARNING,
  continuousSessionWarning,
  type ExecutionTransport,
  type SessionMode,
} from "../src/session-mode";

class ScriptedUi implements OperatorUi {
  readonly events: string[] = [];
  readonly selectPrompts: SelectPrompt<string>[] = [];
  readonly confirmPrompts: ConfirmPrompt[] = [];
  readonly secretPrompts: SecretPrompt[] = [];

  constructor(private readonly responses: unknown[] = []) {}

  private response(kind: string): unknown {
    if (!this.responses.length) {
      throw new Error(`No scripted response remains for ${kind}`);
    }
    return this.responses.shift();
  }

  intro(title: string): void {
    this.events.push(`intro:${title}`);
  }

  outro(message: string): void {
    this.events.push(`outro:${message}`);
  }

  cancel(message: string): void {
    this.events.push(`cancel:${message}`);
  }

  async select<Value extends string>(
    prompt: SelectPrompt<Value>,
  ): Promise<Value | undefined> {
    this.selectPrompts.push(prompt as SelectPrompt<string>);
    this.events.push(`select:${prompt.message}`);
    return this.response("select") as Value | undefined;
  }

  async confirm(prompt: ConfirmPrompt): Promise<boolean | undefined> {
    this.confirmPrompts.push(prompt);
    this.events.push(`confirm:${prompt.message}`);
    return this.response("confirm") as boolean | undefined;
  }

  async text(prompt: TextPrompt): Promise<string | undefined> {
    this.events.push(`text:${prompt.message}`);
    const response = this.response("text") as string | undefined;
    if (response !== undefined) {
      const validation = prompt.validate?.(response);
      if (validation) throw new Error(`Invalid scripted response: ${validation}`);
    }
    return response;
  }

  async secret(prompt: SecretPrompt): Promise<string | undefined> {
    this.secretPrompts.push(prompt);
    this.events.push(`secret:${prompt.message}`);
    const response = this.response("secret") as string | undefined;
    if (response !== undefined) {
      const validation = prompt.validate?.(response);
      if (validation) throw new Error(`Invalid scripted secret: ${validation}`);
    }
    return response;
  }

  note(message: string, title = ""): void {
    this.events.push(`note:${title}:${message}`);
  }

  info(message: string): void {
    this.events.push(`info:${message}`);
  }

  success(message: string): void {
    this.events.push(`success:${message}`);
  }

  warn(message: string): void {
    this.events.push(`warn:${message}`);
  }

  error(message: string): void {
    this.events.push(`error:${message}`);
  }

  async task<Value>(
    message: string,
    operation: () => Promise<Value>,
    successMessage?: string,
  ): Promise<Value> {
    this.events.push(`task:${message}`);
    const result = await operation();
    this.events.push(`success:${successMessage ?? message}`);
    return result;
  }
}

function diagnosticReport(
  overrides: Partial<DiagnosticReport> = {},
): DiagnosticReport {
  const report: DiagnosticReport = {
    checks: [],
    browserRuntime: { mode: "direct", reason: "test fixture" },
    chromiumInstalled: true,
    profilePresent: true,
    environmentFilePresent: false,
    ready: true,
    ...overrides,
  };
  if (!overrides.checks) {
    report.checks = [
      { id: "os", label: "Operating system", status: "info", detail: "Linux" },
      {
        id: "node",
        label: "Node.js",
        status: "ok",
        detail: "v24.14.0 (requires 20.12 or newer)",
      },
      { id: "npm", label: "npm", status: "ok", detail: "11.9.0" },
      {
        id: "dependencies",
        label: "Dependencies",
        status: "ok",
        detail: "installed",
      },
      {
        id: "playwright",
        label: "Playwright",
        status: "ok",
        detail: "1.62.1",
      },
      {
        id: "chromium",
        label: "Playwright Chromium",
        status: report.chromiumInstalled ? "ok" : "error",
        detail: report.chromiumInstalled ? "installed" : "missing",
      },
      {
        id: "configuration",
        label: "Configuration",
        status: "ok",
        detail: "valid",
      },
      {
        id: "env",
        label: ".env",
        status: report.environmentFilePresent ? "ok" : "warn",
        detail: report.environmentFilePresent ? "found" : "not found",
      },
      {
        id: "source-workbook",
        label: "Source workbook",
        status: "ok",
        detail: "found",
      },
      {
        id: "executed-workbook",
        label: "Executed workbook",
        status: "ok",
        detail: "found",
      },
      {
        id: "whatsapp-profile",
        label: "WhatsApp profile",
        status: report.profilePresent ? "ok" : "error",
        detail: report.profilePresent ? "present" : "missing",
      },
      {
        id: "whatsapp-target",
        label: "WhatsApp target",
        status: "ok",
        detail: "configured by phone",
      },
      {
        id: "drive",
        label: "Google Drive",
        status: "ok",
        detail: "credentials and folder access verified",
      },
      {
        id: "discord",
        label: "Discord notifications",
        status: "warn",
        detail: "disabled; webhook not configured; connectivity not tested",
      },
      {
        id: "browser-runtime",
        label: "Browser runtime",
        status: report.browserRuntime.mode === "unavailable" ? "error" : "ok",
        detail: report.browserRuntime.reason,
      },
    ];
  }
  return report;
}

function stubActions(
  overrides: Partial<OperatorActions> = {},
): OperatorActions {
  return {
    validatePgn: async () => true,
    prepareFresh: async () => undefined,
    runPgn: async () => undefined,
    validateRetest: async () => ({
      selectedCount: 0,
      finalCleanupOnly: false,
      shouldExecute: false,
      readyToExecute: true,
    }),
    runRetest: async () => undefined,
    validateEvidence: async () => ({ ready: true }),
    migrateEvidence: async () => undefined,
    validateDiscord: async (sendTest) => ({
      enabled: true,
      configured: true,
      valid: true,
      connectivity: "ok",
      testNotificationSent: sendTest,
    }),
    configureNotifications: async () => undefined,
    loginWhatsApp: async () => undefined,
    verifyWhatsApp: async () => undefined,
    recreateWhatsApp: async () => undefined,
    setup: async () => undefined,
    diagnostics: async () => undefined,
    typecheck: async () => undefined,
    regressionTests: async () => undefined,
    createTemplate: async () => undefined,
    ...overrides,
  };
}

function recoveryFixture(isDemo = false, sessionMode?: SessionMode, transport?: ExecutionTransport): {
  discovery: RecoveryDiscovery;
  validation: RecoveryValidation;
} {
  const state: RecoveryRunState = {
    schemaVersion: RECOVERY_SCHEMA_VERSION,
    ...(isDemo ? { isDemo: true as const } : {}),
    ...(sessionMode ? { sessionMode } : {}),
    ...(sessionMode || transport ? { transport: transport ?? "whatsapp" } : {}),
    runId: "RECOVERY-OPERATOR-001",
    mode: "full",
    status: "INTERRUPTED",
    sourceWorkbookPath: "data/source.xlsx",
    sourceWorkbookHash: "a".repeat(64),
    executedWorkbookPath: "reports/executed.xlsx",
    selectedScenarioIds: ["TC-001", "TC-002"],
    completedScenarioIds: ["TC-001"],
    skippedScenarioIds: [],
    totalScenarios: 2,
    lastCompletedScenarioId: "TC-001",
    activeScenarioId: "TC-002",
    activeScenarioAttempt: 1,
    activeScenarioStartedAt: "2026-09-05T10:01:00.000Z",
    scenarioAttempts: [
      {
        scenarioId: "TC-001",
        attempt: 1,
        status: "COMPLETED",
        startedAt: "2026-09-05T10:00:00.000Z",
        finishedAt: "2026-09-05T10:00:30.000Z",
      },
      {
        scenarioId: "TC-002",
        attempt: 1,
        status: "INTERRUPTED",
        startedAt: "2026-09-05T10:01:00.000Z",
        finishedAt: "2026-09-05T10:01:30.000Z",
      },
    ],
    reconciliationDecisions: [],
    driveRunFolderId: "fixture-folder",
    driveRunFolderUrl:
      "https://drive.google.com/drive/folders/fixture-folder",
    finalCleanupComplete: false,
    workbookProgress: "Saved through TC-002 turn 1",
    metrics: {
      executedScenarios: 1,
      capturedScenarios: 1,
      timeouts: 0,
      technicalErrors: 0,
      evidenceCaptured: 2,
      evidenceUploaded: 2,
      evidenceUploadErrors: 0,
    },
    startedAt: "2026-09-05T10:00:00.000Z",
    updatedAt: "2026-09-05T10:01:30.000Z",
    heartbeatAt: "2026-09-05T10:01:30.000Z",
    interruptedAt: "2026-09-05T10:01:30.000Z",
    interruptionReason: isDemo
      ? "Interrupted during Turn 2 of 2; demo only."
      : "Process received SIGTERM",
    resumeCount: 0,
  };
  const manifest: RecoveryRunManifest = {
    schemaVersion: RECOVERY_SCHEMA_VERSION,
    ...(isDemo ? { isDemo: true as const } : {}),
    ...(sessionMode ? { sessionMode } : {}),
    ...(sessionMode || transport ? { transport: transport ?? "whatsapp" } : {}),
    runId: state.runId,
    sourceWorkbookHash: state.sourceWorkbookHash,
    createdAt: state.startedAt,
    scenarios: state.selectedScenarioIds.map((testCaseId, order) => ({
      testCaseId,
      sheetKind: "kb",
      sheetName: "Test Case Knowledge Base",
      sourceRowNumber: order + 2,
      order,
      turnCount: testCaseId === "TC-002" ? 2 : 1,
      inputHash: String(order + 1).repeat(64),
    })),
  };
  const discovery: RecoveryDiscovery = {
    kind: "recoverable",
    state,
    manifest,
    lock: { status: "unlocked" },
  };
  return {
    discovery,
    validation: {
      runId: state.runId,
      mode: state.mode,
      state,
      manifest,
      lock: { status: "unlocked" },
      sourceDrift: "unchanged",
      reconciliation: {
        checkpointCompletedIds: ["TC-001"],
        workbookCompletedIds: ["TC-001"],
        transcriptCompletedIds: ["TC-001"],
        artifactConfirmedIds: ["TC-001"],
        safeCompletedIds: ["TC-001"],
        reconciledScenarioIds: [],
        mismatchedScenarioIds: [],
        evidenceCapturedScenarioIds: ["TC-001"],
        evidenceUploadedScenarioIds: ["TC-001"],
        nextScenarioId: "TC-002",
        interruptedScenarioId: "TC-002",
        restartInterruptedScenarioFromTurnOne: true,
      },
      checks: [],
      ready: sessionMode !== "continuous",
      ...(sessionMode === "continuous" ? { restartReady: true } : {}),
    },
  };
}

function stubRecoveryActions(
  recovery: ReturnType<typeof recoveryFixture>,
  calls: string[],
  overrides: Partial<OperatorActions> = {},
): OperatorActions {
  const mutate = (action: string) => async (runId: string) => {
    calls.push(`${action}:${runId}`);
    return { runId };
  };
  return stubActions({
    inspectRecovery: async () => recovery.discovery,
    validateRecovery: async (runId) => {
      calls.push(`validate:${runId}`);
      return recovery.validation;
    },
    resumeRecovery: async (runId) => { calls.push(`resume:${runId}`); },
    restartRecovery: async (runId, acceptSourceDrift) => {
      calls.push(`restart:${runId}:${acceptSourceDrift}`);
    },
    skipRecoveryScenario: mutate("skip"),
    repairRecovery: mutate("repair"),
    abandonRecovery: mutate("abandon"),
    ...overrides,
  });
}

const REST_ENVIRONMENT = {
  LIVEPERSON_REST_ENABLED: "true",
  LIVEPERSON_ACCOUNT_ID: "123456",
  LIVEPERSON_CLIENT_ID: "mock-rest-client",
  LIVEPERSON_CLIENT_SECRET: "mock-rest-secret-never-log",
  LIVEPERSON_SKILL_ID: "42",
  LIVEPERSON_SENTINEL_DOMAIN: "sentinel.example.invalid",
  LIVEPERSON_IDP_DOMAIN: "idp.example.invalid",
  LIVEPERSON_ASYNC_MESSAGING_DOMAIN: "async.example.invalid",
  LIVEPERSON_MESSAGING_REST_DOMAIN: "messaging.example.invalid",
};

function restSetupReport(): DiagnosticReport {
  const report = diagnosticReport();
  report.checks.push({ id: "liveperson-rest", label: "LivePerson REST", status: "info", detail: "fixture; authentication not checked" });
  return report;
}

function restDiagnosticDependencies(projectRoot: string): DiagnosticDependencies {
  return {
    projectRoot,
    transport: "rest",
    platform: "linux",
    environment: { ...REST_ENVIRONMENT },
    npmVersion: async () => "11.0.0",
    packageVersion: async (name) => {
      assert(!["playwright", "googleapis", "sharp"].includes(name), `REST must not probe ${name}`);
      return "1.0.0";
    },
    pathExists: async (filePath) => {
      assert.doesNotMatch(filePath, /chromium|whatsapp-profile/i);
      return true;
    },
    chromiumExecutablePath: async () => assert.fail("REST must not inspect Chromium"),
    hasCommand: async () => assert.fail("REST must not probe Xvfb"),
    validateDrive: async () => assert.fail("REST must not validate Drive"),
    inspectDiscord: async () => assert.fail("Discord inspection must be explicit"),
    validateRest: async () => assert.fail("REST authentication must be explicit"),
    inspectWorkbookSchema: async () => ({ ready: true, detail: "fixture schema" }),
    inspectRecovery: async () => ({ kind: "none", lock: { status: "unlocked" } }),
  };
}

test("browser runtime is direct on Windows, macOS, displayed Linux, and headless Linux", () => {
  for (const input of [
    { platform: "win32" as const, headless: false, xvfbAvailable: false },
    { platform: "darwin" as const, headless: false, xvfbAvailable: false },
    {
      platform: "linux" as const,
      display: ":0",
      headless: false,
      xvfbAvailable: false,
    },
    { platform: "linux" as const, headless: true, xvfbAvailable: false },
  ]) {
    assert.equal(planBrowserRuntime(input).mode, "direct");
  }
});

test("headless Linux workspace uses xvfb only when it is available and needed", async () => {
  let probes = 0;
  const available = await detectBrowserRuntime({
    platform: "linux",
    environment: { CODESPACES: "true" },
    headless: false,
    hasCommand: async (command) => {
      probes += 1;
      assert.equal(command, "xvfb-run");
      return true;
    },
  });
  assert.equal(available.mode, "xvfb");
  assert.equal(probes, 1);
  assert.equal(
    planBrowserRuntime({
      platform: "linux",
      headless: false,
      xvfbAvailable: false,
    }).mode,
    "unavailable",
  );
});

test("browser actions stay in-process on Windows and use shell-free xvfb on Codespaces", async () => {
  let directCalls = 0;
  let spawnedCommand = "";
  let spawnedArgs: readonly string[] = [];
  await runBrowserAction({
    platform: "win32",
    environment: {},
    headless: false,
    projectRoot: "/fixture/project",
    scriptPath: "/fixture/project/scripts/action.ts",
    direct: async () => {
      directCalls += 1;
    },
    runCommand: async () => {
      throw new Error("Windows must not spawn xvfb");
    },
  });
  await runBrowserAction({
    platform: "linux",
    environment: { CODESPACES: "true" },
    headless: false,
    projectRoot: "/fixture/project",
    scriptPath: "/fixture/project/scripts/action.ts",
    args: ["--limit", "2"],
    hasCommand: async () => true,
    direct: async () => {
      directCalls += 1;
    },
    runCommand: async (command, args) => {
      spawnedCommand = command;
      spawnedArgs = args;
    },
  });
  assert.equal(directCalls, 1);
  assert.equal(spawnedCommand, "xvfb-run");
  assert.deepEqual(spawnedArgs.slice(0, 5), [
    "-a",
    process.execPath,
    "--import",
    "tsx",
    "/fixture/project/scripts/action.ts",
  ]);
  assert.deepEqual(spawnedArgs.slice(-2), ["--limit", "2"]);
});

test("a no-op workbook action does not require DISPLAY or probe xvfb", async () => {
  let directCalls = 0;
  let displayProbes = 0;
  await runBrowserAction({
    platform: "linux",
    environment: {},
    headless: false,
    projectRoot: "/fixture/project",
    scriptPath: "/fixture/project/scripts/action.ts",
    browserRequired: async () => false,
    hasCommand: async () => {
      displayProbes += 1;
      return false;
    },
    direct: async () => {
      directCalls += 1;
    },
  });
  assert.equal(directCalls, 1);
  assert.equal(displayProbes, 0);
});

test("first-time setup writes only prompted local configuration", async (context) => {
  const projectRoot = await mkdtemp(path.join(tmpdir(), "pgn-operator-first-"));
  context.after(() => rm(projectRoot, { recursive: true, force: true }));
  const ui = new ScriptedUi([
    true,
    "phone",
    "+62 812 3456 7890",
    false,
    false,
    false,
    "exit",
  ]);
  const result = await runSetupWizard(ui, {
    projectRoot,
    environment: {},
    diagnose: async () => diagnosticReport(),
  });
  const written = await readFile(path.join(projectRoot, ".env"), "utf8");
  assert.equal(result.environmentUpdated, true);
  assert.match(written, /PGN_WHATSAPP_PHONE=6281234567890/);
  assert.match(written, /PGN_WHATSAPP_CHAT=\n/);
  assert.match(written, /WHATSAPP_HEADLESS=false/);
  assert.match(written, /GOOGLE_DRIVE_EVIDENCE_ENABLED=false/);
  assert.match(written, /DISCORD_NOTIFICATIONS_ENABLED=false/);
  assert.equal(
    ui.confirmPrompts[0].message,
    "Would you like to review or update your setup?",
  );
  assert.equal(ui.confirmPrompts[0].active, "Yes");
  assert.equal(ui.confirmPrompts[0].inactive, "No, keep current settings");
});

test("REST setup opts in with masked credentials and no implicit authentication", async (context) => {
  const projectRoot = await mkdtemp(path.join(tmpdir(), "pgn-operator-rest-setup-"));
  context.after(() => rm(projectRoot, { recursive: true, force: true }));
  const ui = new ScriptedUi([
    true, "skip", false, false, false, true,
    "123456", REST_ENVIRONMENT.LIVEPERSON_CLIENT_ID,
    REST_ENVIRONMENT.LIVEPERSON_CLIENT_SECRET, "42", "", "", "", "", false, "exit",
  ]);
  const diagnostics: boolean[] = [];
  const result = await runSetupWizard(ui, {
    projectRoot, environment: {},
    diagnose: async (access) => { diagnostics.push(access); return restSetupReport(); },
    validateRest: async () => assert.fail("Declined REST authentication must not run"),
  });
  assert.equal(result.environmentUpdated, true);
  assert.deepEqual(diagnostics, [false, false, true]);
  const written = await readFile(path.join(projectRoot, ".env"), "utf8");
  assert.match(written, /LIVEPERSON_REST_ENABLED=true/);
  assert.match(written, /LIVEPERSON_ACCOUNT_ID=123456/);
  assert.match(written, /LIVEPERSON_SKILL_ID=42/);
  assert.match(written, /LIVEPERSON_MESSAGING_REST_DOMAIN=\n/);
  assert(written.includes(REST_ENVIRONMENT.LIVEPERSON_CLIENT_SECRET));
  assert.equal(ui.secretPrompts.length, 2);
  assert(ui.secretPrompts.every((prompt) => prompt.mask === "*" && prompt.clearOnError));
  assert.equal(ui.events.join("\n").includes(REST_ENVIRONMENT.LIVEPERSON_CLIENT_SECRET), false);
  assert.equal(ui.events.join("\n").includes(REST_ENVIRONMENT.LIVEPERSON_CLIENT_ID), false);
  assert.doesNotMatch(written, /(?:^|\n)(?:SESSION_MODE|EXECUTION_TRANSPORT|PGN_TRANSPORT)=/);
});

test("REST setup secret cancellation preserves the entire existing environment file", async (context) => {
  const projectRoot = await mkdtemp(path.join(tmpdir(), "pgn-operator-rest-cancel-"));
  context.after(() => rm(projectRoot, { recursive: true, force: true }));
  const source = "# unchanged\nUNRELATED_SECRET=fixture-existing\nLIVEPERSON_REST_ENABLED=false\n";
  await writeFile(path.join(projectRoot, ".env"), source);
  const ui = new ScriptedUi([true, "skip", false, false, false, true, "123456", "mock-client", undefined]);
  const result = await runSetupWizard(ui, {
    projectRoot, environment: {}, diagnose: async () => restSetupReport(),
    validateRest: async () => assert.fail("Cancellation must not authenticate"),
  });
  assert.equal(result.cancelled, true);
  assert.equal(result.environmentUpdated, false);
  assert.equal(await readFile(path.join(projectRoot, ".env"), "utf8"), source);
  assert.equal(ui.secretPrompts.at(-1)?.message, "LIVEPERSON_CLIENT_SECRET");
  assert.doesNotMatch(ui.events.join("\n"), /fixture-existing|mock-client/);
});

test("REST setup preserves process-managed settings and never copies their secrets into .env", async (context) => {
  const projectRoot = await mkdtemp(path.join(tmpdir(), "pgn-operator-rest-managed-"));
  context.after(() => rm(projectRoot, { recursive: true, force: true }));
  const original = "LIVEPERSON_CLIENT_SECRET=existing-file-secret\nUNRELATED=keep\n";
  await writeFile(path.join(projectRoot, ".env"), original);
  const environment = { ...REST_ENVIRONMENT };
  const ui = new ScriptedUi([true, "skip", false, false, false, false, "exit"]);
  await runSetupWizard(ui, {
    projectRoot, environment, diagnose: async () => restSetupReport(),
    validateRest: async () => assert.fail("Declined REST authentication must not run"),
  });
  const written = await readFile(path.join(projectRoot, ".env"), "utf8");
  assert(written.startsWith(original));
  assert.doesNotMatch(written, /LIVEPERSON_REST_ENABLED|mock-rest/);
  for (const name of Object.keys(REST_ENVIRONMENT) as Array<keyof typeof REST_ENVIRONMENT>) {
    assert.equal(environment[name], REST_ENVIRONMENT[name]);
  }
  assert.equal(ui.secretPrompts.length, 0);
  assert.match(ui.events.join("\n"), /managed by the process environment/);
  assert.doesNotMatch(ui.events.join("\n"), /mock-rest|existing-file-secret/);
});

test("REST setup keeps an existing local secret without displaying or replacing it", async (context) => {
  const projectRoot = await mkdtemp(path.join(tmpdir(), "pgn-operator-rest-keep-"));
  context.after(() => rm(projectRoot, { recursive: true, force: true }));
  await writeFile(path.join(projectRoot, ".env"), "LIVEPERSON_CLIENT_SECRET=existing-file-secret\n");
  const { LIVEPERSON_CLIENT_SECRET: _secret, ...environment } = REST_ENVIRONMENT;
  const ui = new ScriptedUi([true, "skip", false, false, false, true, false, "exit"]);
  await runSetupWizard(ui, { projectRoot, environment, diagnose: async () => restSetupReport() });
  assert.match(await readFile(path.join(projectRoot, ".env"), "utf8"), /LIVEPERSON_CLIENT_SECRET=existing-file-secret/);
  assert.equal(ui.secretPrompts.length, 0);
  assert.doesNotMatch(ui.events.join("\n"), /existing-file-secret/);
});

for (const outcome of ["decline", "success", "failure", "cancel"] as const) {
  test(`REST setup ${outcome} authentication is explicit and safely mocked`, async (context) => {
    const projectRoot = await mkdtemp(path.join(tmpdir(), "pgn-operator-rest-auth-"));
    context.after(() => rm(projectRoot, { recursive: true, force: true }));
    const ui = new ScriptedUi([false, outcome === "cancel" ? undefined : outcome !== "decline", ...(outcome === "cancel" ? [] : ["exit"])]);
    let calls = 0;
    const result = await runSetupWizard(ui, {
      projectRoot, environment: { ...REST_ENVIRONMENT }, diagnose: async () => restSetupReport(),
      validateRest: async (config) => {
        calls += 1;
        assert.equal(config.livePersonRest?.clientSecret, REST_ENVIRONMENT.LIVEPERSON_CLIENT_SECRET);
        if (outcome === "failure") throw new Error(`Rejected ${config.livePersonRest?.clientSecret}`);
      },
    });
    assert.equal(calls, ["success", "failure"].includes(outcome) ? 1 : 0);
    assert.equal(result.cancelled, outcome === "cancel");
    const prompt = ui.confirmPrompts.find((candidate) => candidate.message.startsWith("Validate LivePerson REST domains"));
    assert.equal(prompt?.initialValue, false);
    assert.match(prompt?.message ?? "", /real, harmless auth\/domain requests only; no conversations/);
    assert.equal(ui.events.join("\n").includes(REST_ENVIRONMENT.LIVEPERSON_CLIENT_SECRET), false);
    if (outcome === "failure") assert.match(ui.events.join("\n"), /REST validation did not pass/);
    await assert.rejects(readFile(path.join(projectRoot, ".env"), "utf8"), /ENOENT/);
  });
}

test("existing setup preserves comments and unrelated secret fields", async (context) => {
  const projectRoot = await mkdtemp(path.join(tmpdir(), "pgn-operator-existing-"));
  context.after(() => rm(projectRoot, { recursive: true, force: true }));
  const fixtureSecret = "fixture-secret-that-must-not-be-logged";
  await writeFile(
    path.join(projectRoot, ".env"),
    `# retained comment\nUNRELATED_SECRET=${fixtureSecret}\nPGN_WHATSAPP_CHAT=Existing chat\nWHATSAPP_HEADLESS=false\nGOOGLE_DRIVE_EVIDENCE_ENABLED=false\n`,
  );
  const ui = new ScriptedUi([true, "keep", true, false, false, "exit"]);
  await runSetupWizard(ui, {
    projectRoot,
    environment: {},
    diagnose: async () =>
      diagnosticReport({ environmentFilePresent: true }),
  });
  const written = await readFile(path.join(projectRoot, ".env"), "utf8");
  assert.match(written, /# retained comment/);
  assert.match(written, new RegExp(`UNRELATED_SECRET=${fixtureSecret}`));
  assert.match(written, /PGN_WHATSAPP_CHAT=Existing chat/);
  assert.match(written, /WHATSAPP_HEADLESS=true/);
  assert.doesNotMatch(ui.events.join("\n"), new RegExp(fixtureSecret));
});

test("notification setup uses a masked prompt and never sends implicitly", async (context) => {
  const projectRoot = await mkdtemp(path.join(tmpdir(), "pgn-operator-discord-"));
  context.after(() => rm(projectRoot, { recursive: true, force: true }));
  const webhook = "https://discord.com/api/webhooks/123456789/setup-token";
  let testCalls = 0;
  const ui = new ScriptedUi([true, webhook, "5", false]);
  const configured = await configureDiscordNotifications(ui, {
    projectRoot,
    environment: {},
    testDiscordWebhook: async () => {
      testCalls += 1;
      throw new Error("A declined test must not run");
    },
  });

  const written = await readFile(path.join(projectRoot, ".env"), "utf8");
  assert.equal(configured, true);
  assert.equal(testCalls, 0);
  assert.equal(ui.secretPrompts.length, 1);
  assert.equal(ui.secretPrompts[0]?.clearOnError, true);
  assert.equal(written.includes(webhook), true);
  assert.match(written, /DISCORD_NOTIFICATIONS_ENABLED=true/);
  assert.match(written, /DISCORD_PROGRESS_EVERY=5/);
  assert.match(written, /DISCORD_NOTIFY_PROGRESS=true/);
  assert.equal(ui.events.join("\n").includes(webhook), false);
});

test("notification setup posts one explicit test and stores custom progress", async (context) => {
  const projectRoot = await mkdtemp(path.join(tmpdir(), "pgn-operator-discord-test-"));
  context.after(() => rm(projectRoot, { recursive: true, force: true }));
  const webhook = "https://discord.com/api/webhooks/123456789/custom-token";
  const tested: string[] = [];
  const ui = new ScriptedUi([true, webhook, "custom", "7", "3", true]);
  await configureDiscordNotifications(ui, {
    projectRoot,
    environment: {},
    testDiscordWebhook: async (candidate) => {
      tested.push(candidate);
      return {
        enabled: true,
        configured: true,
        valid: true,
        connectivity: "ok",
        testNotificationSent: true,
      };
    },
  });

  const written = await readFile(path.join(projectRoot, ".env"), "utf8");
  assert.equal(tested.length, 1);
  assert.equal(tested[0] === webhook, true);
  assert.match(written, /DISCORD_PROGRESS_EVERY=7/);
  assert.match(written, /DISCORD_PROGRESS_MINUTES=3/);
  assert.match(ui.events.join("\n"), /Discord test notification sent/);
  assert.equal(ui.events.join("\n").includes(webhook), false);
});

test("notification setup preserves an existing webhook without displaying it", async (context) => {
  const projectRoot = await mkdtemp(path.join(tmpdir(), "pgn-operator-discord-keep-"));
  context.after(() => rm(projectRoot, { recursive: true, force: true }));
  const webhook = "https://discord.com/api/webhooks/123456789/existing-token";
  await writeFile(
    path.join(projectRoot, ".env"),
    `DISCORD_NOTIFICATIONS_ENABLED=true\nDISCORD_WEBHOOK_URL=${webhook}\nDISCORD_NOTIFY_PROGRESS=true\nDISCORD_PROGRESS_EVERY=5\n`,
  );
  const ui = new ScriptedUi([true, true, "final", false]);
  await configureDiscordNotifications(ui, {
    projectRoot,
    environment: {},
  });

  const written = await readFile(path.join(projectRoot, ".env"), "utf8");
  assert.equal(ui.secretPrompts.length, 0);
  assert.equal(written.includes(webhook), true);
  assert.match(written, /DISCORD_NOTIFY_START=false/);
  assert.match(written, /DISCORD_NOTIFY_PROGRESS=false/);
  assert.equal(ui.events.join("\n").includes(webhook), false);
});

test("cancelling the Discord secret prompt writes nothing", async (context) => {
  const projectRoot = await mkdtemp(path.join(tmpdir(), "pgn-operator-discord-cancel-"));
  context.after(() => rm(projectRoot, { recursive: true, force: true }));
  const ui = new ScriptedUi([true, undefined]);
  const configured = await configureDiscordNotifications(ui, {
    projectRoot,
    environment: {},
  });

  assert.equal(configured, false);
  await assert.rejects(
    readFile(path.join(projectRoot, ".env"), "utf8"),
    /ENOENT/,
  );
});

test("notification setup does not shadow process-managed Discord settings", async (context) => {
  const projectRoot = await mkdtemp(path.join(tmpdir(), "pgn-operator-discord-env-"));
  context.after(() => rm(projectRoot, { recursive: true, force: true }));
  const ui = new ScriptedUi(["5", false]);
  const configured = await configureDiscordNotifications(ui, {
    projectRoot,
    environment: {
      DISCORD_NOTIFICATIONS_ENABLED: "true",
      DISCORD_WEBHOOK_URL:
        "https://discord.com/api/webhooks/123456789/process-token",
    },
  });

  assert.equal(configured, true);
  assert.equal(
    ui.confirmPrompts.some(
      (prompt) => prompt.message === "Enable Discord notifications?",
    ),
    false,
  );
  assert.equal(ui.secretPrompts.length, 0);
  assert.match(ui.events.join("\n"), /managed by the process environment/);
  const written = await readFile(path.join(projectRoot, ".env"), "utf8");
  assert.match(written, /DISCORD_PROGRESS_EVERY=5/);
  assert.doesNotMatch(written, /DISCORD_NOTIFICATIONS_ENABLED/);
  assert.doesNotMatch(written, /DISCORD_WEBHOOK_URL/);
});

test("notification setup configures local cadence with a process-managed webhook", async (context) => {
  const projectRoot = await mkdtemp(path.join(tmpdir(), "pgn-operator-discord-mixed-"));
  context.after(() => rm(projectRoot, { recursive: true, force: true }));
  const webhook =
    "https://discord.com/api/webhooks/123456789/process-secret-token";
  const ui = new ScriptedUi([true, "10", false]);
  const configured = await configureDiscordNotifications(ui, {
    projectRoot,
    environment: { DISCORD_WEBHOOK_URL: webhook },
  });

  const written = await readFile(path.join(projectRoot, ".env"), "utf8");
  assert.equal(configured, true);
  assert.equal(ui.secretPrompts.length, 0);
  assert.match(written, /DISCORD_NOTIFICATIONS_ENABLED=true/);
  assert.match(written, /DISCORD_PROGRESS_EVERY=10/);
  assert.equal(written.includes(webhook), false);
  assert.equal(ui.events.join("\n").includes(webhook), false);
  assert.match(ui.events.join("\n"), /managed by the process environment/);
});

test("setup refuses to write a missing service-account file", async (context) => {
  const projectRoot = await mkdtemp(path.join(tmpdir(), "pgn-operator-credential-"));
  context.after(() => rm(projectRoot, { recursive: true, force: true }));
  const ui = new ScriptedUi([
    true,
    "skip",
    false,
    true,
    "https://drive.google.com/drive/folders/abcdefghijklmno",
    "file",
    ".secrets/missing.json",
    false,
  ]);
  const result = await runSetupWizard(ui, {
    projectRoot,
    environment: {},
    diagnose: async () => diagnosticReport(),
    validateCredentialFile: async () => {
      throw new Error("file not found");
    },
  });
  assert.equal(result.environmentUpdated, false);
  await assert.rejects(
    readFile(path.join(projectRoot, ".env"), "utf8"),
    /ENOENT/,
  );
  assert.match(ui.events.join("\n"), /Credential file is not valid/);
});

test("setup offers the platform-appropriate Chromium installation", async (context) => {
  const projectRoot = await mkdtemp(path.join(tmpdir(), "pgn-operator-browser-"));
  context.after(() => rm(projectRoot, { recursive: true, force: true }));
  let diagnosticsCalls = 0;
  const installs: boolean[] = [];
  const ui = new ScriptedUi([false, "dependencies", "exit"]);
  const result = await runSetupWizard(ui, {
    projectRoot,
    platform: "linux",
    environment: {},
    diagnose: async () => {
      diagnosticsCalls += 1;
      return diagnosticReport({
        chromiumInstalled: diagnosticsCalls >= 3,
      });
    },
    installChromium: async (withDependencies) => {
      installs.push(withDependencies);
    },
  });
  assert.equal(result.chromiumInstalled, true);
  assert.deepEqual(installs, [true]);
  assert.match(
    ui.events.join("\n"),
    /npx playwright install --with-deps chromium/,
  );
});

test("fully configured setup shows one detailed inspection and a compact final checklist", async (context) => {
  const projectRoot = await mkdtemp(path.join(tmpdir(), "pgn-operator-ready-"));
  context.after(() => rm(projectRoot, { recursive: true, force: true }));
  await writeFile(path.join(projectRoot, ".env"), "PGN_WHATSAPP_PHONE=628123456789\n");
  const ui = new ScriptedUi([false, "exit"]);
  await runSetupWizard(ui, {
    projectRoot,
    environment: {},
    diagnose: async () =>
      diagnosticReport({ environmentFilePresent: true, ready: true }),
  });

  const output = ui.events.join("\n");
  assert.equal(output.match(/note:Environment:/g)?.length, 1);
  assert.equal(output.match(/note:Setup checklist:/g)?.length, 1);
  assert.match(output, /Google Drive configuration: detected/);
  assert.match(output, /✓ Node\.js 24\.14\.0/);
  assert.match(output, /✓ npm 11\.9\.0/);
  assert.match(output, /✓ Dependencies installed/);
  assert.match(output, /✓ Chromium installed/);
  assert.match(output, /✓ Source workbook found/);
  assert.match(output, /✓ Executed workbook found/);
  assert.match(output, /✓ WhatsApp profile present/);
  assert.match(output, /✓ Google Drive access verified/);
  assert.match(output, /success:Setup complete/);

  const nextPrompt = ui.selectPrompts.at(-1);
  assert.equal(nextPrompt?.message, "What would you like to do next?");
  assert.deepEqual(
    nextPrompt?.options.map((option) => option.label),
    ["Run diagnostics", "Open main menu", "Start full test", "Exit"],
  );
});

test("setup with missing Drive credentials keeps details out of the final checklist", async (context) => {
  const projectRoot = await mkdtemp(path.join(tmpdir(), "pgn-operator-no-drive-"));
  context.after(() => rm(projectRoot, { recursive: true, force: true }));
  const missingCredential = diagnosticReport({
    environmentFilePresent: true,
    ready: false,
  });
  missingCredential.checks = missingCredential.checks.map((check) =>
    check.id === "drive"
      ? {
          ...check,
          status: "error",
          detail:
            "Configure GOOGLE_SERVICE_ACCOUNT_JSON, GOOGLE_SERVICE_ACCOUNT_JSON_BASE64, or GOOGLE_SERVICE_ACCOUNT_FILE",
        }
      : check,
  );
  const ui = new ScriptedUi([false, "exit"]);
  await runSetupWizard(ui, {
    projectRoot,
    environment: {},
    diagnose: async () => missingCredential,
  });

  const output = ui.events.join("\n");
  assert.match(output, /Google Drive configuration: needs attention/);
  assert.match(output, /Google Drive access not verified/);
  const finalChecklist = output.split("note:Setup checklist:")[1] ?? "";
  assert.doesNotMatch(finalChecklist, /GOOGLE_SERVICE_ACCOUNT_JSON/);
  assert.equal(
    ui.selectPrompts.at(-1)?.options.find(
      (option) => option.value === "full-test",
    )?.disabled,
    true,
  );
});

test("setup with a missing WhatsApp profile offers login and reports it concisely", async (context) => {
  const projectRoot = await mkdtemp(path.join(tmpdir(), "pgn-operator-no-profile-"));
  context.after(() => rm(projectRoot, { recursive: true, force: true }));
  let loginCalls = 0;
  const ui = new ScriptedUi([false, false, "exit"]);
  await runSetupWizard(ui, {
    projectRoot,
    environment: {},
    diagnose: async () =>
      diagnosticReport({ profilePresent: false, ready: false }),
    loginWhatsApp: async () => {
      loginCalls += 1;
    },
  });
  assert.equal(loginCalls, 0);
  assert.match(ui.events.join("\n"), /✗ WhatsApp profile missing/);
  assert.ok(
    ui.confirmPrompts.some(
      (prompt) => prompt.message === "Open WhatsApp login now?",
    ),
  );
});

test("Codespaces setup explains automatic Xvfb without implementation noise", async (context) => {
  const projectRoot = await mkdtemp(path.join(tmpdir(), "pgn-operator-codespaces-"));
  context.after(() => rm(projectRoot, { recursive: true, force: true }));
  const report = diagnosticReport({
    environmentFilePresent: true,
    browserRuntime: {
      mode: "xvfb",
      reason: "headed Chromium needs a virtual display",
    },
  });
  report.checks = report.checks.map((check) => {
    if (check.id === "os") {
      return { ...check, detail: "Linux (GitHub Codespaces)" };
    }
    if (check.id === "browser-runtime") {
      return { ...check, detail: report.browserRuntime.reason };
    }
    return check;
  });
  const ui = new ScriptedUi([false, "exit"]);
  await runSetupWizard(ui, {
    projectRoot,
    platform: "linux",
    environment: { CODESPACES: "true" },
    diagnose: async () => report,
  });
  const output = ui.events.join("\n");
  assert.match(
    output,
    /Browser runtime: Codespaces detected; Xvfb will be used automatically/,
  );
  assert.doesNotMatch(output, /headed Chromium needs a virtual display/);
});

test("Windows setup reports a ready browser runtime without mentioning Xvfb", async (context) => {
  const projectRoot = await mkdtemp(path.join(tmpdir(), "pgn-operator-windows-"));
  context.after(() => rm(projectRoot, { recursive: true, force: true }));
  const report = diagnosticReport({ environmentFilePresent: true });
  report.checks = report.checks.map((check) =>
    check.id === "os" ? { ...check, detail: "Windows" } : check,
  );
  const ui = new ScriptedUi([false, "exit"]);
  await runSetupWizard(ui, {
    projectRoot,
    platform: "win32",
    environment: {},
    diagnose: async () => report,
  });
  const output = ui.events.join("\n");
  assert.match(output, /Browser runtime: ready/);
  assert.doesNotMatch(output, /Xvfb/);
});

test("setup cancellation exits before configuration or completion choices", async (context) => {
  const projectRoot = await mkdtemp(path.join(tmpdir(), "pgn-operator-cancel-"));
  context.after(() => rm(projectRoot, { recursive: true, force: true }));
  const ui = new ScriptedUi([undefined]);
  const result = await runSetupWizard(ui, {
    projectRoot,
    environment: {},
    diagnose: async () => diagnosticReport(),
  });
  assert.equal(result.cancelled, true);
  assert.equal(result.nextAction, undefined);
  assert.doesNotMatch(ui.events.join("\n"), /Setup complete/);
});

test("diagnostics report missing Chromium, env, workbook output, and WhatsApp session", async (context) => {
  const projectRoot = await mkdtemp(path.join(tmpdir(), "pgn-operator-diagnostics-"));
  context.after(() => rm(projectRoot, { recursive: true, force: true }));
  const report = await collectDiagnostics({
    projectRoot,
    platform: "linux",
    environment: {},
    npmVersion: async () => "11.0.0",
    packageVersion: async () => "1.0.0",
    chromiumExecutablePath: async () => "/fixture/chromium",
    pathExists: async (filePath) =>
      filePath.endsWith(
        "PGN AI Assistant - Knowledge Base Testing Report - User Inputs.xlsx",
      ),
    hasCommand: async () => false,
    checkDriveAccess: false,
  });
  const byId = new Map(report.checks.map((check) => [check.id, check]));
  assert.equal(byId.get("chromium")?.status, "error");
  assert.equal(byId.get("env")?.status, "warn");
  assert.equal(byId.get("executed-workbook")?.status, "warn");
  assert.equal(byId.get("whatsapp-profile")?.status, "error");
  assert.equal(byId.get("browser-runtime")?.status, "error");
  assert.equal(byId.get("liveperson-rest")?.status, "info");
  assert.match(byId.get("liveperson-rest")?.detail ?? "", /disabled/);
});

test("REST-only diagnostics check common readiness without browser, profile, Drive, or implicit auth probes", async (context) => {
  const projectRoot = await mkdtemp(path.join(tmpdir(), "pgn-operator-rest-doctor-"));
  context.after(() => rm(projectRoot, { recursive: true, force: true }));
  let schemas = 0;
  const report = await collectDiagnostics({
    ...restDiagnosticDependencies(projectRoot),
    environment: {
      ...REST_ENVIRONMENT,
      GOOGLE_DRIVE_EVIDENCE_ENABLED: "true",
      GOOGLE_DRIVE_EVIDENCE_PARENT_FOLDER: "abcdefghijklmno",
      GOOGLE_SERVICE_ACCOUNT_FILE: ".secrets/must-not-read.json",
    },
    checkDriveAccess: true,
    inspectWorkbookSchema: async () => { schemas += 1; return { ready: true, detail: "mocked common schema" }; },
  });
  assert.equal(report.ready, true);
  assert.equal(report.transport, "rest");
  assert.equal(schemas, 1);
  const checks = new Map(report.checks.map((check) => [check.id, check]));
  for (const id of ["playwright", "chromium", "whatsapp-profile", "whatsapp-target", "browser-runtime", "drive"]) {
    assert.equal(checks.get(id)?.status, "info", id);
    assert.match(checks.get(id)?.detail ?? "", /not required for REST/, id);
  }
  assert.equal(checks.get("configuration")?.status, "ok");
  assert.equal(checks.get("liveperson-rest")?.status, "ok");
  assert.match(checks.get("liveperson-rest")?.detail ?? "", /authentication not checked; no requests made/);
  for (const format of [formatDiagnosticReport, formatSetupInspection, formatSetupChecklist]) {
    const output = format(report);
    assert.doesNotMatch(output, /Chromium missing|WhatsApp profile missing|Drive access not verified|mock-rest|must-not-read/);
    assert.match(output, /LivePerson REST:/);
  }
});

for (const blocker of ["disabled", "secret", "schema", "configuration"] as const) {
  test(`REST-only readiness is blocked by ${blocker} without authentication`, async (context) => {
    const projectRoot = await mkdtemp(path.join(tmpdir(), "pgn-operator-rest-blocked-"));
    context.after(() => rm(projectRoot, { recursive: true, force: true }));
    const dependencies = restDiagnosticDependencies(projectRoot);
    if (blocker === "disabled") dependencies.environment!.LIVEPERSON_REST_ENABLED = "false";
    if (blocker === "secret") dependencies.environment!.LIVEPERSON_CLIENT_SECRET = "";
    if (blocker === "configuration") dependencies.environment!.REST_RESPONSE_TIMEOUT_MS = "invalid";
    if (blocker === "schema") dependencies.inspectWorkbookSchema = async () => ({ ready: false, detail: "missing User Input" });
    const report = await collectDiagnostics(dependencies);
    assert.equal(report.ready, false);
    assert(report.checks.some((check) => check.status === "error" && ["configuration", "liveperson-rest", "workbook-schema"].includes(check.id)));
  });
}

for (const fail of [false, true]) {
  test(`explicit REST diagnostic authentication ${fail ? "failure is redacted" : "success reports domains and auth only"}`, async (context) => {
    const projectRoot = await mkdtemp(path.join(tmpdir(), "pgn-operator-rest-check-"));
    context.after(() => rm(projectRoot, { recursive: true, force: true }));
    let calls = 0;
    const report = await collectDiagnostics({
      ...restDiagnosticDependencies(projectRoot), checkRestAccess: true,
      validateRest: async (config) => {
        calls += 1;
        if (fail) throw new Error(`Rejected ${config.livePersonRest?.clientSecret} ${config.livePersonRest?.clientId}`);
      },
    });
    assert.equal(calls, 1);
    assert.equal(report.ready, !fail);
    const rest = report.checks.find((check) => check.id === "liveperson-rest");
    assert.equal(rest?.status, fail ? "error" : "ok");
    assert.doesNotMatch(formatDiagnosticReport(report), /mock-rest/);
    if (!fail) assert.match(rest?.detail ?? "", /domains, app JWT, and synthetic consumer JWS validated; no conversations or testcase messages/);
  });
}

test("diagnostics identify missing Drive credentials without making a Drive call", async (context) => {
  const projectRoot = await mkdtemp(path.join(tmpdir(), "pgn-operator-drive-config-"));
  context.after(() => rm(projectRoot, { recursive: true, force: true }));
  let driveCalls = 0;
  const report = await collectDiagnostics({
    projectRoot,
    platform: "win32",
    environment: {
      GOOGLE_DRIVE_EVIDENCE_ENABLED: "true",
      GOOGLE_DRIVE_EVIDENCE_PARENT_FOLDER: "abcdefghijklmno",
    },
    npmVersion: async () => "11.0.0",
    packageVersion: async () => "1.0.0",
    chromiumExecutablePath: async () => "/fixture/chromium",
    pathExists: async () => true,
    validateDrive: async () => {
      driveCalls += 1;
    },
  });
  const drive = report.checks.find((check) => check.id === "drive");
  assert.equal(drive?.status, "error");
  assert.match(drive?.detail ?? "", /Configure GOOGLE_SERVICE_ACCOUNT/);
  assert.equal(driveCalls, 0);
});

test("diagnostics redact credential material from Drive failures", async (context) => {
  const projectRoot = await mkdtemp(path.join(tmpdir(), "pgn-operator-drive-error-"));
  context.after(() => rm(projectRoot, { recursive: true, force: true }));
  const credential = JSON.stringify({
    type: "service_account",
    client_email: "fixture@example.invalid",
    private_key: "fixture-private-key",
  });
  const report = await collectDiagnostics({
    projectRoot,
    platform: "darwin",
    environment: {
      GOOGLE_DRIVE_EVIDENCE_ENABLED: "true",
      GOOGLE_DRIVE_EVIDENCE_PARENT_FOLDER: "abcdefghijklmno",
      GOOGLE_SERVICE_ACCOUNT_JSON: credential,
    },
    npmVersion: async () => "11.0.0",
    packageVersion: async () => "1.0.0",
    chromiumExecutablePath: async () => "/fixture/chromium",
    pathExists: async () => true,
    validateDrive: async () => {
      throw new Error(`Drive rejected ${credential}`);
    },
  });
  const detail = report.checks.find((check) => check.id === "drive")?.detail ?? "";
  assert.match(detail, /\[REDACTED\]/);
  assert.doesNotMatch(detail, /fixture-private-key/);
});

test("diagnostics report Discord without inspecting it by default", async (context) => {
  const projectRoot = await mkdtemp(path.join(tmpdir(), "pgn-operator-discord-diagnostic-"));
  context.after(() => rm(projectRoot, { recursive: true, force: true }));
  const webhook = "https://discord.com/api/webhooks/123456789/diagnostic-token";
  let inspections = 0;
  const report = await collectDiagnostics({
    projectRoot,
    platform: "win32",
    environment: {
      PGN_WHATSAPP_PHONE: "628123456789",
      DISCORD_NOTIFICATIONS_ENABLED: "true",
      DISCORD_WEBHOOK_URL: webhook,
    },
    npmVersion: async () => "11.0.0",
    packageVersion: async () => "1.0.0",
    chromiumExecutablePath: async () => "/fixture/chromium",
    pathExists: async () => true,
    inspectDiscord: async () => {
      inspections += 1;
      throw new Error("Inspection must be explicitly enabled");
    },
  });

  const discord = report.checks.find((check) => check.id === "discord");
  assert.equal(inspections, 0);
  assert.equal(discord?.status, "ok");
  assert.match(discord?.detail ?? "", /connectivity not tested/);
  assert.equal((discord?.detail ?? "").includes(webhook), false);
});

test("diagnostics identify malformed Discord event settings", async (context) => {
  const projectRoot = await mkdtemp(path.join(tmpdir(), "pgn-operator-discord-invalid-"));
  context.after(() => rm(projectRoot, { recursive: true, force: true }));
  const report = await collectDiagnostics({
    projectRoot,
    platform: "win32",
    environment: {
      PGN_WHATSAPP_PHONE: "628123456789",
      DISCORD_NOTIFICATIONS_ENABLED: "true",
      DISCORD_WEBHOOK_URL:
        "https://discord.com/api/webhooks/123456789/diagnostic-token",
      DISCORD_NOTIFY_COMPLETE: "typo",
    },
    npmVersion: async () => "11.0.0",
    packageVersion: async () => "1.0.0",
    chromiumExecutablePath: async () => "/fixture/chromium",
    pathExists: async () => true,
  });

  const discord = report.checks.find((check) => check.id === "discord");
  assert.equal(discord?.status, "warn");
  assert.match(discord?.detail ?? "", /DISCORD_NOTIFY_COMPLETE/);
  assert.match(discord?.detail ?? "", /affected notifications are disabled/);
});

test("explicit Discord diagnostics inspect safely and remain non-blocking on failure", async (context) => {
  const projectRoot = await mkdtemp(path.join(tmpdir(), "pgn-operator-discord-inspect-"));
  context.after(() => rm(projectRoot, { recursive: true, force: true }));
  const webhook = "https://discord.com/api/webhooks/123456789/inspection-token";
  let inspections = 0;
  const report = await collectDiagnostics({
    projectRoot,
    platform: "win32",
    environment: {
      PGN_WHATSAPP_PHONE: "628123456789",
      DISCORD_NOTIFICATIONS_ENABLED: "true",
      DISCORD_WEBHOOK_URL: webhook,
    },
    npmVersion: async () => "11.0.0",
    packageVersion: async () => "1.0.0",
    chromiumExecutablePath: async () => "/fixture/chromium",
    pathExists: async () => true,
    checkDiscordAccess: true,
    inspectDiscord: async () => {
      inspections += 1;
      return {
        enabled: true,
        configured: true,
        valid: true,
        connectivity: "failed",
        testNotificationSent: false,
        reason: `Could not inspect ${webhook}`,
      };
    },
    inspectWorkbookSchema: async () => ({ ready: true, detail: "fixture mapping" }),
  });

  const discord = report.checks.find((check) => check.id === "discord");
  assert.equal(inspections, 1);
  assert.equal(discord?.status, "warn");
  assert.equal((discord?.detail ?? "").includes(webhook), false);
  assert.equal(report.ready, true);
});

for (const kind of ["recoverable", "running"] as const) {
  test(`diagnostics label ${kind} demos and skip enabled external checks`, async (context) => {
    const projectRoot = await mkdtemp(path.join(tmpdir(), "pgn-operator-demo-diagnostic-"));
    context.after(() => rm(projectRoot, { recursive: true, force: true }));
    const recovery = recoveryFixture(true);
    const credential = JSON.stringify({
      type: "service_account",
      client_email: "demo@example.invalid",
      private_key: "demo-private-key-must-not-be-displayed",
    });
    const webhook = "https://discord.com/api/webhooks/123456789/demo-secret-token";
    const externalCalls: string[] = [];
    const report = await collectDiagnostics({
      projectRoot,
      platform: "win32",
      environment: {
        ...REST_ENVIRONMENT,
        PGN_WHATSAPP_PHONE: "628123456789",
        GOOGLE_DRIVE_EVIDENCE_ENABLED: "true",
        GOOGLE_DRIVE_EVIDENCE_PARENT_FOLDER: "abcdefghijklmno",
        ...(kind === "recoverable"
          ? { GOOGLE_SERVICE_ACCOUNT_JSON: credential }
          : { GOOGLE_SERVICE_ACCOUNT_FILE: ".secrets/must-not-read.json" }),
        DISCORD_NOTIFICATIONS_ENABLED: "true",
        DISCORD_WEBHOOK_URL: webhook,
      },
      npmVersion: async () => "11.0.0",
      packageVersion: async () => "1.0.0",
      chromiumExecutablePath: async () => "/fixture/chromium",
      pathExists: async () => true,
      hasCommand: async () => false,
      inspectRecovery: async () => ({
        kind,
        state: recovery.validation.state,
        manifest: recovery.validation.manifest,
        lock: { status: "unlocked" },
      }),
      inspectWorkbookSchema: async () => ({ ready: true, detail: "fixture mapping" }),
      checkDriveAccess: true,
      checkRestAccess: true,
      validateRest: async () => { externalCalls.push("rest"); },
      validateDrive: async () => {
        externalCalls.push("drive");
      },
      checkDiscordAccess: true,
      inspectDiscord: async () => {
        externalCalls.push("discord");
        return {
          enabled: true,
          configured: true,
          valid: true,
          connectivity: "ok",
          testNotificationSent: false,
        };
      },
    });

    assert.deepEqual(externalCalls, []);
    assert.equal(report.ready, true);
    const byId = new Map(report.checks.map((check) => [check.id, check]));
    assert.match(byId.get("recovery")?.detail ?? "", /DEMO Run RECOVERY-OPERATOR-001/);
    for (const id of ["drive", "discord", "liveperson-rest"]) {
      assert.equal(byId.get(id)?.status, "info");
      assert.match(byId.get(id)?.detail ?? "", /DEMO: skipped/);
    }
    for (const format of [formatDiagnosticReport, formatSetupInspection, formatSetupChecklist]) {
      const output = format(report);
      assert.match(output, /Google Drive: DEMO: skipped/);
      assert.match(output, /Discord notifications: DEMO: skipped/);
      assert.doesNotMatch(output, /demo-private-key|demo-secret-token|must-not-read|628123456789/);
    }
  });
}

test("recoverable runs are surfaced before the main menu and never resume without confirmation", async () => {
  const recovery = recoveryFixture();
  let resumes = 0;
  const ui = new ScriptedUi(["resume", false, "menu", "exit"]);
  await runControlPanel(
    ui,
    stubActions({
      inspectRecovery: async () => recovery.discovery,
      validateRecovery: async () => recovery.validation,
      resumeRecovery: async () => {
        resumes += 1;
      },
      skipRecoveryScenario: async (runId) => ({ runId }),
      repairRecovery: async (runId) => ({ runId }),
      abandonRecovery: async (runId) => ({ runId }),
    }),
  );
  assert.equal(resumes, 0);
  assert.equal(ui.selectPrompts[0]?.message, "Recover interrupted Run RECOVERY-OPERATOR-001");
  assert.deepEqual(
    ui.selectPrompts[0]?.options.map((option) => option.label),
    [
      "Inspect recovery details",
      "Resume safely",
      "Skip current scenario and continue",
      "Abandon recovery",
      "Continue to main menu",
      "Exit",
    ],
  );
  const output = ui.events.join("\n");
  assert.match(output, /Interrupted scenario: TC-002/);
  assert.match(output, /Recovery execution cancelled/);
});

test("confirmed recovery keeps the Run ID and starts at the interrupted scenario", async () => {
  const recovery = recoveryFixture();
  let available = true;
  const resumed: Array<{ runId: string; acceptSourceDrift?: boolean }> = [];
  const ui = new ScriptedUi(["resume", true, "exit"]);
  await runControlPanel(
    ui,
    stubActions({
      inspectRecovery: async (): Promise<RecoveryDiscovery> =>
        available
          ? recovery.discovery
          : { kind: "none", lock: { status: "unlocked" } },
      validateRecovery: async () => recovery.validation,
      resumeRecovery: async (runId, acceptSourceDrift) => {
        resumed.push({ runId, acceptSourceDrift });
        available = false;
      },
      skipRecoveryScenario: async (runId) => ({ runId }),
      repairRecovery: async (runId) => ({ runId }),
      abandonRecovery: async (runId) => ({ runId }),
    }),
  );
  assert.deepEqual(resumed, [
    { runId: "RECOVERY-OPERATOR-001", acceptSourceDrift: false },
  ]);
  assert(
    ui.confirmPrompts.some((prompt) =>
      prompt.message.includes("at TC-002 from Turn 1"),
    ),
  );
});

test("continuous recovery offers only full restart or abandonment and guards stale dispatch", async () => {
  const recovery = recoveryFixture(false, "continuous");
  const before = structuredClone(recovery);
  const calls: string[] = [];
  const ui = new ScriptedUi(["inspect", "resume", "skip", "repair", "restart", false, "menu", "exit"]);
  await runControlPanel(ui, stubRecoveryActions(recovery, calls));

  assert.deepEqual(calls, ["validate:RECOVERY-OPERATOR-001", "validate:RECOVERY-OPERATOR-001"]);
  assert.deepEqual(recovery, before);
  assert.equal(ui.selectPrompts[0]?.initialValue, "inspect");
  assert.deepEqual(ui.selectPrompts[0]?.options.map((option) => option.label), [
    "Inspect details",
    "Restart continuous run from beginning",
    "Abandon",
    "Main menu",
    "Exit",
  ]);
  assert(ui.events.includes(`warn:${CONTINUOUS_RECOVERY_WARNING}`));
  assert(ui.events.includes(`warn:${CONTINUOUS_SESSION_WARNING}`));
  assert.equal(ui.confirmPrompts.length, 1);
  assert.equal(ui.confirmPrompts[0]?.initialValue, false);
  assert.match(ui.events.join("\n"), /Full continuous restart cancelled/);
  assert.equal(ui.selectPrompts.at(-1)?.message, "Choose an operation");
});

for (const mode of ["full", "retest"] as const) {
  for (const sourceDrift of ["unchanged", "formatting-only"] as const) {
    test(`${mode} continuous restart validates ${sourceDrift} source and explicitly confirms the full original run`, async () => {
      const recovery = recoveryFixture(false, "continuous");
      recovery.validation.mode = mode;
      recovery.validation.state.mode = mode;
      recovery.validation.sourceDrift = sourceDrift;
      // A full restart must not offer partial progress repair.
      recovery.validation.reconciliation!.mismatchedScenarioIds = ["TC-001"];
      const before = structuredClone(recovery);
      const calls: string[] = [];
      const ui = new ScriptedUi(["restart", ...(sourceDrift === "formatting-only" ? [true] : []), true, "exit"]);
      await runControlPanel(ui, stubRecoveryActions(recovery, calls));

      assert.deepEqual(calls, [
        "validate:RECOVERY-OPERATOR-001",
        `restart:RECOVERY-OPERATOR-001:${sourceDrift === "formatting-only"}`,
      ]);
      assert.deepEqual(recovery, before);
      assert(ui.confirmPrompts.every((prompt) => prompt.initialValue === false));
      assert.equal(ui.confirmPrompts.length, sourceDrift === "formatting-only" ? 2 : 1);
      const confirmation = ui.confirmPrompts.at(-1)?.message ?? "";
      assert.match(confirmation, /ALL 2 originally selected scenarios/);
      assert.match(confirmation, /including completed\/skipped scenarios, in their original order/);
      assert.match(confirmation, /NEW Run ID and new evidence folder/);
      assert.match(confirmation, /ABANDONED only after the new checkpoint is saved/);
      assert(ui.events.includes(`warn:${CONTINUOUS_SESSION_WARNING}`));
    });
  }

  test(`${mode} continuous demo restart is preview only even when validation loses demo flags`, async () => {
    const recovery = recoveryFixture(true, "continuous");
    recovery.validation.mode = mode;
    recovery.validation.state.mode = mode;
    const before = structuredClone(recovery);
    const calls: string[] = [];
    const ui = new ScriptedUi(["restart", "resume", "skip", "repair", "exit"]);
    await runControlPanel(ui, stubRecoveryActions(recovery, calls, {
      validateRecovery: async () => {
        calls.push("validate");
        const validation = structuredClone(recovery.validation);
        delete validation.state.isDemo;
        delete validation.manifest.isDemo;
        return validation;
      },
    }));
    assert.deepEqual(calls, ["validate"]);
    assert.deepEqual(recovery, before);
    assert.equal(ui.confirmPrompts.length, 0);
    assert.match(ui.events.join("\n"), /DEMO continuous restart preview/);
    assert.match(ui.events.join("\n"), /all 2 originally selected scenarios from the beginning \(preview only\)/);
    assert.match(ui.events.join("\n"), /Original order: TC-001, TC-002/);
  });
}

test("continuous restart recognizes demo flags found only during validation", async () => {
  const recovery = recoveryFixture(false, "continuous");
  const validation = structuredClone(recovery.validation);
  validation.manifest.isDemo = true;
  const calls: string[] = [];
  const ui = new ScriptedUi(["restart", "exit"]);
  await runControlPanel(ui, stubRecoveryActions(recovery, calls, {
    validateRecovery: async () => validation,
  }));
  assert.deepEqual(calls, []);
  assert.equal(ui.confirmPrompts.length, 0);
  assert.match(ui.events.join("\n"), /DEMO continuous restart preview/);
});

test("continuous restart requires restartReady, safe source drift, and an available adapter, never ready alone", async () => {
  for (const blockedBy of ["false", "missing", "structural", "unavailable", "adapter"] as const) {
    const recovery = recoveryFixture(false, "continuous");
    recovery.validation.ready = true;
    if (blockedBy === "false") recovery.validation.restartReady = false;
    if (blockedBy === "missing") delete recovery.validation.restartReady;
    if (blockedBy === "structural" || blockedBy === "unavailable") {
      recovery.validation.sourceDrift = blockedBy;
    }
    const calls: string[] = [];
    const ui = new ScriptedUi(["restart", "exit"]);
    await runControlPanel(ui, stubRecoveryActions(recovery, calls,
      blockedBy === "adapter" ? { restartRecovery: undefined } : {},
    ));
    assert.deepEqual(calls, ["validate:RECOVERY-OPERATOR-001"], blockedBy);
    assert.equal(ui.confirmPrompts.length, 0, blockedBy);
    assert.match(ui.events.join("\n"), /No testcase was executed/);
    assert.match(ui.selectPrompts[0]?.message ?? "", /Recover interrupted Run/);
  }
});

test("declining or cancelling continuous restart confirmations never executes", async () => {
  for (const answer of [false, undefined]) {
    for (const sourceDrift of ["unchanged", "formatting-only"] as const) {
      const recovery = recoveryFixture(false, "continuous");
      recovery.validation.sourceDrift = sourceDrift;
      const calls: string[] = [];
      const ui = new ScriptedUi(["restart", answer, "exit"]);
      await runControlPanel(ui, stubRecoveryActions(recovery, calls));
      assert.deepEqual(calls, ["validate:RECOVERY-OPERATOR-001"]);
      assert.equal(ui.confirmPrompts.length, 1);
      assert.equal(ui.confirmPrompts[0]?.initialValue, false);
    }
  }
});

test("continuous recovery can be explicitly abandoned without resuming or restarting", async () => {
  const recovery = recoveryFixture(false, "continuous");
  let available = true;
  const calls: string[] = [];
  const ui = new ScriptedUi(["abandon", false, "abandon", true, "exit"]);
  await runControlPanel(ui, stubRecoveryActions(recovery, calls, {
    inspectRecovery: async () => available ? recovery.discovery : { kind: "none", lock: { status: "unlocked" } },
    abandonRecovery: async (runId) => {
      calls.push(`abandon:${runId}`);
      available = false;
      return { runId };
    },
  }));
  assert.deepEqual(calls, ["abandon:RECOVERY-OPERATOR-001"]);
  assert(ui.confirmPrompts.every((prompt) => prompt.initialValue === false));
  assert.match(ui.events.join("\n"), /marked ABANDONED; artifacts were preserved/);
  assert.equal(ui.selectPrompts.at(-1)?.message, "Choose an operation");
});

test("stale isolated recovery menus cannot resume or repair continuous validation", async () => {
  for (const source of ["state", "manifest"] as const) {
    const recovery = recoveryFixture();
    const validation = structuredClone(recovery.validation);
    validation[source].sessionMode = "continuous";
    validation.ready = true;
    validation.reconciliation!.mismatchedScenarioIds = ["TC-001"];
    const calls: string[] = [];
    const ui = new ScriptedUi(["resume", "exit"]);
    await runControlPanel(ui, stubRecoveryActions(recovery, calls, {
      validateRecovery: async () => { calls.push("validate"); return validation; },
    }));
    assert.deepEqual(calls, ["validate"]);
    assert.equal(ui.confirmPrompts.length, 0);
    assert(ui.events.includes(`warn:${CONTINUOUS_RECOVERY_WARNING}`));
  }
});

test("revalidated continuous recovery cannot pass through the isolated resume path", async () => {
  const recovery = recoveryFixture();
  recovery.validation.reconciliation!.mismatchedScenarioIds = ["TC-001"];
  const continuous = recoveryFixture(false, "continuous").validation;
  continuous.ready = true;
  let validations = 0;
  const calls: string[] = [];
  const ui = new ScriptedUi(["resume", "rerun", true, "exit"]);
  await runControlPanel(ui, stubRecoveryActions(recovery, calls, {
    validateRecovery: async () => {
      calls.push("validate");
      return ++validations === 1 ? recovery.validation : continuous;
    },
  }));
  assert.deepEqual(calls, ["validate", "repair:RECOVERY-OPERATOR-001", "validate"]);
  assert.equal(ui.confirmPrompts.length, 1);
  assert(ui.events.includes(`warn:${CONTINUOUS_RECOVERY_WARNING}`));
});

test("legacy isolated recovery still allows confirmed skip followed by resume", async () => {
  const recovery = recoveryFixture();
  const calls: string[] = [];
  const ui = new ScriptedUi(["skip", true, true, "exit"]);
  await runControlPanel(ui, stubRecoveryActions(recovery, calls, { restartRecovery: undefined }));
  assert.deepEqual(calls, [
    "skip:RECOVERY-OPERATOR-001",
    "validate:RECOVERY-OPERATOR-001",
    "resume:RECOVERY-OPERATOR-001",
  ]);
  assert.equal(ui.confirmPrompts.length, 2);
  assert(ui.confirmPrompts.every((prompt) => prompt.initialValue === false));
});

for (const mode of ["full", "retest"] as const) {
  test(`${mode} demo resume and restart use the same recovery menu without execution or mutation`, async () => {
    const recovery = recoveryFixture(true);
    recovery.validation.mode = mode;
    recovery.validation.state.mode = mode;
    recovery.validation.sourceDrift = "formatting-only";
    recovery.validation.checks = [
      {
        id: "source",
        label: "Source workbook",
        status: "warn",
        detail: "file hash changed, but selected scenario inputs are unchanged",
      },
      {
        id: "reconciliation",
        label: "Progress reconciliation",
        status: "ok",
        detail: "fixture workbook and transcript agree with checkpoint",
      },
    ];
    const before = structuredClone(recovery);
    const calls: string[] = [];
    const execute = async (): Promise<void> => { calls.push("execute"); };
    const mutate = async (runId: string) => {
      calls.push("mutate");
      return { runId };
    };
    const ui = new ScriptedUi(["inspect", "resume", "restart", "menu", "exit"]);
    await runControlPanel(
      ui,
      stubActions({
        inspectRecovery: async () => recovery.discovery,
        validateRecovery: async (runId) => {
          calls.push(`validate:${runId}`);
          const validation = structuredClone(recovery.validation);
          // Discovery must keep the demo safe even if validation loses both flags.
          delete validation.state.isDemo;
          delete validation.manifest.isDemo;
          return validation;
        },
        resumeRecovery: execute,
        runPgn: execute,
        runRetest: execute,
        skipRecoveryScenario: mutate,
        repairRecovery: mutate,
        abandonRecovery: mutate,
      }),
    );

    assert.deepEqual(calls, Array(3).fill("validate:RECOVERY-OPERATOR-001"));
    assert.deepEqual(recovery, before);
    assert.equal(ui.confirmPrompts.length, 0);
    assert.equal(ui.selectPrompts.length, 5);
    assert(
      ui.selectPrompts.slice(0, -1).every(
        (prompt) => prompt.message === "Recover interrupted Run RECOVERY-OPERATOR-001 [DEMO]",
      ),
    );
    assert.deepEqual(
      ui.selectPrompts[0]?.options.map((option) => option.label),
      [
        "Inspect recovery details",
        "Resume safely",
        "Restart interrupted scenario",
        "Skip current scenario and continue",
        "Abandon recovery",
        "Continue to main menu",
        "Exit",
      ],
    );
    assert.equal(ui.selectPrompts.at(-1)?.message, "Choose an operation");
    const output = ui.events.join("\n");
    assert.match(output, /Recoverable run found \[DEMO\]/);
    assert.match(output, /Source workbook: file hash changed, but selected scenario inputs are unchanged/);
    assert.match(output, /Progress reconciliation: fixture workbook and transcript agree with checkpoint/);
    assert.equal(output.match(/note:DEMO recovery preview:/g)?.length, 2);
    assert.match(output, /1 completed, 0 skipped, 1 remaining/);
    assert.match(output, /Previous interruption: Interrupted during Turn 2 of 2; demo only\./);
    assert.match(output, /Restart interrupted scenario: TC-002 from Turn 1 of 2 \(preview only\)/);
    assert.match(output, /Next scenario: TC-002 from Turn 1 \(preview only\)/);
    assert.match(output, /DEMO: UI\/testing only; no live execution/);
    assert.doesNotMatch(output, /This will open WhatsApp|How should recovery reconcile/);
  });
}

test("declining or cancelling a demo skip does not mutate or validate recovery", async () => {
  const recovery = recoveryFixture(true);
  const before = structuredClone(recovery);
  const calls: string[] = [];
  const mutate = async (runId: string) => {
    calls.push("mutate");
    return { runId };
  };
  const ui = new ScriptedUi(["skip", false, "skip", undefined, "exit"]);
  await runControlPanel(
    ui,
    stubActions({
      inspectRecovery: async () => recovery.discovery,
      validateRecovery: async () => {
        calls.push("validate");
        return recovery.validation;
      },
      resumeRecovery: async () => { calls.push("execute"); },
      skipRecoveryScenario: mutate,
      repairRecovery: mutate,
      abandonRecovery: mutate,
    }),
  );
  assert.deepEqual(calls, []);
  assert.deepEqual(recovery, before);
  assert.equal(ui.confirmPrompts.length, 2);
  assert(ui.confirmPrompts.every((prompt) => prompt.initialValue === false));
  assert.equal(
    ui.confirmPrompts[0]?.message,
    "Explicitly skip TC-002 in Run RECOVERY-OPERATOR-001? Existing evidence will be preserved.",
  );
});

test("confirmed demo skip calls the existing action before preview without resuming or final cleanup", async () => {
  const recovery = recoveryFixture(true);
  const calls: string[] = [];
  const execute = async (): Promise<void> => { calls.push("execute"); };
  const ui = new ScriptedUi(["skip", true, "exit"]);
  await runControlPanel(
    ui,
    stubActions({
      inspectRecovery: async () => recovery.discovery,
      skipRecoveryScenario: async (runId) => {
        calls.push(`skip:${runId}`);
        assert.equal(ui.confirmPrompts.length, 1);
        const state = recovery.validation.state;
        state.skippedScenarioIds.push("TC-002");
        state.activeScenarioId = undefined;
        state.activeScenarioAttempt = undefined;
        state.activeScenarioStartedAt = undefined;
        state.status = "RECOVERABLE";
        state.interruptionReason = "Scenario TC-002 skipped by operator";
        const reconciliation = recovery.validation.reconciliation!;
        reconciliation.nextScenarioId = undefined;
        reconciliation.interruptedScenarioId = undefined;
        reconciliation.restartInterruptedScenarioFromTurnOne = false;
        return { runId, scenarioId: "TC-002", warning: "fixture skip audit warning" };
      },
      validateRecovery: async (runId) => {
        calls.push(`validate:${runId}`);
        const validation = structuredClone(recovery.validation);
        delete validation.state.isDemo;
        delete validation.manifest.isDemo;
        return validation;
      },
      resumeRecovery: execute,
      runPgn: execute,
      runRetest: execute,
      repairRecovery: async (runId) => {
        calls.push("repair");
        return { runId };
      },
      abandonRecovery: async (runId) => {
        calls.push("abandon");
        return { runId };
      },
    }),
  );
  assert.deepEqual(calls, ["skip:RECOVERY-OPERATOR-001", "validate:RECOVERY-OPERATOR-001"]);
  assert.equal(ui.confirmPrompts.length, 1);
  assert.equal(recovery.validation.state.finalCleanupComplete, false);
  assert.deepEqual(recovery.validation.state.completedScenarioIds, ["TC-001"]);
  assert.deepEqual(recovery.validation.state.skippedScenarioIds, ["TC-002"]);
  assert.equal(ui.selectPrompts[1]?.message, ui.selectPrompts[0]?.message);
  const output = ui.events.join("\n");
  assert.match(output, /fixture skip audit warning/);
  assert.match(output, /1 completed, 1 skipped, 0 remaining/);
  assert.match(output, /Next scenario: none; no remaining scenarios/);
  assert.match(output, /note:DEMO recovery preview:/);
  assert.doesNotMatch(output, /This will open WhatsApp/);
});

test("demo abandonment uses the existing confirmed action and preserves artifacts", async () => {
  const recovery = recoveryFixture(true);
  let available = true;
  const calls: string[] = [];
  const ui = new ScriptedUi(["abandon", false, "abandon", true, "exit"]);
  await runControlPanel(
    ui,
    stubActions({
      inspectRecovery: async () => available
        ? recovery.discovery
        : { kind: "none", lock: { status: "unlocked" } },
      validateRecovery: async () => {
        calls.push("validate");
        return recovery.validation;
      },
      resumeRecovery: async () => { calls.push("execute"); },
      skipRecoveryScenario: async (runId) => {
        calls.push("skip");
        return { runId };
      },
      repairRecovery: async (runId) => {
        calls.push("repair");
        return { runId };
      },
      abandonRecovery: async (runId) => {
        calls.push(`abandon:${runId}`);
        assert.equal(ui.confirmPrompts.length, 2);
        available = false;
        return { runId };
      },
    }),
  );
  assert.deepEqual(calls, ["abandon:RECOVERY-OPERATOR-001"]);
  assert.equal(ui.confirmPrompts.length, 2);
  assert(ui.confirmPrompts.every((prompt) => prompt.initialValue === false));
  assert.match(ui.confirmPrompts[0]?.message ?? "", /keeps the workbook, transcript, evidence, and recovery history/);
  assert.match(ui.events.join("\n"), /marked ABANDONED; artifacts were preserved/);
  assert.equal(ui.selectPrompts.at(-1)?.message, "Choose an operation");
});

test("unknown recovery choices never fall through to abandonment in live or demo menus", async () => {
  for (const isDemo of [false, true]) {
    const recovery = recoveryFixture(isDemo);
    const calls: string[] = [];
    const mutate = async (runId: string) => {
      calls.push("mutate");
      return { runId };
    };
    const ui = new ScriptedUi([isDemo ? "unknown" : "restart", "exit"]);
    await runControlPanel(
      ui,
      stubActions({
        inspectRecovery: async () => recovery.discovery,
        validateRecovery: async () => {
          calls.push("validate");
          return recovery.validation;
        },
        resumeRecovery: async () => { calls.push("execute"); },
        skipRecoveryScenario: mutate,
        repairRecovery: mutate,
        abandonRecovery: mutate,
      }),
    );
    assert.deepEqual(calls, []);
    assert.equal(ui.confirmPrompts.length, 0);
    assert.equal(ui.selectPrompts.length, 2);
  }
});

test("demo mismatch validation remains visible without repair, mutation, or execution", async () => {
  const recovery = recoveryFixture(true);
  recovery.validation.ready = false;
  recovery.validation.reconciliation!.mismatchedScenarioIds = ["TC-001"];
  recovery.validation.reconciliation!.safeCompletedIds = [];
  recovery.validation.checks = [{
    id: "reconciliation",
    label: "Progress reconciliation",
    status: "error",
    detail: "TC-001 disagrees across checkpoint, workbook, and transcript",
  }];
  const before = structuredClone(recovery);
  const calls: string[] = [];
  const mutate = async (runId: string) => {
    calls.push("mutate");
    return { runId };
  };
  const ui = new ScriptedUi(["resume", "restart", "exit"]);
  await runControlPanel(
    ui,
    stubActions({
      inspectRecovery: async () => recovery.discovery,
      validateRecovery: async () => {
        calls.push("validate");
        return recovery.validation;
      },
      resumeRecovery: async () => { calls.push("execute"); },
      skipRecoveryScenario: mutate,
      repairRecovery: mutate,
      abandonRecovery: mutate,
    }),
  );
  assert.deepEqual(calls, ["validate", "validate"]);
  assert.deepEqual(recovery, before);
  assert.equal(ui.confirmPrompts.length, 0);
  assert.equal(ui.selectPrompts.length, 3);
  const output = ui.events.join("\n");
  assert.match(output, /ERROR Progress reconciliation: TC-001 disagrees/);
  assert.match(output, /DEMO validation remains BLOCKED; no repair or execution was attempted/);
  assert.match(output, /Next scenario: TC-002 from Turn 1 \(preview only\)/);
  assert.doesNotMatch(output, /How should recovery reconcile/);
});

test("fresh-run cancellation never invokes workbook preparation", async () => {
  let preparations = 0;
  const ui = new ScriptedUi(["run", "fresh", false, "back", "exit"]);
  await runControlPanel(
    ui,
    stubActions({
      prepareFresh: async () => {
        preparations += 1;
      },
    }),
  );
  assert.equal(preparations, 0);
  assert.match(ui.events.join("\n"), /Fresh-run preparation cancelled/);
});

test("full execution, evidence migration, and auth recreation require confirmation", async () => {
  let fullRuns = 0;
  let migrations = 0;
  let recreations = 0;
  const ui = new ScriptedUi([
    "run",
    "remaining",
    "isolated",
    false,
    "back",
    "evidence",
    "migrate",
    false,
    "back",
    "whatsapp",
    "recreate",
    false,
    "back",
    "exit",
  ]);
  await runControlPanel(
    ui,
    stubActions({
      runPgn: async () => {
        fullRuns += 1;
      },
      migrateEvidence: async () => {
        migrations += 1;
      },
      recreateWhatsApp: async () => {
        recreations += 1;
      },
    }),
  );
  assert.deepEqual([fullRuns, migrations, recreations], [0, 0, 0]);
});

for (const sessionMode of ["isolated", "continuous"] as const) {
  for (const selection of ["full", "ids", "sheet", "retest"] as const) {
    test(`REST ${selection} routes ${sessionMode} through shared actions with no implicit rerun`, async () => {
      const filters = selection === "ids" ? ["--test", "TC-001,TC-002", "--rerun"]
        : selection === "sheet" ? ["--sheet", "negative"] : [];
      const expected = ["--transport=rest", ...filters, ...(sessionMode === "continuous" ? ["--session=continuous"] : [])];
      const received: Array<{ action: string; args: string[] }> = [];
      const ui = new ScriptedUi([
        "run", "rest", sessionMode, selection,
        ...(selection === "ids" ? ["TC-001, TC-002", true] : selection === "sheet" ? ["negative"] : selection === "retest" ? ["ready"] : []),
        true, "back", "exit",
      ]);
      await runControlPanel(ui, stubActions({
        prepareFresh: async () => assert.fail("REST full must not implicitly prepare a fresh workbook"),
        runPgn: async (args) => { received.push({ action: "full", args }); },
        validateRetest: async (args) => {
          received.push({ action: "validate-retest", args: [...args] });
          return { selectedCount: 2, shouldExecute: true, readyToExecute: true, finalCleanupOnly: false };
        },
        runRetest: async (args) => { received.push({ action: "retest", args }); },
      }));
      assert.deepEqual(received, selection === "retest"
        ? [{ action: "validate-retest", args: expected }, { action: "retest", args: expected }]
        : [{ action: "full", args: expected }]);
      const runMenu = ui.selectPrompts.find((prompt) => prompt.message === "Run tests");
      assert(runMenu?.options.some((option) => option.label === "REST Bulk Test" && option.value === "rest"));
      for (const value of ["validate", "remaining", "sheet", "ids", "fresh", "back"]) {
        assert(runMenu?.options.some((option) => option.value === value), `Existing WhatsApp choice ${value}`);
      }
      const selector = ui.selectPrompts.find((prompt) => prompt.message === "Session Mode for this REST run");
      assert.equal(selector?.initialValue, "isolated");
      assert.match(selector?.options[0]?.hint ?? "", /Fresh conversation per scenario/);
      assert.match(selector?.options[1]?.hint ?? "", /One initial conversation/);
      assert.deepEqual(ui.selectPrompts.find((prompt) => prompt.message === "REST Bulk Test")?.options.map((option) => option.value), ["full", "ids", "sheet", "retest", "validate", "back"]);
      const confirmation = ui.confirmPrompts.at(-1);
      assert.equal(confirmation?.initialValue, false);
      assert.match(confirmation?.message ?? "", /real testcase messages through LivePerson REST/);
      assert.match(confirmation?.message ?? "", /No WhatsApp, debug reset, screenshots, or Drive uploads/);
      assert.doesNotMatch(confirmation?.message ?? "", /will open WhatsApp|send reset|upload evidence/);
      assert.equal(ui.events.includes(`warn:${continuousSessionWarning("rest")}`), sessionMode === "continuous");
      assert.equal(ui.events.includes(`warn:${CONTINUOUS_SESSION_WARNING}`), false);
    });
  }
}

test("REST transport and continuous session do not become defaults for later runs", async () => {
  const received: string[][] = [];
  const ui = new ScriptedUi([
    "run", "rest", "continuous", "full", true,
    "rest", "isolated", "full", true,
    "remaining", "isolated", true, "back", "exit",
  ]);
  await runControlPanel(ui, stubActions({ runPgn: async (args) => { received.push(args); } }));
  assert.deepEqual(received, [["--transport=rest", "--session=continuous"], ["--transport=rest"], []]);
  assert(ui.selectPrompts.filter((prompt) => prompt.message.startsWith("Session Mode")).every((prompt) => prompt.initialValue === "isolated"));
});

test("REST validation requires explicit confirmation and never runs testcases", async () => {
  const received: string[][] = [];
  const ui = new ScriptedUi([
    "run", "rest", "continuous", "validate", false, "validate", true, "back", "back", "exit",
  ]);
  await runControlPanel(ui, stubActions({
    validateRest: async (args) => { received.push(args); return { ready: true, selectedCount: 7 }; },
    runPgn: async () => assert.fail("Validation must not execute"),
    runRetest: async () => assert.fail("Validation must not retest"),
    validatePgn: async () => assert.fail("REST validation uses its own readiness action"),
  }));
  assert.deepEqual(received, [["--transport=rest", "--session=continuous"]]);
  assert(ui.confirmPrompts.every((prompt) => prompt.initialValue === false));
  assert.match(ui.confirmPrompts[0]?.message ?? "", /real LivePerson domain and authentication requests/);
  assert.match(ui.events.join("\n"), /7 scenario\(s\) selected; no testcase was executed/);
});

test("REST validation without an adapter is disabled and never falls back to execution", async () => {
  const ui = new ScriptedUi(["run", "rest", "isolated", "validate", "back", "back", "exit"]);
  await runControlPanel(ui, stubActions({ runPgn: async () => assert.fail("Unavailable validation must not run") }));
  assert.equal(ui.selectPrompts.find((prompt) => prompt.message === "REST Bulk Test")?.options.find((option) => option.value === "validate")?.disabled, true);
  assert.equal(ui.confirmPrompts.length, 0);
  assert.match(ui.events.join("\n"), /REST validation is unavailable/);
});

test("REST back, selector cancellation, and declined execution never invoke a runner", async () => {
  for (const responses of [
    ["run", "rest", undefined],
    ["run", "rest", "isolated", "back", "back", "exit"],
    ["run", "rest", "isolated", "retest", "back", "back", "back", "exit"],
    ["run", "rest", "continuous", "full", false, "back", "back", "exit"],
    ["run", "rest", "continuous", "full", undefined],
  ]) {
    await runControlPanel(new ScriptedUi(responses), stubActions({
      runPgn: async () => assert.fail("Cancelled execution must not run"),
      runRetest: async () => assert.fail("Cancelled execution must not retest"),
    }));
  }
});

test("REST retest review and selected IDs preserve transport and session arguments", async () => {
  const received: string[][] = [];
  const ui = new ScriptedUi(["run", "rest", "continuous", "retest", "review", "ids", "TC-002", true, "back", "exit"]);
  await runControlPanel(ui, stubActions({
    validateRetest: async (args) => {
      received.push(args);
      return { selectedCount: 1, shouldExecute: true, readyToExecute: true, finalCleanupOnly: false };
    },
    runRetest: async (args) => { received.push(args); },
  }));
  assert.deepEqual(received, [
    ["--transport=rest", "--session=continuous"],
    ["--transport=rest", "--test", "TC-002", "--session=continuous"],
    ["--transport=rest", "--test", "TC-002", "--session=continuous"],
  ]);
  assert.equal(ui.selectPrompts.find((prompt) => prompt.message === "REST retest fixed cases")?.options.some((option) => option.value === "resume"), false);
});

for (const mode of ["full", "retest"] as const) {
  test(`isolated REST ${mode} recovery confirms a fresh conversation from Turn 1 without Drive promises`, async () => {
    const recovery = recoveryFixture(false, "isolated", "rest");
    recovery.validation.state.mode = mode;
    recovery.validation.mode = mode;
    const calls: string[] = [];
    const ui = new ScriptedUi(["resume", true, "exit"]);
    await runControlPanel(ui, stubRecoveryActions(recovery, calls));
    assert.deepEqual(calls, ["validate:RECOVERY-OPERATOR-001", "resume:RECOVERY-OPERATOR-001"]);
    const confirmation = ui.confirmPrompts.at(-1)?.message ?? "";
    assert.match(confirmation, /fresh conversation per scenario.*Turn 1/);
    assert.match(confirmation, /No WhatsApp, debug reset, screenshots, or Drive uploads/);
    assert.doesNotMatch(confirmation, /will open WhatsApp|reuse existing Drive artifacts/);
  });

  test(`continuous REST ${mode} recovery allows only full restart or abandonment`, async () => {
    const recovery = recoveryFixture(false, "continuous", "rest");
    recovery.validation.state.mode = mode;
    recovery.validation.mode = mode;
    const calls: string[] = [];
    const ui = new ScriptedUi(["resume", "skip", "repair", "restart", true, "abandon", true, "exit"]);
    await runControlPanel(ui, stubRecoveryActions(recovery, calls));
    assert.deepEqual(calls, ["validate:RECOVERY-OPERATOR-001", "restart:RECOVERY-OPERATOR-001:false", "abandon:RECOVERY-OPERATOR-001"]);
    const confirmation = ui.confirmPrompts[0]?.message ?? "";
    assert.match(confirmation, /ALL 2 originally selected scenarios/);
    assert.match(confirmation, /NEW Run ID and one new conversation/);
    assert.match(confirmation, /ABANDONED only after the new checkpoint is saved/);
    assert.doesNotMatch(confirmation, /opens WhatsApp|new evidence folder/);
    assert(ui.events.includes(`warn:${continuousSessionWarning("rest")}`));
    assert.equal(ui.events.includes(`warn:${CONTINUOUS_SESSION_WARNING}`), false);
    assert(ui.confirmPrompts.every((prompt) => prompt.initialValue === false));
  });
}

for (const sessionMode of ["isolated", "continuous"] as const) {
  for (const selection of ["remaining", "sheet", "ids", "setup"] as const) {
    test(`${selection} execution selects ${sessionMode} per run without changing isolated arguments`, async () => {
      const filterAnswers = selection === "sheet" ? ["kb"] : selection === "ids" ? ["TC-001, TC-002", true] : [];
      const args = selection === "sheet" ? ["--sheet", "kb"] : selection === "ids" ? ["--test", "TC-001,TC-002", "--rerun"] : [];
      if (sessionMode === "continuous") args.push("--session=continuous");
      const received: string[][] = [];
      const ui = new ScriptedUi([
        ...(selection === "setup" ? ["setup"] : ["run", selection, ...filterAnswers]),
        sessionMode, true,
        ...(selection === "setup" ? [] : ["back"]),
        "exit",
      ]);
      await runControlPanel(ui, stubActions({
        setup: async () => "full-test",
        runPgn: async (receivedArgs) => {
          assert(ui.confirmPrompts.at(-1)?.message.includes("will open WhatsApp"));
          if (sessionMode === "continuous") assert(ui.events.includes(`warn:${CONTINUOUS_SESSION_WARNING}`));
          received.push(receivedArgs);
        },
      }));
      assert.deepEqual(received, [args]);
      const selector = ui.selectPrompts.find((prompt) => prompt.message === "Session Mode for this run");
      assert.equal(selector?.initialValue, "isolated");
      assert.deepEqual(selector?.options.map((option) => option.value), ["isolated", "continuous"]);
      assert.equal(ui.confirmPrompts.at(-1)?.initialValue, false);
      assert.equal(ui.events.includes(`warn:${CONTINUOUS_SESSION_WARNING}`), sessionMode === "continuous");
    });
  }

  for (const selection of ["ready", "ids"] as const) {
    test(`${selection} retest validates and executes the same ${sessionMode} arguments`, async () => {
      const received: Array<{ action: string; args: string[] }> = [];
      const args = selection === "ids" ? ["--test", "TC-001,TC-002"] : [];
      if (sessionMode === "continuous") args.push("--session=continuous");
      const ui = new ScriptedUi([
        "retest", selection, ...(selection === "ids" ? ["TC-001, TC-002"] : []),
        sessionMode, true, "back", "exit",
      ]);
      await runControlPanel(ui, stubActions({
        validateRetest: async (args) => {
          received.push({ action: "validate", args: [...args] });
          return { selectedCount: 2, finalCleanupOnly: false, shouldExecute: true, readyToExecute: true };
        },
        runRetest: async (args) => { received.push({ action: "execute", args: [...args] }); },
      }));
      assert.deepEqual(received, [{ action: "validate", args }, { action: "execute", args }]);
      assert.equal(ui.selectPrompts.find((prompt) => prompt.message === "Session Mode for this run")?.initialValue, "isolated");
      assert.equal(ui.confirmPrompts.at(-1)?.initialValue, false);
      assert.equal(ui.events.includes(`warn:${CONTINUOUS_SESSION_WARNING}`), sessionMode === "continuous");
    });
  }
}

test("session choice resets to isolated for the next run instead of becoming a setting", async () => {
  const received: string[][] = [];
  const ui = new ScriptedUi(["run", "remaining", "continuous", true, "remaining", "isolated", true, "back", "exit"]);
  await runControlPanel(ui, stubActions({ runPgn: async (args) => { received.push(args); } }));
  assert.deepEqual(received, [["--session=continuous"], []]);
  const selectors = ui.selectPrompts.filter((prompt) => prompt.message === "Session Mode for this run");
  assert.equal(selectors.length, 2);
  assert(selectors.every((prompt) => prompt.initialValue === "isolated"));
});

test("session selection and continuous execution cancellation never launch tests", async () => {
  for (const selection of ["full", "retest", "setup"] as const) {
    for (const answer of ["selector", false, undefined]) {
      const calls: string[] = [];
      const ui = new ScriptedUi([
        ...(selection === "setup" ? ["setup"] : selection === "full" ? ["run", "remaining"] : ["retest", "ready"]),
        ...(answer === "selector" ? [undefined] : ["continuous", answer]),
        ...(answer === false ? [...(selection === "setup" ? [] : ["back"]), "exit"] : []),
      ]);
      await runControlPanel(ui, stubActions({
        setup: async () => "full-test",
        validateRetest: async () => {
          calls.push("validate");
          return { selectedCount: 2, finalCleanupOnly: false, shouldExecute: true, readyToExecute: true };
        },
        runPgn: async () => { calls.push("execute"); },
        runRetest: async () => { calls.push("execute"); },
      }));
      assert.deepEqual(calls, selection === "retest" && answer !== "selector" ? ["validate"] : []);
      assert(ui.confirmPrompts.every((prompt) => prompt.initialValue === false));
      if (answer !== "selector") assert(ui.events.includes(`warn:${CONTINUOUS_SESSION_WARNING}`));
    }
  }
});

test("fresh workbook preparation remains separate from session selection and execution", async () => {
  const calls: string[] = [];
  const ui = new ScriptedUi(["run", "fresh", true, "back", "exit"]);
  await runControlPanel(ui, stubActions({
    prepareFresh: async () => { calls.push("prepare"); },
    runPgn: async () => { calls.push("execute"); },
  }));
  assert.deepEqual(calls, ["prepare"]);
  assert.equal(ui.selectPrompts.some((prompt) => prompt.message === "Session Mode for this run"), false);
});

test("legacy retest resume retains its stored mode and final-cleanup-only flow without a new selector", async () => {
  const received: string[][] = [];
  const ui = new ScriptedUi(["retest", "resume", "RETEST-FIXTURE", true, "back", "exit"]);
  await runControlPanel(ui, stubActions({
    validateRetest: async (args) => {
      received.push([...args]);
      return { selectedCount: 0, finalCleanupOnly: true, shouldExecute: true, readyToExecute: true };
    },
    runRetest: async (args) => { received.push([...args]); },
  }));
  assert.deepEqual(received, [["--resume", "RETEST-FIXTURE"], ["--resume", "RETEST-FIXTURE"]]);
  assert.equal(ui.selectPrompts.some((prompt) => prompt.message === "Session Mode for this run"), false);
  assert.match(ui.confirmPrompts[0]?.message ?? "", /final WhatsApp session cleanup/);
});

test("Notifications menu separates safe status, confirmed test, and configuration", async () => {
  const validations: boolean[] = [];
  let configurations = 0;
  const ui = new ScriptedUi([
    "notifications",
    "status",
    "test",
    false,
    "test",
    true,
    "configure",
    "back",
    "exit",
  ]);
  await runControlPanel(
    ui,
    stubActions({
      validateDiscord: async (sendTest) => {
        validations.push(sendTest);
        return {
          enabled: true,
          configured: true,
          valid: true,
          connectivity: "ok",
          testNotificationSent: sendTest,
        };
      },
      configureNotifications: async () => {
        configurations += 1;
      },
    }),
  );

  assert.deepEqual(validations, [false, true]);
  assert.equal(configurations, 1);
  assert.match(ui.events.join("\n"), /Discord test notification cancelled/);
  assert.match(ui.events.join("\n"), /Discord test notification sent/);
});

test("setup completion routes diagnostics and confirmed full tests through existing actions", async () => {
  let diagnostics = 0;
  const diagnosticsUi = new ScriptedUi(["setup", "exit"]);
  await runControlPanel(
    diagnosticsUi,
    stubActions({
      setup: async () => "diagnostics",
      diagnostics: async () => {
        diagnostics += 1;
      },
    }),
  );
  assert.equal(diagnostics, 1);

  let fullRuns = 0;
  const fullTestUi = new ScriptedUi(["setup", "isolated", true, "exit"]);
  await runControlPanel(
    fullTestUi,
    stubActions({
      setup: async () => "full-test",
      runPgn: async (args) => {
        assert.deepEqual(args, []);
        fullRuns += 1;
      },
    }),
  );
  assert.equal(fullRuns, 1);
});

test("zero-candidate retest returns to the menu without execution", async () => {
  let executions = 0;
  const ui = new ScriptedUi(["retest", "ready", "isolated", "back", "exit"]);
  await runControlPanel(
    ui,
    stubActions({
      runRetest: async () => {
        executions += 1;
      },
    }),
  );
  assert.equal(executions, 0);
  assert.match(ui.events.join("\n"), /Nothing will be executed/);
});

test("nonzero ready retest selection executes only after confirmation", async () => {
  const received: string[][] = [];
  const ui = new ScriptedUi(["retest", "ready", "isolated", true, "back", "exit"]);
  await runControlPanel(
    ui,
    stubActions({
      validateRetest: async () => ({
        selectedCount: 3,
        finalCleanupOnly: false,
        shouldExecute: true,
        readyToExecute: true,
      }),
      runRetest: async (args) => {
        received.push(args);
      },
    }),
  );
  assert.deepEqual(received, [[]]);
});

test("failed retest prerequisites prevent execution", async () => {
  let executions = 0;
  const ui = new ScriptedUi(["retest", "ready", "isolated", "back", "exit"]);
  await runControlPanel(
    ui,
    stubActions({
      validateRetest: async () => ({
        selectedCount: 2,
        finalCleanupOnly: false,
        shouldExecute: true,
        readyToExecute: false,
      }),
      runRetest: async () => {
        executions += 1;
      },
    }),
  );
  assert.equal(executions, 0);
  assert.match(ui.events.join("\n"), /prerequisites are not ready/);
});

test("Ctrl+C cancellation closes the control panel cleanly", async () => {
  const ui = new ScriptedUi([undefined]);
  await runControlPanel(ui, stubActions());
  assert.match(ui.events.join("\n"), /cancel:Operator control panel closed/);
});

test("back navigation performs no submenu action", async () => {
  let validations = 0;
  const ui = new ScriptedUi(["validate", "back", "exit"]);
  await runControlPanel(
    ui,
    stubActions({
      validatePgn: async () => {
        validations += 1;
        return true;
      },
    }),
  );
  assert.equal(validations, 0);
  assert.match(ui.events.at(-1) ?? "", /outro:Operator control panel closed/);
});

test("environment text updates one key without duplicating or removing unknown fields", () => {
  const updated = updateEnvironmentText(
    "# comment\nVALUE=old\nUNKNOWN=keep\nVALUE=duplicate\n",
    { VALUE: "new value", EMPTY: "" },
  );
  assert.equal(updated.match(/^VALUE=/gm)?.length, 1);
  assert.match(updated, /VALUE="new value"/);
  assert.match(updated, /UNKNOWN=keep/);
  assert.match(updated, /EMPTY=\n/);
});

test("workbook validation CLI accepts only central session flags, including no-op execution filters rejection", () => {
  assert.equal(parsePgnValidationArgs([]), "isolated");
  for (const sessionMode of ["isolated", "continuous"] as const) {
    assert.equal(parsePgnValidationArgs([`--session=${sessionMode}`]), sessionMode);
    assert.equal(parsePgnValidationArgs(["--session", sessionMode]), sessionMode);
  }
  assert.equal(parsePgnValidationArgs(["--fast"]), "continuous");
  assert.equal(parsePgnValidationArgs(["--fast", "--session=continuous"]), "continuous");
  for (const args of [
    ["--limit", "1"], ["--sheet", "all"], ["--sheet", "kb"],
    ["--test", "TC-001"], ["--test", ",,"], ["--rerun"], ["--rerun", "TC-001"],
    ["--resume", "RUN-001"], ["--restart-run", "RUN-001"],
    ["--resume", "RUN-001", "--accept-source-drift"],
  ]) {
    assert.throws(() => parsePgnValidationArgs(args), /execution and recovery filters are not allowed/);
    assert.throws(() => parsePgnValidationArgs(["--session=continuous", ...args]), /execution and recovery filters are not allowed/);
  }
  for (const args of [["--session"], ["--session="], ["--session=invalid"], ["--fast", "--session=isolated"], ["--unknown"]]) {
    assert.throws(() => parsePgnValidationArgs(args));
  }
});

test("workbook validation describes isolated defaults and continuous reset policy without live actions", () => {
  const summary = { scenarios: 0, runnableTurns: 0, missingUserInput: 0, multiTurnScenarios: 0, completedScenarios: 0 };
  const parsed: ParsedPgnWorkbook = {
    scenarios: [], issues: [], summaries: { kb: summary, negative: summary },
    duplicateTestCaseIds: 0, invalidTurnRows: 0,
  };
  const isolation = {
    command: "reset", confirmation: "Session deleted", timeoutMs: 30_000,
    responseIdleMs: 10_000, responseTimeoutMs: 60_000, postResetQuietMs: 10_000,
  };
  const legacy = formatPgnValidation(parsed, isolation);
  const isolated = formatPgnValidation(parsed, { ...isolation, sessionMode: "isolated" });
  assert.equal(legacy, isolated);
  assert.match(isolated, /Session Mode: Isolated/);
  assert.match(isolated, /Between-scenario reset: Enabled/);
  assert.match(isolated, /Final cleanup reset: Enabled/);
  assert.equal(isolated.includes(CONTINUOUS_SESSION_WARNING), false);

  const continuous = formatPgnValidation(parsed, { ...isolation, sessionMode: "continuous" });
  assert.match(continuous, /Session Mode: Continuous/);
  assert.match(continuous, /Between-scenario reset: Disabled/);
  assert.match(continuous, /Final cleanup reset: Disabled/);
  assert(continuous.includes(CONTINUOUS_SESSION_WARNING));
  for (const output of [isolated, continuous]) {
    assert.match(output, /Transport: WhatsApp/);
    assert.match(output, /Initial reset: Required/);
    assert.match(output, /Expected confirmation: "Session deleted"/);
    assert.match(output, /Reset timeout: 30000 ms/);
    assert.match(output, /Post-reset quiet window: 10000 ms/);
    assert.match(output, /READY TO EXECUTE/);
  }
  parsed.issues.push({ code: "MISSING_SHEET", severity: "ERROR", sheetName: "Negative Case", message: "Missing sheet" });
  assert.match(formatPgnValidation(parsed, { ...isolation, sessionMode: "continuous" }), /NOT READY$/);
});
