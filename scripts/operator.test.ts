import assert from "node:assert/strict";
import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import test from "node:test";
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
});

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
  });

  const discord = report.checks.find((check) => check.id === "discord");
  assert.equal(inspections, 1);
  assert.equal(discord?.status, "warn");
  assert.equal((discord?.detail ?? "").includes(webhook), false);
  assert.equal(report.ready, true);
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
  const fullTestUi = new ScriptedUi(["setup", true, "exit"]);
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
  const ui = new ScriptedUi(["retest", "ready", "back", "exit"]);
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
  const ui = new ScriptedUi(["retest", "ready", true, "back", "exit"]);
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
  const ui = new ScriptedUi(["retest", "ready", "back", "exit"]);
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
