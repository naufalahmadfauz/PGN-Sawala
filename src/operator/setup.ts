import { access, readFile, rename, rm, writeFile } from "node:fs/promises";
import path from "node:path";
import dotenv from "dotenv";
import { normalizeGoogleDriveFolderId } from "../config";
import {
  resolveGoogleServiceAccount,
  safeGoogleCredentialError,
  type GoogleServiceAccountConfiguration,
} from "../evidence/google-service-account";
import {
  REPOSITORY_ROOT,
  loadEnvironment,
  synchronizeEnvironmentFileUpdates,
} from "../environment";
import {
  safeDiscordError,
  validateDiscordWebhook,
  validateDiscordWebhookUrl,
  type DiscordValidationResult,
} from "../notifications/discord";
import {
  collectDiagnostics,
  formatSetupChecklist,
  formatSetupInspection,
  type DiagnosticReport,
} from "./diagnostics";
import { runInheritedCommand } from "./process";
import type { OperatorUi } from "./ui";

const CONFIGURATION_KEYS = [
  "PGN_WHATSAPP_PHONE",
  "PGN_WHATSAPP_CHAT",
  "WHATSAPP_HEADLESS",
  "GOOGLE_DRIVE_EVIDENCE_ENABLED",
  "GOOGLE_DRIVE_EVIDENCE_PARENT_FOLDER",
  "GOOGLE_SERVICE_ACCOUNT_JSON",
  "GOOGLE_SERVICE_ACCOUNT_JSON_BASE64",
  "GOOGLE_SERVICE_ACCOUNT_FILE",
  "DISCORD_NOTIFICATIONS_ENABLED",
  "DISCORD_WEBHOOK_URL",
  "DISCORD_PROGRESS_EVERY",
  "DISCORD_PROGRESS_MINUTES",
  "DISCORD_NOTIFY_START",
  "DISCORD_NOTIFY_PROGRESS",
] as const;
const DISCORD_SETUP_KEYS = [
  "DISCORD_NOTIFICATIONS_ENABLED",
  "DISCORD_WEBHOOK_URL",
  "DISCORD_PROGRESS_EVERY",
  "DISCORD_PROGRESS_MINUTES",
  "DISCORD_NOTIFY_START",
  "DISCORD_NOTIFY_PROGRESS",
] as const;
type DiscordSetupKey = (typeof DISCORD_SETUP_KEYS)[number];

export interface DiscordSetupDependencies {
  projectRoot?: string;
  environment?: NodeJS.ProcessEnv;
  testDiscordWebhook?: (
    webhookUrl: string,
  ) => Promise<DiscordValidationResult>;
}

export interface SetupDependencies extends DiscordSetupDependencies {
  platform?: NodeJS.Platform;
  diagnose?: (checkDriveAccess: boolean) => Promise<DiagnosticReport>;
  installChromium?: (withDependencies: boolean) => Promise<void>;
  loginWhatsApp?: () => Promise<void>;
  validateCredentialFile?: (
    projectRoot: string,
    configuredPath: string,
  ) => Promise<void>;
}

export interface SetupResult {
  cancelled: boolean;
  environmentUpdated: boolean;
  chromiumInstalled: boolean;
  loginStarted: boolean;
  nextAction?: SetupNextAction;
}

export type SetupNextAction =
  | "diagnostics"
  | "main-menu"
  | "full-test"
  | "exit";

function envValue(value: string): string {
  if (!value) return "";
  return /^[A-Za-z0-9_./:+-]+$/.test(value) ? value : JSON.stringify(value);
}

export function updateEnvironmentText(
  source: string,
  updates: Readonly<Record<string, string>>,
): string {
  const hadFinalNewline = source.endsWith("\n");
  const lines = source ? source.replace(/\r\n/g, "\n").split("\n") : [];
  if (hadFinalNewline && lines.at(-1) === "") {
    lines.pop();
  }
  const remaining = new Set(Object.keys(updates));
  const output: string[] = [];
  for (const line of lines) {
    const match = line.match(/^\s*(?:export\s+)?([A-Za-z_][A-Za-z0-9_]*)\s*=/);
    const name = match?.[1];
    if (!name || !(name in updates)) {
      output.push(line);
      continue;
    }
    if (remaining.has(name)) {
      output.push(`${name}=${envValue(updates[name])}`);
      remaining.delete(name);
    }
  }
  if (remaining.size && output.length && output.at(-1)?.trim()) {
    output.push("");
  }
  for (const name of remaining) {
    output.push(`${name}=${envValue(updates[name])}`);
  }
  return `${output.join("\n")}\n`;
}

export async function writeEnvironmentUpdates(
  projectRoot: string,
  updates: Readonly<Record<string, string>>,
): Promise<void> {
  const environmentPath = path.join(projectRoot, ".env");
  const temporaryPath = `${environmentPath}.${process.pid}.tmp`;
  let source = "";
  try {
    source = await readFile(environmentPath, "utf8");
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code !== "ENOENT") throw error;
  }
  await rm(temporaryPath, { force: true });
  try {
    await writeFile(temporaryPath, updateEnvironmentText(source, updates), {
      encoding: "utf8",
      mode: 0o600,
    });
    await rename(temporaryPath, environmentPath);
  } catch (error) {
    await rm(temporaryPath, { force: true }).catch(() => undefined);
    throw error;
  }
}

async function readEnvironmentValues(
  projectRoot: string,
): Promise<Record<string, string>> {
  try {
    return dotenv.parse(await readFile(path.join(projectRoot, ".env"), "utf8"));
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code === "ENOENT") return {};
    throw error;
  }
}

async function defaultValidateCredentialFile(
  projectRoot: string,
  configuredPath: string,
): Promise<void> {
  const absolutePath = path.resolve(projectRoot, configuredPath);
  await access(absolutePath);
  const configuration: GoogleServiceAccountConfiguration = {
    source: "GOOGLE_SERVICE_ACCOUNT_FILE",
    configuredSources: ["GOOGLE_SERVICE_ACCOUNT_FILE"],
    environmentSource: ".env",
    filePath: absolutePath,
    fileDisplayPath: configuredPath,
  };
  resolveGoogleServiceAccount(configuration);
}

export async function installPlaywrightChromium(
  projectRoot: string,
  withDependencies: boolean,
): Promise<void> {
  const cliPath = path.join(projectRoot, "node_modules", "playwright", "cli.js");
  await access(cliPath);
  await runInheritedCommand(
    process.execPath,
    [
      cliPath,
      "install",
      ...(withDependencies ? ["--with-deps"] : []),
      "chromium",
    ],
    { cwd: projectRoot },
  );
}

function configuredValue(
  name: (typeof CONFIGURATION_KEYS)[number],
  fileValues: Record<string, string>,
  environment: NodeJS.ProcessEnv,
): string {
  return environment[name] ?? fileValues[name] ?? "";
}

function configuredBoolean(
  name: (typeof CONFIGURATION_KEYS)[number],
  fileValues: Record<string, string>,
  environment: NodeJS.ProcessEnv,
  fallback: boolean,
): boolean {
  const value = configuredValue(name, fileValues, environment).trim();
  if (!value) return fallback;
  if (/^(?:true|1)$/i.test(value)) return true;
  if (/^(?:false|0)$/i.test(value)) return false;
  return fallback;
}

function configuredPositiveInteger(
  name: (typeof CONFIGURATION_KEYS)[number],
  fileValues: Record<string, string>,
  environment: NodeJS.ProcessEnv,
  fallback: number,
): number {
  const value = Number(configuredValue(name, fileValues, environment));
  return Number.isInteger(value) && value >= 1 ? value : fallback;
}

async function defaultTestDiscordWebhook(
  webhookUrl: string,
): Promise<DiscordValidationResult> {
  return validateDiscordWebhook(
    {
      discordNotificationsEnabled: true,
      discordWebhookUrl: webhookUrl,
    },
    { sendTest: true },
  );
}

async function promptDiscordUpdates(
  ui: OperatorUi,
  fileValues: Record<string, string>,
  environment: NodeJS.ProcessEnv,
  testDiscordWebhook: (webhookUrl: string) => Promise<DiscordValidationResult>,
  processManaged: ReadonlySet<DiscordSetupKey>,
): Promise<Record<string, string> | undefined> {
  const currentlyEnabled = configuredBoolean(
    "DISCORD_NOTIFICATIONS_ENABLED",
    fileValues,
    environment,
    false,
  );
  const updates: Record<string, string> = {};
  let enabled = currentlyEnabled;
  if (processManaged.has("DISCORD_NOTIFICATIONS_ENABLED")) {
    ui.warn(
      "Discord enablement is managed by the process environment. Update that environment or Codespaces Secret instead of .env.",
    );
  } else {
    const selected = await ui.confirm({
      message: "Enable Discord notifications?",
      initialValue: currentlyEnabled,
    });
    if (selected === undefined) return undefined;
    enabled = selected;
    updates.DISCORD_NOTIFICATIONS_ENABLED = String(enabled);
  }
  if (!enabled) return updates;

  const existingWebhook = configuredValue(
    "DISCORD_WEBHOOK_URL",
    fileValues,
    environment,
  );
  const existingValidation = validateDiscordWebhookUrl(existingWebhook);
  let webhookUrl: string | undefined;
  if (processManaged.has("DISCORD_WEBHOOK_URL")) {
    if (existingValidation.valid) {
      webhookUrl = existingValidation.url;
      ui.info(
        "The Discord webhook is managed by the process environment and remains hidden.",
      );
    } else {
      ui.warn(
        "The process-managed Discord webhook is missing or invalid. Update that environment or Codespaces Secret; .env cannot override it.",
      );
    }
  } else if (existingValidation.valid) {
    const keep = await ui.confirm({
      message: "Keep the currently configured Discord webhook?",
      initialValue: true,
    });
    if (keep === undefined) return undefined;
    if (keep) {
      webhookUrl = existingValidation.url;
    } else {
      const replacement = await ui.secret({
        message: "Discord Incoming Webhook URL",
        mask: "*",
        clearOnError: true,
        validate: (value) => {
          const validation = validateDiscordWebhookUrl(value);
          return validation.valid ? undefined : validation.reason;
        },
      });
      if (replacement === undefined) return undefined;
      const validation = validateDiscordWebhookUrl(replacement);
      if (!validation.valid) return undefined;
      webhookUrl = validation.url;
      updates.DISCORD_WEBHOOK_URL = webhookUrl;
    }
  } else {
    const entered = await ui.secret({
      message: "Discord Incoming Webhook URL",
      mask: "*",
      clearOnError: true,
      validate: (value) => {
        const validation = validateDiscordWebhookUrl(value);
        return validation.valid ? undefined : validation.reason;
      },
    });
    if (entered === undefined) return undefined;
    const validation = validateDiscordWebhookUrl(entered);
    if (!validation.valid) return undefined;
    webhookUrl = validation.url;
    updates.DISCORD_WEBHOOK_URL = webhookUrl;
  }

  const managedProgressKeys = [
    "DISCORD_NOTIFY_PROGRESS",
    "DISCORD_PROGRESS_EVERY",
    "DISCORD_PROGRESS_MINUTES",
  ].filter((name): name is DiscordSetupKey =>
    processManaged.has(name as DiscordSetupKey),
  );
  if (managedProgressKeys.length) {
    ui.warn(
      `Discord progress settings are managed by the process environment (${managedProgressKeys.join(", ")}) and were left unchanged.`,
    );
  } else {
    const currentProgressEnabled = configuredBoolean(
      "DISCORD_NOTIFY_PROGRESS",
      fileValues,
      environment,
      true,
    );
    const currentProgressEvery = configuredPositiveInteger(
      "DISCORD_PROGRESS_EVERY",
      fileValues,
      environment,
      5,
    );
    const currentProgressMinutes = configuredPositiveInteger(
      "DISCORD_PROGRESS_MINUTES",
      fileValues,
      environment,
      2,
    );
    const progressChoice = await ui.select({
      message: "Discord progress updates",
      options: [
        {
          value: "5",
          label: "Every 5 scenarios",
          hint: `or every ${currentProgressMinutes} minutes`,
        },
        {
          value: "10",
          label: "Every 10 scenarios",
          hint: `or every ${currentProgressMinutes} minutes`,
        },
        { value: "final", label: "Final result only" },
        { value: "custom", label: "Custom frequency" },
      ],
      initialValue: !currentProgressEnabled
        ? "final"
        : currentProgressEvery === 5
          ? "5"
          : currentProgressEvery === 10
            ? "10"
            : "custom",
    });
    if (progressChoice === undefined) return undefined;
    const finalOnly = progressChoice === "final";
    updates.DISCORD_NOTIFY_PROGRESS = String(!finalOnly);
    if (finalOnly) {
      if (processManaged.has("DISCORD_NOTIFY_START")) {
        ui.warn(
          "Start notifications remain process-managed, so the final-only preset cannot change them.",
        );
      } else {
        updates.DISCORD_NOTIFY_START = "false";
      }
    }
    if (progressChoice === "5" || progressChoice === "10") {
      updates.DISCORD_PROGRESS_EVERY = progressChoice;
      updates.DISCORD_PROGRESS_MINUTES = String(currentProgressMinutes);
    } else if (progressChoice === "custom") {
      const every = await ui.text({
        message: "Scenarios between Discord progress updates",
        initialValue: String(currentProgressEvery),
        validate: (value) =>
          Number.isInteger(Number(value)) && Number(value) >= 1
            ? undefined
            : "Enter a whole number of 1 or greater",
      });
      if (every === undefined) return undefined;
      const minutes = await ui.text({
        message: "Minutes between Discord progress updates",
        initialValue: String(currentProgressMinutes),
        validate: (value) =>
          Number.isInteger(Number(value)) && Number(value) >= 1
            ? undefined
            : "Enter a whole number of 1 or greater",
      });
      if (minutes === undefined) return undefined;
      updates.DISCORD_PROGRESS_EVERY = String(Number(every));
      updates.DISCORD_PROGRESS_MINUTES = String(Number(minutes));
    }
  }

  if (!webhookUrl) return updates;
  const sendTest = await ui.confirm({
    message: "Send one visible Discord test notification now?",
    initialValue: false,
  });
  if (sendTest === undefined) return undefined;
  if (sendTest) {
    try {
      const result = await ui.task(
        "Sending one Discord test notification",
        () => testDiscordWebhook(webhookUrl),
        "Discord test request completed",
      );
      if (result.testNotificationSent) {
        ui.success("Discord test notification sent.");
      } else if (result.testNotificationDeliveryUncertain) {
        ui.warn(
          `Discord test notification delivery could not be confirmed; check the channel before retrying: ${safeDiscordError(new Error(result.reason ?? "request outcome is unknown"), webhookUrl)}`,
        );
      } else {
        ui.warn(
          `Discord test notification was not sent: ${safeDiscordError(new Error(result.reason ?? "validation failed"), webhookUrl)}`,
        );
      }
    } catch (error) {
      ui.warn(
        `Discord test notification delivery could not be confirmed; check the channel before retrying: ${safeDiscordError(error, webhookUrl)}`,
      );
    }
  }
  return updates;
}

export async function configureDiscordNotifications(
  ui: OperatorUi,
  dependencies: DiscordSetupDependencies = {},
): Promise<boolean> {
  const projectRoot = path.resolve(dependencies.projectRoot ?? REPOSITORY_ROOT);
  const loadedEnvironment = loadEnvironment({
    repositoryRoot: projectRoot,
    values: dependencies.environment ?? process.env,
  });
  const environment = loadedEnvironment.values;
  const fileValues = await readEnvironmentValues(projectRoot);
  const updates = await promptDiscordUpdates(
    ui,
    fileValues,
    environment,
    dependencies.testDiscordWebhook ?? defaultTestDiscordWebhook,
    new Set(
      DISCORD_SETUP_KEYS.filter(
        (name) => loadedEnvironment.sourceFor(name) === "process environment",
      ),
    ),
  );
  if (!updates) {
    ui.info("Notification configuration cancelled");
    return false;
  }
  if (Object.keys(updates).length === 0) {
    ui.info("No notification settings were changed.");
    return true;
  }
  await ui.task(
    "Writing notification settings safely",
    () => writeEnvironmentUpdates(projectRoot, updates),
    "Notification settings saved",
  );
  if (
    dependencies.environment === undefined &&
    projectRoot === REPOSITORY_ROOT
  ) {
    synchronizeEnvironmentFileUpdates(updates);
  }
  ui.success("Notification settings saved without displaying the webhook.");
  return true;
}

function cancelled(ui: OperatorUi): SetupResult {
  ui.cancel("Setup cancelled");
  return {
    cancelled: true,
    environmentUpdated: false,
    chromiumInstalled: false,
    loginStarted: false,
  };
}

export async function runSetupWizard(
  ui: OperatorUi,
  dependencies: SetupDependencies = {},
): Promise<SetupResult> {
  const projectRoot = path.resolve(dependencies.projectRoot ?? REPOSITORY_ROOT);
  const platform = dependencies.platform ?? process.platform;
  const loadedEnvironment = loadEnvironment({
    repositoryRoot: projectRoot,
    values: dependencies.environment ?? process.env,
  });
  const environment = loadedEnvironment.values;
  const diagnose =
    dependencies.diagnose ??
    ((checkDriveAccess: boolean) =>
      collectDiagnostics({
        projectRoot,
        platform,
        environment,
        checkDriveAccess,
        checkDiscordAccess: false,
      }));
  const installChromium =
    dependencies.installChromium ??
    ((withDependencies: boolean) =>
      installPlaywrightChromium(projectRoot, withDependencies));
  const validateCredentialFile =
    dependencies.validateCredentialFile ?? defaultValidateCredentialFile;

  ui.intro("PGN Sawala setup");
  let report = await ui.task(
    "Inspecting this workstation",
    () => diagnose(false),
    "Workstation inspected",
  );
  ui.note(formatSetupInspection(report), "Environment");

  const configure = await ui.confirm({
    message: "Would you like to review or update your setup?",
    initialValue: !report.environmentFilePresent,
    active: "Yes",
    inactive: "No, keep current settings",
  });
  if (configure === undefined) return cancelled(ui);

  let environmentUpdated = false;
  if (configure) {
    const fileValues = await readEnvironmentValues(projectRoot);
    const updates: Record<string, string> = {};
    const existingPhone = configuredValue(
      "PGN_WHATSAPP_PHONE",
      fileValues,
      environment,
    );
    const existingChat = configuredValue(
      "PGN_WHATSAPP_CHAT",
      fileValues,
      environment,
    );
    const target = await ui.select({
      message: "How should the WhatsApp test chat be identified?",
      options: [
        ...(existingPhone || existingChat
          ? [
              {
                value: "keep" as const,
                label: "Keep current target",
                hint: "Leave the configured target unchanged",
              },
            ]
          : []),
        {
          value: "phone" as const,
          label: "Phone number",
          hint: "Recommended; international digits only",
        },
        { value: "chat" as const, label: "Chat name" },
        { value: "skip" as const, label: "Configure later" },
      ],
      initialValue: existingPhone || existingChat ? "keep" : "phone",
    });
    if (target === undefined) return cancelled(ui);
    if (target === "phone") {
      const phone = await ui.text({
        message: "PGN WhatsApp phone number",
        placeholder: "International number without + or spaces",
        initialValue: existingPhone || undefined,
        validate(value) {
          const digits = value.replace(/\D/g, "");
          return digits.length >= 8
            ? undefined
            : "Enter at least 8 digits including the country code";
        },
      });
      if (phone === undefined) return cancelled(ui);
      updates.PGN_WHATSAPP_PHONE = phone.replace(/\D/g, "");
      updates.PGN_WHATSAPP_CHAT = "";
    } else if (target === "chat") {
      const chat = await ui.text({
        message: "WhatsApp chat name",
        initialValue: existingChat || undefined,
        validate: (value) =>
          value.trim() ? undefined : "Chat name must not be empty",
      });
      if (chat === undefined) return cancelled(ui);
      updates.PGN_WHATSAPP_CHAT = chat.trim();
      updates.PGN_WHATSAPP_PHONE = "";
    }

    const currentHeadless = /^(?:true|1)$/i.test(
      configuredValue("WHATSAPP_HEADLESS", fileValues, environment),
    );
    const headless = await ui.confirm({
      message: "Run Chromium headless?",
      initialValue: currentHeadless,
    });
    if (headless === undefined) return cancelled(ui);
    updates.WHATSAPP_HEADLESS = String(headless);

    const currentDriveEnabled = /^(?:true|1)$/i.test(
      configuredValue(
        "GOOGLE_DRIVE_EVIDENCE_ENABLED",
        fileValues,
        environment,
      ),
    );
    const driveEnabled = await ui.confirm({
      message: "Upload evidence to Google Drive?",
      initialValue: currentDriveEnabled,
    });
    if (driveEnabled === undefined) return cancelled(ui);
    updates.GOOGLE_DRIVE_EVIDENCE_ENABLED = String(driveEnabled);
    if (driveEnabled) {
      const currentFolder = configuredValue(
        "GOOGLE_DRIVE_EVIDENCE_PARENT_FOLDER",
        fileValues,
        environment,
      );
      const folder = await ui.text({
        message: "Shared Drive parent folder URL or ID",
        initialValue: currentFolder || undefined,
        validate(value) {
          try {
            normalizeGoogleDriveFolderId(value);
            return undefined;
          } catch (error) {
            return error instanceof Error ? error.message : String(error);
          }
        },
      });
      if (folder === undefined) return cancelled(ui);
      updates.GOOGLE_DRIVE_EVIDENCE_PARENT_FOLDER = folder.trim();

      const hasExistingCredential = [
        "GOOGLE_SERVICE_ACCOUNT_JSON",
        "GOOGLE_SERVICE_ACCOUNT_JSON_BASE64",
        "GOOGLE_SERVICE_ACCOUNT_FILE",
      ].some((name) => Boolean(environment[name] ?? fileValues[name]));
      const credentialChoice = await ui.select({
        message: "Google service-account credentials",
        options: [
          ...(hasExistingCredential
            ? [
                {
                  value: "keep" as const,
                  label: "Keep current credentials",
                  hint: "No credential value will be displayed",
                },
              ]
            : []),
          {
            value: "file" as const,
            label: "Service-account file",
            hint: "Recommended for local development",
          },
        ],
        initialValue: hasExistingCredential ? "keep" : "file",
      });
      if (credentialChoice === undefined) return cancelled(ui);
      if (credentialChoice === "file") {
        const currentCredentialFile = configuredValue(
          "GOOGLE_SERVICE_ACCOUNT_FILE",
          fileValues,
          environment,
        );
        while (true) {
          const credentialPath = await ui.text({
            message: "Service-account JSON file",
            placeholder: ".secrets/google-service-account.json",
            defaultValue: ".secrets/google-service-account.json",
            initialValue: currentCredentialFile || undefined,
            validate: (value) =>
              value.trim() ? undefined : "Credential path must not be empty",
          });
          if (credentialPath === undefined) return cancelled(ui);
          try {
            await validateCredentialFile(projectRoot, credentialPath.trim());
            updates.GOOGLE_SERVICE_ACCOUNT_FILE = credentialPath.trim();
            updates.GOOGLE_SERVICE_ACCOUNT_JSON = "";
            updates.GOOGLE_SERVICE_ACCOUNT_JSON_BASE64 = "";
            break;
          } catch (error) {
            ui.error(`Credential file is not valid: ${safeGoogleCredentialError(error)}`);
            const retry = await ui.confirm({
              message: "Choose a different credential file?",
              initialValue: true,
            });
            if (retry === undefined) return cancelled(ui);
            if (!retry) {
              ui.warn("Configuration was not changed.");
              ui.outro("Setup stopped before writing an invalid Drive configuration");
              return {
                cancelled: false,
                environmentUpdated: false,
                chromiumInstalled: false,
                loginStarted: false,
              };
            }
          }
        }
      }
    }

    const discordUpdates = await promptDiscordUpdates(
      ui,
      fileValues,
      environment,
      dependencies.testDiscordWebhook ?? defaultTestDiscordWebhook,
      new Set(
        DISCORD_SETUP_KEYS.filter(
          (name) => loadedEnvironment.sourceFor(name) === "process environment",
        ),
      ),
    );
    if (!discordUpdates) return cancelled(ui);
    Object.assign(updates, discordUpdates);

    await ui.task(
      "Writing .env safely",
      () => writeEnvironmentUpdates(projectRoot, updates),
      ".env updated",
    );
    if (
      dependencies.environment === undefined &&
      projectRoot === REPOSITORY_ROOT
    ) {
      synchronizeEnvironmentFileUpdates(updates);
    }
    environmentUpdated = true;
    ui.success("Configuration saved without exposing credential values.");
  }

  report = await diagnose(false);
  let chromiumInstalled = false;
  if (!report.chromiumInstalled) {
    const installChoice = await ui.select({
      message: "Playwright Chromium is missing. Install it now?",
      options: [
        { value: "browser" as const, label: "Install Chromium" },
        ...(platform === "linux"
          ? [
              {
                value: "dependencies" as const,
                label: "Install Chromium + OS deps",
                hint: "Runs playwright install --with-deps",
              },
            ]
          : []),
        { value: "skip" as const, label: "Skip for now" },
      ],
      initialValue: platform === "linux" ? "dependencies" : "browser",
    });
    if (installChoice === undefined) return cancelled(ui);
    if (installChoice !== "skip") {
      const withDependencies = installChoice === "dependencies";
      ui.note(
        `npx playwright install${withDependencies ? " --with-deps" : ""} chromium`,
        "Install command",
      );
      await ui.task(
        "Installing Playwright Chromium",
        () => installChromium(withDependencies),
        "Playwright Chromium installed",
      );
      chromiumInstalled = true;
    }
  }

  report = await diagnose(true);
  let loginStarted = false;
  if (!report.profilePresent && dependencies.loginWhatsApp) {
    const login = await ui.confirm({
      message: "Open WhatsApp login now?",
      initialValue: false,
    });
    if (login === undefined) return cancelled(ui);
    if (login) {
      await dependencies.loginWhatsApp();
      loginStarted = true;
      report = await diagnose(true);
    }
  }
  ui.note(formatSetupChecklist(report), "Setup checklist");
  ui.success("Setup complete");
  if (!report.ready) {
    ui.warn("Some items need attention. Run diagnostics to show details.");
  }
  const nextAction = await ui.select<SetupNextAction>({
    message: "What would you like to do next?",
    options: [
      {
        value: "diagnostics",
        label: "Run diagnostics",
        hint: "Show detailed technical output",
      },
      { value: "main-menu", label: "Open main menu" },
      {
        value: "full-test",
        label: "Start full test",
        hint: report.ready ? undefined : "Resolve setup issues first",
        disabled: !report.ready,
      },
      { value: "exit", label: "Exit" },
    ],
    initialValue: report.ready ? "main-menu" : "diagnostics",
  });
  return {
    cancelled: false,
    environmentUpdated,
    chromiumInstalled,
    loginStarted,
    nextAction: nextAction ?? "exit",
  };
}
