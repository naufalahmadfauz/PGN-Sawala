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
  synchronizeEnvironmentFileUpdates,
} from "../environment";
import {
  collectDiagnostics,
  formatDiagnosticReport,
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
] as const;

export interface SetupDependencies {
  projectRoot?: string;
  platform?: NodeJS.Platform;
  environment?: NodeJS.ProcessEnv;
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
}

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
  const environment = dependencies.environment ?? process.env;
  const diagnose =
    dependencies.diagnose ??
    ((checkDriveAccess: boolean) =>
      collectDiagnostics({
        projectRoot,
        platform,
        environment,
        checkDriveAccess,
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
  ui.note(formatDiagnosticReport(report), "Environment");

  const configure = await ui.confirm({
    message: report.environmentFilePresent
      ? "Update repository-local configuration?"
      : "Create a repository-local .env configuration?",
    initialValue: !report.environmentFilePresent,
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
  ui.note(formatDiagnosticReport(report), "Final checks");
  ui.outro(
    report.ready
      ? "Setup complete. Run npm run pgn to open the control panel."
      : "Setup complete with items to resolve; run npm run doctor for details.",
  );
  return {
    cancelled: false,
    environmentUpdated,
    chromiumInstalled,
    loginStarted,
  };
}
