import { access, readFile } from "node:fs/promises";
import path from "node:path";
import { loadConfig, type AppConfig } from "../config";
import { createGoogleDriveEvidencePublisher } from "../evidence/google-drive";
import {
  resolveGoogleServiceAccount,
  safeGoogleCredentialError,
} from "../evidence/google-service-account";
import { REPOSITORY_ROOT } from "../environment";
import { detectBrowserRuntime, type BrowserRuntimePlan } from "./browser-runtime";
import { commandAvailable, runProcess } from "./process";

export type DiagnosticStatus = "ok" | "warn" | "error" | "info";

export interface DiagnosticCheck {
  id: string;
  label: string;
  status: DiagnosticStatus;
  detail: string;
}

export interface DiagnosticReport {
  checks: DiagnosticCheck[];
  browserRuntime: BrowserRuntimePlan;
  chromiumInstalled: boolean;
  profilePresent: boolean;
  environmentFilePresent: boolean;
  ready: boolean;
}

export interface DiagnosticDependencies {
  projectRoot?: string;
  platform?: NodeJS.Platform;
  environment?: NodeJS.ProcessEnv;
  pathExists?: (filePath: string) => Promise<boolean>;
  npmVersion?: () => Promise<string | undefined>;
  packageVersion?: (
    packageName: string,
    projectRoot: string,
  ) => Promise<string | undefined>;
  chromiumExecutablePath?: () => Promise<string | undefined>;
  hasCommand?: (command: string, args?: readonly string[]) => Promise<boolean>;
  validateDrive?: (config: AppConfig) => Promise<void>;
  checkDriveAccess?: boolean;
}

async function defaultPathExists(filePath: string): Promise<boolean> {
  return access(filePath)
    .then(() => true)
    .catch(() => false);
}

async function defaultNpmVersion(): Promise<string | undefined> {
  const fromUserAgent = process.env.npm_config_user_agent?.match(/\bnpm\/([^\s]+)/)?.[1];
  if (fromUserAgent) {
    return fromUserAgent;
  }
  const candidates = [
    process.env.npm_execpath,
    path.join(path.dirname(process.execPath), "node_modules", "npm", "bin", "npm-cli.js"),
    path.resolve(
      path.dirname(process.execPath),
      "..",
      "lib",
      "node_modules",
      "npm",
      "bin",
      "npm-cli.js",
    ),
  ].filter((candidate): candidate is string => Boolean(candidate));
  for (const candidate of candidates) {
    try {
      await access(candidate);
      const result = await runProcess(process.execPath, [candidate, "--version"]);
      if (result.exitCode === 0) return result.stdout.trim();
    } catch {
      // Try the next portable npm CLI location.
    }
  }
  try {
    const executable = process.platform === "win32" ? "npm.cmd" : "npm";
    const result = await runProcess(executable, ["--version"]);
    return result.exitCode === 0 ? result.stdout.trim() : undefined;
  } catch {
    return undefined;
  }
}

async function defaultPackageVersion(
  packageName: string,
  projectRoot: string,
): Promise<string | undefined> {
  try {
    const packageJson = JSON.parse(
      await readFile(
        path.join(projectRoot, "node_modules", packageName, "package.json"),
        "utf8",
      ),
    ) as { version?: unknown };
    return typeof packageJson.version === "string"
      ? packageJson.version
      : undefined;
  } catch {
    return undefined;
  }
}

async function defaultChromiumExecutablePath(): Promise<string | undefined> {
  try {
    const { chromium } = await import("playwright");
    return chromium.executablePath();
  } catch {
    return undefined;
  }
}

function nodeVersionSupported(version: string): boolean {
  const [major, minor] = version
    .replace(/^v/, "")
    .split(".")
    .map((part) => Number(part));
  return major > 20 || (major === 20 && minor >= 12);
}

function platformLabel(
  platform: NodeJS.Platform,
  environment: NodeJS.ProcessEnv,
): string {
  if (platform === "win32") return "Windows";
  if (platform === "darwin") return "macOS";
  if (platform === "linux") {
    return /^(?:true|1)$/i.test(environment.CODESPACES ?? "")
      ? "Linux (GitHub Codespaces)"
      : "Linux";
  }
  return platform;
}

export async function collectDiagnostics(
  dependencies: DiagnosticDependencies = {},
): Promise<DiagnosticReport> {
  const projectRoot = path.resolve(dependencies.projectRoot ?? REPOSITORY_ROOT);
  const platform = dependencies.platform ?? process.platform;
  const environment = { ...(dependencies.environment ?? process.env) };
  const pathExists = dependencies.pathExists ?? defaultPathExists;
  const getPackageVersion = dependencies.packageVersion ?? defaultPackageVersion;
  const checks: DiagnosticCheck[] = [];
  const add = (
    id: string,
    label: string,
    status: DiagnosticStatus,
    detail: string,
  ): void => {
    checks.push({ id, label, status, detail });
  };

  add("os", "Operating system", "info", platformLabel(platform, environment));
  add(
    "node",
    "Node.js",
    nodeVersionSupported(process.version) ? "ok" : "error",
    `${process.version} (requires 20.12 or newer)`,
  );
  const npmVersion = await (dependencies.npmVersion ?? defaultNpmVersion)();
  add(
    "npm",
    "npm",
    npmVersion ? "ok" : "error",
    npmVersion ?? "not available",
  );

  const dependencyNames = [
    "@clack/prompts",
    "dotenv",
    "exceljs",
    "googleapis",
    "jszip",
    "playwright",
    "sharp",
    "tsx",
    "typescript",
  ];
  const dependencyVersions = new Map(
    await Promise.all(
      dependencyNames.map(async (name) => [
        name,
        await getPackageVersion(name, projectRoot),
      ] as const),
    ),
  );
  const missingDependencies = dependencyNames.filter(
    (name) => !dependencyVersions.get(name),
  );
  const playwrightVersion = dependencyVersions.get("playwright");
  add(
    "dependencies",
    "Dependencies",
    missingDependencies.length === 0 ? "ok" : "error",
    missingDependencies.length === 0
      ? "installed"
      : `missing ${missingDependencies.join(", ")}; run npm install`,
  );
  add(
    "playwright",
    "Playwright",
    playwrightVersion ? "ok" : "error",
    playwrightVersion ?? "not installed",
  );

  const chromiumPath = await (
    dependencies.chromiumExecutablePath ?? defaultChromiumExecutablePath
  )();
  const chromiumInstalled = Boolean(
    chromiumPath && (await pathExists(chromiumPath)),
  );
  add(
    "chromium",
    "Playwright Chromium",
    chromiumInstalled ? "ok" : "error",
    chromiumInstalled ? "installed" : "missing; install Chromium",
  );

  let config: AppConfig | undefined;
  let configError: unknown;
  try {
    config = loadConfig({ repositoryRoot: projectRoot, environment });
  } catch (error) {
    configError = error;
  }
  if (configError) {
    add(
      "configuration",
      "Configuration",
      "error",
      safeGoogleCredentialError(configError),
    );
  } else {
    add("configuration", "Configuration", "ok", "valid");
  }

  const environmentFilePath = path.join(projectRoot, ".env");
  const environmentFilePresent = await pathExists(environmentFilePath);
  add(
    "env",
    ".env",
    environmentFilePresent ? "ok" : "warn",
    environmentFilePresent
      ? "repository-local configuration found"
      : "not found; process environment and defaults remain available",
  );

  const sourceWorkbookPath =
    config?.pgnSourceWorkbookPath ??
    path.join(
      projectRoot,
      "data",
      "PGN AI Assistant - Knowledge Base Testing Report - User Inputs.xlsx",
    );
  const executedWorkbookPath =
    config?.pgnExecutedWorkbookPath ??
    path.join(
      projectRoot,
      "reports",
      "PGN AI Assistant - Knowledge Base Testing Report - Executed.xlsx",
    );
  const [sourceWorkbookPresent, executedWorkbookPresent] = await Promise.all([
    pathExists(sourceWorkbookPath),
    pathExists(executedWorkbookPath),
  ]);
  add(
    "source-workbook",
    "Source workbook",
    sourceWorkbookPresent ? "ok" : "error",
    sourceWorkbookPresent ? "found" : "missing",
  );
  add(
    "executed-workbook",
    "Executed workbook",
    executedWorkbookPresent ? "ok" : "warn",
    executedWorkbookPresent ? "found" : "not created yet",
  );

  const profilePath = config?.profileDir ?? path.join(projectRoot, ".whatsapp-profile");
  const profilePresent = await pathExists(profilePath);
  add(
    "whatsapp-profile",
    "WhatsApp profile",
    profilePresent ? "ok" : "error",
    profilePresent ? "present" : "missing; sign in before execution",
  );
  add(
    "whatsapp-target",
    "WhatsApp target",
    config?.target ? "ok" : "error",
    config?.target ? `configured by ${config.target.kind}` : "not configured",
  );

  if (!config) {
    add("drive", "Google Drive", "error", "configuration is invalid");
  } else if (!config.googleDriveEvidenceEnabled) {
    add("drive", "Google Drive", "warn", "disabled");
  } else {
    try {
      resolveGoogleServiceAccount(config.googleServiceAccount);
      if (!config.googleDriveEvidenceParentFolderId) {
        throw new Error("parent folder is not configured");
      }
      if (dependencies.checkDriveAccess !== false) {
        if (dependencies.validateDrive) {
          await dependencies.validateDrive(config);
        } else {
          await createGoogleDriveEvidencePublisher(config).validateParentFolder();
        }
        add("drive", "Google Drive", "ok", "credentials and folder access verified");
      } else {
        add("drive", "Google Drive", "ok", "configuration is valid; access not checked");
      }
    } catch (error) {
      add(
        "drive",
        "Google Drive",
        "error",
        safeGoogleCredentialError(error, config.googleServiceAccount?.value),
      );
    }
  }

  const browserRuntime = await detectBrowserRuntime({
    platform,
    environment,
    headless: config?.headless ?? false,
    hasCommand: dependencies.hasCommand ?? commandAvailable,
  });
  add(
    "browser-runtime",
    "Browser runtime",
    browserRuntime.mode === "unavailable" ? "error" : "ok",
    browserRuntime.reason,
  );

  return {
    checks,
    browserRuntime,
    chromiumInstalled,
    profilePresent,
    environmentFilePresent,
    ready: !checks.some((check) => check.status === "error"),
  };
}

export function formatDiagnosticReport(report: DiagnosticReport): string {
  const labels: Record<DiagnosticStatus, string> = {
    ok: "OK",
    warn: "WARN",
    error: "ERROR",
    info: "INFO",
  };
  return report.checks
    .map(
      (check) =>
        `${labels[check.status].padEnd(5)} ${check.label}: ${check.detail}`,
    )
    .join("\n");
}
