import path from "node:path";
import {
  GOOGLE_SERVICE_ACCOUNT_SOURCES,
  type GoogleServiceAccountConfiguration,
} from "./evidence/google-service-account";
import {
  REPOSITORY_ROOT,
  loadEnvironment,
  summarizeEnvironmentSources,
  type EnvironmentConfigurationSource,
} from "./environment";

export interface WhatsAppTarget {
  kind: "phone" | "chat";
  value: string;
}

export interface AppConfig {
  projectRoot: string;
  environmentFilePath: string;
  environmentFileLoaded: boolean;
  whatsappUrl: string;
  profileDir: string;
  artifactsDir: string;
  debugDir: string;
  evidenceDir: string;
  evidenceCleanDir: string;
  reportArchiveDir: string;
  loginScreenshotPath: string;
  dataFilePath: string;
  reportFilePath: string;
  pgnSourceWorkbookPath: string;
  pgnExecutedWorkbookPath: string;
  loginTimeoutMs: number;
  authTimeoutMs: number;
  responseTimeoutMs: number;
  responseIdleMs: number;
  resetCommand: string;
  resetConfirmation: string;
  resetTimeoutMs: number;
  postResetQuietMs: number;
  betweenTestsMs: number;
  headless: boolean;
  browserChannel?: string;
  googleDriveEvidenceEnabled: boolean;
  googleDriveEvidenceParentFolderId?: string;
  googleDriveEvidenceFolderPrefix: string;
  googleDriveRetestFolderPrefix: string;
  googleDriveConfigurationSource: EnvironmentConfigurationSource;
  googleServiceAccount?: GoogleServiceAccountConfiguration;
  legacyEvidenceCropLeft?: number;
  discordNotificationsEnabled: boolean;
  discordWebhookUrl?: string;
  discordProgressEvery: number;
  discordProgressMinutes: number;
  discordNotifyStart: boolean;
  discordNotifyProgress: boolean;
  discordNotifyComplete: boolean;
  discordNotifyFailure: boolean;
  discordConfigurationIssues: string[];
  target?: WhatsAppTarget;
}

function integerFromEnvironment(
  environment: NodeJS.ProcessEnv,
  name: string,
  fallback: number,
  minimum: number,
): number {
  const rawValue = environment[name]?.trim();
  if (!rawValue) {
    return fallback;
  }

  const value = Number(rawValue);
  if (!Number.isInteger(value) || value < minimum) {
    throw new Error(`${name} must be an integer greater than or equal to ${minimum}`);
  }

  return value;
}

function booleanFromEnvironment(
  environment: NodeJS.ProcessEnv,
  name: string,
  fallback: boolean,
): boolean {
  const rawValue = environment[name]?.trim().toLowerCase();
  if (!rawValue) {
    return fallback;
  }
  if (rawValue === "true" || rawValue === "1") {
    return true;
  }
  if (rawValue === "false" || rawValue === "0") {
    return false;
  }
  throw new Error(`${name} must be true, false, 1, or 0`);
}

function textFromEnvironment(
  environment: NodeJS.ProcessEnv,
  name: string,
  fallback: string,
): string {
  return environment[name]?.trim() || fallback;
}

function optionalBooleanFromEnvironment(
  environment: NodeJS.ProcessEnv,
  name: string,
  fallback: boolean,
  invalidFallback = fallback,
): boolean {
  try {
    return booleanFromEnvironment(environment, name, fallback);
  } catch {
    return invalidFallback;
  }
}

function optionalIntegerFromEnvironment(
  environment: NodeJS.ProcessEnv,
  name: string,
  fallback: number,
): number {
  try {
    return integerFromEnvironment(environment, name, fallback, 1);
  } catch {
    return fallback;
  }
}

function validOptionalPositiveInteger(
  environment: NodeJS.ProcessEnv,
  name: string,
): boolean {
  const rawValue = environment[name]?.trim();
  if (!rawValue) return true;
  const value = Number(rawValue);
  return Number.isInteger(value) && value >= 1;
}

function validOptionalBoolean(
  environment: NodeJS.ProcessEnv,
  name: string,
): boolean {
  const rawValue = environment[name]?.trim().toLowerCase();
  return (
    !rawValue ||
    rawValue === "true" ||
    rawValue === "false" ||
    rawValue === "1" ||
    rawValue === "0"
  );
}

export function normalizeGoogleDriveFolderId(value: string): string {
  const trimmed = value.trim();
  if (!trimmed) {
    throw new Error("GOOGLE_DRIVE_EVIDENCE_PARENT_FOLDER must not be empty");
  }

  let folderId = trimmed;
  if (/^https?:\/\//i.test(trimmed)) {
    let url: URL;
    try {
      url = new URL(trimmed);
    } catch {
      throw new Error("GOOGLE_DRIVE_EVIDENCE_PARENT_FOLDER is not a valid URL");
    }
    if (url.hostname !== "drive.google.com") {
      throw new Error(
        "GOOGLE_DRIVE_EVIDENCE_PARENT_FOLDER URL must use drive.google.com",
      );
    }
    const match = url.pathname.match(/\/folders\/([A-Za-z0-9_-]+)/);
    if (!match) {
      throw new Error(
        "GOOGLE_DRIVE_EVIDENCE_PARENT_FOLDER URL must contain /folders/<ID>",
      );
    }
    folderId = match[1];
  }

  if (!/^[A-Za-z0-9_-]{10,}$/.test(folderId)) {
    throw new Error(
      "GOOGLE_DRIVE_EVIDENCE_PARENT_FOLDER must be a Drive folder ID or folder URL",
    );
  }
  return folderId;
}

function resolveInsideDirectory(
  projectRoot: string,
  directory: string,
  value: string,
  environmentName: string,
): string {
  const allowedDirectory = path.join(projectRoot, directory);
  const resolved = path.resolve(projectRoot, value);
  const relative = path.relative(allowedDirectory, resolved);
  if (!relative || relative.startsWith("..") || path.isAbsolute(relative)) {
    throw new Error(
      `${environmentName} must point to a file inside the ${directory}/ directory`,
    );
  }
  if (path.extname(resolved).toLowerCase() !== ".xlsx") {
    throw new Error(`${environmentName} must point to an .xlsx file`);
  }
  return resolved;
}

function readTarget(environment: NodeJS.ProcessEnv): WhatsAppTarget | undefined {
  const phone = environment.PGN_WHATSAPP_PHONE?.replace(/\D/g, "");
  if (phone) {
    if (phone.length < 8) {
      throw new Error("PGN_WHATSAPP_PHONE must include a valid international phone number");
    }
    return { kind: "phone", value: phone };
  }

  const chat = environment.PGN_WHATSAPP_CHAT?.trim();
  return chat ? { kind: "chat", value: chat } : undefined;
}

export interface LoadConfigOptions {
  repositoryRoot?: string;
  environment?: NodeJS.ProcessEnv;
}

export function loadConfig(options: LoadConfigOptions = {}): AppConfig {
  const projectRoot = path.resolve(options.repositoryRoot ?? REPOSITORY_ROOT);
  const loadedEnvironment = loadEnvironment({
    repositoryRoot: projectRoot,
    values: options.environment,
  });
  const environment = loadedEnvironment.values;
  const artifactsDir = path.resolve(projectRoot, "artifacts");
  const profileDir = path.resolve(projectRoot, ".whatsapp-profile");
  const configuredProfile = environment.WHATSAPP_PROFILE_DIR?.trim();
  if (
    configuredProfile &&
    path.resolve(projectRoot, configuredProfile) !== profileDir
  ) {
    throw new Error(
      "WHATSAPP_PROFILE_DIR is restricted to .whatsapp-profile so authentication data remains gitignored",
    );
  }
  const dataFilePath = resolveInsideDirectory(
    projectRoot,
    "data",
    environment.PGN_TEST_DATA_FILE?.trim() || "data/PGN_Test_Cases.xlsx",
    "PGN_TEST_DATA_FILE",
  );
  const reportFilePath = resolveInsideDirectory(
    projectRoot,
    "reports",
    environment.PGN_TEST_REPORT_FILE?.trim() ||
      "reports/PGN_Test_Results.xlsx",
    "PGN_TEST_REPORT_FILE",
  );
  const pgnSourceWorkbookPath = resolveInsideDirectory(
    projectRoot,
    "data",
    environment.PGN_SOURCE_WORKBOOK?.trim() ||
      "data/PGN AI Assistant - Knowledge Base Testing Report - User Inputs.xlsx",
    "PGN_SOURCE_WORKBOOK",
  );
  const pgnExecutedWorkbookPath = resolveInsideDirectory(
    projectRoot,
    "reports",
    environment.PGN_EXECUTED_WORKBOOK?.trim() ||
      "reports/PGN AI Assistant - Knowledge Base Testing Report - Executed.xlsx",
    "PGN_EXECUTED_WORKBOOK",
  );
  const responseTimeoutMs = integerFromEnvironment(
    environment,
    "WHATSAPP_RESPONSE_TIMEOUT_MS",
    60_000,
    1_000,
  );
  const responseIdleMs = integerFromEnvironment(
    environment,
    "WHATSAPP_RESPONSE_IDLE_MS",
    10_000,
    500,
  );
  if (responseIdleMs >= responseTimeoutMs) {
    throw new Error(
      "WHATSAPP_RESPONSE_IDLE_MS must be less than WHATSAPP_RESPONSE_TIMEOUT_MS",
    );
  }
  const configuredDriveParent =
    environment.GOOGLE_DRIVE_EVIDENCE_PARENT_FOLDER?.trim();
  const configuredCropLeft = environment.LEGACY_EVIDENCE_CROP_LEFT?.trim();
  const legacyEvidenceCropLeft = configuredCropLeft
    ? integerFromEnvironment(environment, "LEGACY_EVIDENCE_CROP_LEFT", 0, 1)
    : undefined;
  const googleDriveEvidenceFolderPrefix = textFromEnvironment(
    environment,
    "GOOGLE_DRIVE_EVIDENCE_FOLDER_PREFIX",
    "PGN-WhatsApp-Evidence",
  );
  const googleDriveRetestFolderPrefix = textFromEnvironment(
    environment,
    "GOOGLE_DRIVE_RETEST_FOLDER_PREFIX",
    "PGN-WhatsApp-Retest",
  );
  for (const [name, value] of [
    ["GOOGLE_DRIVE_EVIDENCE_FOLDER_PREFIX", googleDriveEvidenceFolderPrefix],
    ["GOOGLE_DRIVE_RETEST_FOLDER_PREFIX", googleDriveRetestFolderPrefix],
  ]) {
    if (value.includes("/") || value.length > 120) {
      throw new Error(
        `${name} must not contain / and must be at most 120 characters`,
      );
    }
  }

  const configuredCredentialSources = GOOGLE_SERVICE_ACCOUNT_SOURCES.filter(
    (source) => Boolean(environment[source]?.trim()),
  );
  const credentialSource = configuredCredentialSources.at(0);
  let googleServiceAccount: GoogleServiceAccountConfiguration | undefined;
  if (credentialSource) {
    const value = environment[credentialSource]!.trim();
    googleServiceAccount = {
      source: credentialSource,
      configuredSources: configuredCredentialSources,
      environmentSource: loadedEnvironment.sourceFor(credentialSource),
      ...(credentialSource === "GOOGLE_SERVICE_ACCOUNT_FILE"
        ? {
            filePath: path.resolve(projectRoot, value),
            fileDisplayPath: path.isAbsolute(value)
              ? path.normalize(value)
              : value.replaceAll("\\", "/"),
          }
        : { value }),
    };
  }
  const discordProgressEvery = optionalIntegerFromEnvironment(
    environment,
    "DISCORD_PROGRESS_EVERY",
    5,
  );
  const discordProgressMinutes = optionalIntegerFromEnvironment(
    environment,
    "DISCORD_PROGRESS_MINUTES",
    2,
  );
  const discordProgressIntervalsValid =
    validOptionalPositiveInteger(environment, "DISCORD_PROGRESS_EVERY") &&
    validOptionalPositiveInteger(environment, "DISCORD_PROGRESS_MINUTES");
  const discordConfigurationIssues = [
    ...[
      "DISCORD_NOTIFICATIONS_ENABLED",
      "DISCORD_NOTIFY_START",
      "DISCORD_NOTIFY_PROGRESS",
      "DISCORD_NOTIFY_COMPLETE",
      "DISCORD_NOTIFY_FAILURE",
    ]
      .filter((name) => !validOptionalBoolean(environment, name))
      .map((name) => `${name} must be true, false, 1, or 0`),
    ...["DISCORD_PROGRESS_EVERY", "DISCORD_PROGRESS_MINUTES"]
      .filter((name) => !validOptionalPositiveInteger(environment, name))
      .map((name) => `${name} must be a whole number of 1 or greater`),
  ];
  const googleDriveEvidenceEnabled = booleanFromEnvironment(
    environment,
    "GOOGLE_DRIVE_EVIDENCE_ENABLED",
    false,
  );
  const driveConfigurationKeys: string[] = [
    "GOOGLE_DRIVE_EVIDENCE_ENABLED",
    "GOOGLE_DRIVE_EVIDENCE_PARENT_FOLDER",
  ];
  if (credentialSource) {
    driveConfigurationKeys.push(credentialSource);
  }

  return {
    projectRoot,
    environmentFilePath: loadedEnvironment.envFilePath,
    environmentFileLoaded: loadedEnvironment.envFileLoaded,
    whatsappUrl: "https://web.whatsapp.com/",
    profileDir,
    artifactsDir,
    debugDir: path.join(artifactsDir, "debug"),
    evidenceDir: path.join(artifactsDir, "evidence"),
    evidenceCleanDir: path.join(artifactsDir, "evidence", "clean"),
    reportArchiveDir: path.join(projectRoot, "reports", "archive"),
    loginScreenshotPath: path.join(artifactsDir, "whatsapp-login.png"),
    dataFilePath,
    reportFilePath,
    pgnSourceWorkbookPath,
    pgnExecutedWorkbookPath,
    loginTimeoutMs: integerFromEnvironment(
      environment,
      "WHATSAPP_LOGIN_TIMEOUT_MS",
      15 * 60_000,
      10_000,
    ),
    authTimeoutMs: integerFromEnvironment(
      environment,
      "WHATSAPP_AUTH_TIMEOUT_MS",
      90_000,
      5_000,
    ),
    responseTimeoutMs,
    responseIdleMs,
    resetCommand: textFromEnvironment(environment, "PGN_RESET_COMMAND", "reset"),
    resetConfirmation: textFromEnvironment(
      environment,
      "PGN_RESET_CONFIRMATION",
      "Session deleted",
    ),
    resetTimeoutMs: integerFromEnvironment(
      environment,
      "PGN_RESET_TIMEOUT_MS",
      30_000,
      1_000,
    ),
    postResetQuietMs: integerFromEnvironment(
      environment,
      "POST_RESET_QUIET_MS",
      10_000,
      1,
    ),
    betweenTestsMs: integerFromEnvironment(
      environment,
      "WHATSAPP_BETWEEN_TESTS_MS",
      3_000,
      0,
    ),
    headless: booleanFromEnvironment(environment, "WHATSAPP_HEADLESS", false),
    browserChannel:
      environment.WHATSAPP_BROWSER_CHANNEL?.trim() || undefined,
    googleDriveEvidenceEnabled,
    googleDriveEvidenceParentFolderId: configuredDriveParent
      ? normalizeGoogleDriveFolderId(configuredDriveParent)
      : undefined,
    googleDriveEvidenceFolderPrefix,
    googleDriveRetestFolderPrefix,
    googleDriveConfigurationSource: summarizeEnvironmentSources(
      loadedEnvironment,
      driveConfigurationKeys,
    ),
    googleServiceAccount,
    legacyEvidenceCropLeft,
    discordNotificationsEnabled: optionalBooleanFromEnvironment(
      environment,
      "DISCORD_NOTIFICATIONS_ENABLED",
      false,
    ),
    discordWebhookUrl: environment.DISCORD_WEBHOOK_URL?.trim() || undefined,
    discordProgressEvery,
    discordProgressMinutes,
    discordNotifyStart: optionalBooleanFromEnvironment(
      environment,
      "DISCORD_NOTIFY_START",
      true,
      false,
    ),
    discordNotifyProgress:
      discordProgressIntervalsValid &&
      optionalBooleanFromEnvironment(
        environment,
        "DISCORD_NOTIFY_PROGRESS",
        true,
        false,
      ),
    discordNotifyComplete: optionalBooleanFromEnvironment(
      environment,
      "DISCORD_NOTIFY_COMPLETE",
      true,
      false,
    ),
    discordNotifyFailure: optionalBooleanFromEnvironment(
      environment,
      "DISCORD_NOTIFY_FAILURE",
      true,
      false,
    ),
    discordConfigurationIssues,
    target: readTarget(environment),
  };
}

export function requireTarget(config: AppConfig): WhatsAppTarget {
  if (!config.target) {
    throw new Error(
      "PGN target is not configured. Set PGN_WHATSAPP_PHONE or PGN_WHATSAPP_CHAT in .env.",
    );
  }
  return config.target;
}
