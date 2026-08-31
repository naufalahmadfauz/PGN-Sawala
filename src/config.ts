import path from "node:path";
import dotenv from "dotenv";

dotenv.config({ path: path.resolve(process.cwd(), ".env"), quiet: true });

export interface WhatsAppTarget {
  kind: "phone" | "chat";
  value: string;
}

export interface AppConfig {
  projectRoot: string;
  whatsappUrl: string;
  profileDir: string;
  artifactsDir: string;
  debugDir: string;
  evidenceDir: string;
  loginScreenshotPath: string;
  dataFilePath: string;
  reportFilePath: string;
  loginTimeoutMs: number;
  authTimeoutMs: number;
  responseTimeoutMs: number;
  responseIdleMs: number;
  betweenTestsMs: number;
  headless: boolean;
  browserChannel?: string;
  target?: WhatsAppTarget;
}

function integerFromEnvironment(
  name: string,
  fallback: number,
  minimum: number,
): number {
  const rawValue = process.env[name]?.trim();
  if (!rawValue) {
    return fallback;
  }

  const value = Number(rawValue);
  if (!Number.isInteger(value) || value < minimum) {
    throw new Error(`${name} must be an integer greater than or equal to ${minimum}`);
  }

  return value;
}

function booleanFromEnvironment(name: string, fallback: boolean): boolean {
  const rawValue = process.env[name]?.trim().toLowerCase();
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

function readTarget(): WhatsAppTarget | undefined {
  const phone = process.env.PGN_WHATSAPP_PHONE?.replace(/\D/g, "");
  if (phone) {
    if (phone.length < 8) {
      throw new Error("PGN_WHATSAPP_PHONE must include a valid international phone number");
    }
    return { kind: "phone", value: phone };
  }

  const chat = process.env.PGN_WHATSAPP_CHAT?.trim();
  return chat ? { kind: "chat", value: chat } : undefined;
}

export function loadConfig(): AppConfig {
  const projectRoot = process.cwd();
  const artifactsDir = path.resolve(projectRoot, "artifacts");
  const profileDir = path.resolve(projectRoot, ".whatsapp-profile");
  const configuredProfile = process.env.WHATSAPP_PROFILE_DIR?.trim();
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
    process.env.PGN_TEST_DATA_FILE?.trim() || "data/PGN_Test_Cases.xlsx",
    "PGN_TEST_DATA_FILE",
  );
  const reportFilePath = resolveInsideDirectory(
    projectRoot,
    "reports",
    process.env.PGN_TEST_REPORT_FILE?.trim() ||
      "reports/PGN_Test_Results.xlsx",
    "PGN_TEST_REPORT_FILE",
  );
  const responseTimeoutMs = integerFromEnvironment(
    "WHATSAPP_RESPONSE_TIMEOUT_MS",
    30_000,
    1_000,
  );
  const responseIdleMs = integerFromEnvironment(
    "WHATSAPP_RESPONSE_IDLE_MS",
    1_800,
    500,
  );
  if (responseIdleMs >= responseTimeoutMs) {
    throw new Error(
      "WHATSAPP_RESPONSE_IDLE_MS must be less than WHATSAPP_RESPONSE_TIMEOUT_MS",
    );
  }

  return {
    projectRoot,
    whatsappUrl: "https://web.whatsapp.com/",
    profileDir,
    artifactsDir,
    debugDir: path.join(artifactsDir, "debug"),
    evidenceDir: path.join(artifactsDir, "evidence"),
    loginScreenshotPath: path.join(artifactsDir, "whatsapp-login.png"),
    dataFilePath,
    reportFilePath,
    loginTimeoutMs: integerFromEnvironment(
      "WHATSAPP_LOGIN_TIMEOUT_MS",
      15 * 60_000,
      10_000,
    ),
    authTimeoutMs: integerFromEnvironment(
      "WHATSAPP_AUTH_TIMEOUT_MS",
      90_000,
      5_000,
    ),
    responseTimeoutMs,
    responseIdleMs,
    betweenTestsMs: integerFromEnvironment(
      "WHATSAPP_BETWEEN_TESTS_MS",
      3_000,
      0,
    ),
    headless: booleanFromEnvironment("WHATSAPP_HEADLESS", false),
    browserChannel:
      process.env.WHATSAPP_BROWSER_CHANNEL?.trim() || undefined,
    target: readTarget(),
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
