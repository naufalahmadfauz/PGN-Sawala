import { loadConfig } from "../src/config";
import { isEntrypoint, runCliMain } from "../src/cli-entrypoint";
import {
  createDiscordNotifier,
  safeDiscordError,
  validateDiscordWebhookUrl,
  type DiscordNotificationSettings,
  type DiscordNotifierDependencies,
  type DiscordRunProgressEvent,
} from "../src/notifications/discord";

const DEMO_SCENARIO_COUNT = 10;
const DEMO_DELAY_MS = 1_000;

export type DiscordDemoOutcome = "complete" | "fail" | "interrupt";

export interface DiscordDemoOptions {
  outcome: DiscordDemoOutcome;
}

export interface DiscordDemoDependencies {
  delay?: (milliseconds: number) => Promise<void>;
  log?: (message: string) => void;
  warn?: (message: string) => void;
  now?: () => Date;
  notifier?: DiscordNotifierDependencies;
}

export interface DiscordDemoResult {
  status:
    | "completed"
    | "simulated-failure"
    | "simulated-interruption"
    | "disabled"
    | "missing-webhook"
    | "invalid-webhook"
    | "invalid-settings"
    | "delivery-warning";
  runId?: string;
  warningCount: number;
}

function defaultDelay(milliseconds: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, milliseconds));
}

function demoRunId(startedAt: Date): string {
  const timestamp = startedAt
    .toISOString()
    .replaceAll("-", "")
    .replaceAll(":", "")
    .replace(/\.\d{3}Z$/, "Z");
  return `DEMO-${timestamp}`;
}

function demoProgress(
  completedScenarios: number,
  updatedAt: Date,
): DiscordRunProgressEvent {
  return {
    completedScenarios,
    totalScenarios: DEMO_SCENARIO_COUNT,
    currentScenarioId: `DEMO-KB-${String(completedScenarios).padStart(3, "0")}`,
    capturedScenarios: completedScenarios,
    timeouts: 0,
    technicalErrors: 0,
    evidenceUploaded: completedScenarios,
    evidenceUploadErrors: 0,
    updatedAt,
  };
}

export function parseDiscordDemoArgs(
  args: readonly string[],
): DiscordDemoOptions {
  const allowed = new Set(["--fail", "--interrupt"]);
  const unknown = args.filter((argument) => !allowed.has(argument));
  const failCount = args.filter((argument) => argument === "--fail").length;
  const interruptCount = args.filter(
    (argument) => argument === "--interrupt",
  ).length;
  if (
    unknown.length ||
    failCount > 1 ||
    interruptCount > 1 ||
    (failCount && interruptCount)
  ) {
    throw new Error(
      "Usage: npm run discord:demo -- [--fail | --interrupt]",
    );
  }
  return {
    outcome: failCount ? "fail" : interruptCount ? "interrupt" : "complete",
  };
}

export async function runDiscordDemo(
  config: DiscordNotificationSettings = loadConfig(),
  options: DiscordDemoOptions = { outcome: "complete" },
  dependencies: DiscordDemoDependencies = {},
): Promise<DiscordDemoResult> {
  const log = dependencies.log ?? console.log;
  const outputWarning = dependencies.warn ?? console.warn;
  if (!config.discordNotificationsEnabled) {
    log("Discord notifications are disabled.");
    log("Enable them in .env before running the demo.");
    return { status: "disabled", warningCount: 0 };
  }
  if (config.discordConfigurationIssues?.length) {
    log("Discord notification configuration is invalid.");
    log("Correct the Discord settings before running the demo.");
    return { status: "invalid-settings", warningCount: 0 };
  }

  const configuredWebhook = config.discordWebhookUrl?.trim();
  if (!configuredWebhook) {
    log("Discord webhook: not configured");
    log("Configure DISCORD_WEBHOOK_URL in .env before running the demo.");
    return { status: "missing-webhook", warningCount: 0 };
  }
  if (!validateDiscordWebhookUrl(configuredWebhook).valid) {
    log("Discord webhook: invalid configuration");
    log("Configure a valid DISCORD_WEBHOOK_URL before running the demo.");
    return { status: "invalid-webhook", warningCount: 0 };
  }

  const now = dependencies.now ?? (() => new Date());
  const delay = dependencies.delay ?? defaultDelay;
  const warnings: string[] = [];
  const notifier = createDiscordNotifier(
    {
      ...config,
      discordProgressEvery: 1,
      discordProgressMinutes: 1,
      discordNotifyStart: true,
      discordNotifyProgress: true,
      discordNotifyComplete: true,
      discordNotifyFailure: true,
    },
    {
      timeoutMs: 1_000,
      maxAttempts: 2,
      maxRetryDelayMs: 500,
      ...dependencies.notifier,
      now,
      minimumProgressIntervalMs: 0,
      warn: (message) => {
        warnings.push(message);
        outputWarning(message);
      },
    },
  );
  const startedAt = now();
  const runId = demoRunId(startedAt);

  log("Discord webhook: configured");
  log(`Demo run ID: ${runId}`);
  log("Demo mode: Discord Demo");
  log(`Selected scenarios: ${DEMO_SCENARIO_COUNT}`);
  log("Demo event: STARTED");
  await notifier.runStarted({
    runId,
    mode: "demo",
    selectedScenarios: DEMO_SCENARIO_COUNT,
    startedAt,
    googleDriveEvidenceEnabled: false,
    workbookPath: "Discord Demo.xlsx",
  });

  if (options.outcome === "complete") {
    for (const completedScenarios of [1, 5]) {
      await delay(DEMO_DELAY_MS);
      log(`Demo event: RUNNING ${completedScenarios} / ${DEMO_SCENARIO_COUNT}`);
      await notifier.runProgress(demoProgress(completedScenarios, now()));
    }
    await delay(DEMO_DELAY_MS);
    log(
      `Demo event: COMPLETED ${DEMO_SCENARIO_COUNT} / ${DEMO_SCENARIO_COUNT}`,
    );
    const completedAt = now();
    await notifier.runCompleted({
      ...demoProgress(DEMO_SCENARIO_COUNT, completedAt),
      completedAt,
    });
  } else {
    await delay(DEMO_DELAY_MS);
    log(`Demo event: RUNNING 5 / ${DEMO_SCENARIO_COUNT}`);
    const progress = demoProgress(5, now());
    await notifier.runProgress(progress);
    await delay(DEMO_DELAY_MS);
    if (options.outcome === "fail") {
      log(`Demo event: FAILED 5 / ${DEMO_SCENARIO_COUNT}`);
      const failedAt = now();
      await notifier.runFailed({
        ...progress,
        updatedAt: failedAt,
        failedAt,
        reason: "Simulated failure for Discord notification testing",
        workbookProgress: "Demo only; no workbook modified",
        evidenceProgress: "5 simulated uploads; no Drive action performed",
      });
    } else {
      log(`Demo event: INTERRUPTED 5 / ${DEMO_SCENARIO_COUNT}`);
      await notifier.runInterrupted("SIGINT", progress, {
        reason: "Simulated SIGINT",
        workbookProgress: "Demo only; no workbook modified",
        evidenceProgress: "5 simulated uploads; no Drive action performed",
      });
    }
  }

  if (warnings.length) {
    log(`Demo result: completed with ${warnings.length} delivery warning(s)`);
    return {
      status: "delivery-warning",
      runId,
      warningCount: warnings.length,
    };
  }
  if (options.outcome === "fail") {
    log("Demo result: simulated failure notification sent successfully");
    return { status: "simulated-failure", runId, warningCount: 0 };
  }
  if (options.outcome === "interrupt") {
    log("Demo result: simulated interruption notification sent successfully");
    return { status: "simulated-interruption", runId, warningCount: 0 };
  }
  log("Demo result: notification lifecycle sent successfully");
  return { status: "completed", runId, warningCount: 0 };
}

if (isEntrypoint(import.meta.url)) {
  runCliMain(async () => {
    const options = parseDiscordDemoArgs(process.argv.slice(2));
    const result = await runDiscordDemo(loadConfig(), options);
    if (
      result.status === "disabled" ||
      result.status === "missing-webhook" ||
      result.status === "invalid-webhook" ||
      result.status === "invalid-settings" ||
      result.status === "delivery-warning"
    ) {
      process.exitCode = 1;
    }
  }, safeDiscordError);
}
