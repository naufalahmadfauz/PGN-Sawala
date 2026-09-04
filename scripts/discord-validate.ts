import { loadConfig, type AppConfig } from "../src/config";
import { isEntrypoint, runCliMain } from "../src/cli-entrypoint";
import {
  discordStatusLines,
  safeDiscordError,
  validateDiscordWebhook,
  type DiscordValidationOptions,
  type DiscordValidationResult,
} from "../src/notifications/discord";

export async function runDiscordValidation(
  config: AppConfig = loadConfig(),
  options: DiscordValidationOptions = {},
): Promise<DiscordValidationResult> {
  return validateDiscordWebhook(config, options);
}

export function parseDiscordValidationArgs(
  args: readonly string[],
): { sendTest: boolean } {
  const unknown = args.filter((argument) => argument !== "--send-test");
  if (
    unknown.length ||
    args.filter((argument) => argument === "--send-test").length > 1
  ) {
    throw new Error("Usage: npm run discord:validate -- [--send-test]");
  }
  return { sendTest: args.includes("--send-test") };
}

if (isEntrypoint(import.meta.url)) {
  runCliMain(async () => {
    const { sendTest } = parseDiscordValidationArgs(process.argv.slice(2));
    const config = loadConfig();
    const result = await runDiscordValidation(config, { sendTest });
    console.log(discordStatusLines(config, result));
    if (result.reason) {
      console.log(
        `Reason .......... ${safeDiscordError(new Error(result.reason), config.discordWebhookUrl)}`,
      );
    }
    const testStatus = !sendTest
      ? "NOT REQUESTED"
      : result.testNotificationSent
        ? "SENT"
        : result.testNotificationDeliveryUncertain
          ? "DELIVERY UNKNOWN; CHECK THE CHANNEL BEFORE RETRYING"
          : "NOT SENT";
    console.log(`Test notification ${testStatus}`);
    if (!result.valid || result.connectivity !== "ok") {
      process.exitCode = 1;
    }
  }, safeDiscordError);
}
