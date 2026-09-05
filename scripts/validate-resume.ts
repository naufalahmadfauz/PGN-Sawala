import { loadConfig, type AppConfig } from "../src/config";
import { isEntrypoint, runCliMain } from "../src/cli-entrypoint";
import {
  formatRecoveryValidation,
  validateRecoveryRun,
  type RecoveryValidation,
  type RecoveryValidationDependencies,
} from "../src/recovery/recovery-service";

export function parseResumeValidationArgs(args: readonly string[]): {
  runId?: string;
} {
  if (args.length === 0) return {};
  if (args.length === 2 && args[0] === "--run" && args[1]?.trim()) {
    return { runId: args[1].trim() };
  }
  throw new Error("Usage: npm run test:pgn:resume:validate -- [--run <Run ID>]");
}

export async function validateResume(
  config: AppConfig = loadConfig(),
  runId?: string,
  dependencies: RecoveryValidationDependencies = {},
): Promise<RecoveryValidation> {
  return validateRecoveryRun(config, runId, dependencies);
}

if (isEntrypoint(import.meta.url)) {
  runCliMain(async () => {
    const options = parseResumeValidationArgs(process.argv.slice(2));
    const validation = await validateResume(loadConfig(), options.runId);
    console.log(formatRecoveryValidation(validation));
    if (!validation.ready) {
      throw new Error("Recovery validation is blocked; no testcase was executed");
    }
  });
}
