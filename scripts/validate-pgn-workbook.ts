import { loadConfig, type AppConfig } from "../src/config";
import { isEntrypoint, runCliMain } from "../src/cli-entrypoint";
import { loadPgnWorkbook } from "../src/excel/pgn-workbook-loader";
import { formatPgnValidation } from "../src/excel/pgn-workbook-validator";
import { formatWorkbookMappings, inspectWorkbookMappings } from "../src/operator/workbook-configuration";
import { parseCliOptions } from "../src/pgn-cli";
import type { SessionMode } from "../src/session-mode";

export function parsePgnValidationArgs(args: string[]): SessionMode {
  const options = parseCliOptions(args);
  for (let index = 0; index < args.length; index += 1) {
    const argument = args[index];
    if (argument === "--session") {
      index += 1;
    } else if (argument !== "--fast" && !argument.startsWith("--session=")) {
      throw new Error("Workbook validation accepts only --session=isolated, --session=continuous, --session isolated|continuous, or --fast; execution and recovery filters are not allowed.");
    }
  }
  return options.sessionMode;
}

export async function validatePgnWorkbook(
  config: AppConfig = loadConfig(),
  sessionMode: SessionMode = "isolated",
): Promise<boolean> {
  const inspection = await inspectWorkbookMappings(config, true);
  console.log(formatWorkbookMappings(inspection));
  const { parsed } = await loadPgnWorkbook(config.pgnSourceWorkbookPath);
  console.log(
    formatPgnValidation(parsed, {
      sessionMode,
      command: config.resetCommand,
      confirmation: config.resetConfirmation,
      timeoutMs: config.resetTimeoutMs,
      responseIdleMs: config.responseIdleMs,
      responseTimeoutMs: config.responseTimeoutMs,
      postResetQuietMs: config.postResetQuietMs,
    }),
  );
  return inspection.ready && !parsed.issues.some((issue) => issue.severity === "ERROR");
}

if (isEntrypoint(import.meta.url)) {
  runCliMain(async () => {
    const sessionMode = parsePgnValidationArgs(process.argv.slice(2));
    if (!(await validatePgnWorkbook(undefined, sessionMode))) {
      process.exitCode = 1;
    }
  });
}
