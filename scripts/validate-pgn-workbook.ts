import { loadConfig, type AppConfig } from "../src/config";
import { isEntrypoint, runCliMain } from "../src/cli-entrypoint";
import { loadPgnWorkbook } from "../src/excel/pgn-workbook-loader";
import { formatPgnValidation } from "../src/excel/pgn-workbook-validator";

export async function validatePgnWorkbook(
  config: AppConfig = loadConfig(),
): Promise<boolean> {
  const { parsed } = await loadPgnWorkbook(config.pgnSourceWorkbookPath);
  console.log(
    formatPgnValidation(parsed, {
      command: config.resetCommand,
      confirmation: config.resetConfirmation,
      timeoutMs: config.resetTimeoutMs,
      responseIdleMs: config.responseIdleMs,
      responseTimeoutMs: config.responseTimeoutMs,
      postResetQuietMs: config.postResetQuietMs,
    }),
  );
  return !parsed.issues.some((issue) => issue.severity === "ERROR");
}

if (isEntrypoint(import.meta.url)) {
  runCliMain(async () => {
    if (!(await validatePgnWorkbook())) {
      process.exitCode = 1;
    }
  });
}
