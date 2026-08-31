import { loadConfig } from "../src/config";
import { loadPgnWorkbook } from "../src/excel/pgn-workbook-loader";
import { formatPgnValidation } from "../src/excel/pgn-workbook-validator";

const config = loadConfig();
const { parsed } = await loadPgnWorkbook(config.pgnSourceWorkbookPath);
console.log(
  formatPgnValidation(parsed, {
    command: config.resetCommand,
    confirmation: config.resetConfirmation,
    timeoutMs: config.resetTimeoutMs,
  }),
);
if (parsed.issues.some((issue) => issue.severity === "ERROR")) {
  process.exitCode = 1;
}
