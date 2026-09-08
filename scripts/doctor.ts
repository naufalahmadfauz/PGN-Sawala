import { runCliMain } from "../src/cli-entrypoint";
import {
  collectDiagnostics,
  formatDiagnosticReport,
} from "../src/operator/diagnostics";
import { createClackUi } from "../src/operator/ui";
import { readExecutionTransport } from "../src/session-mode";

const ui = createClackUi();
runCliMain(async () => {
  const args = process.argv.slice(2);
  if (args.some((arg) => !arg.startsWith("--transport=") && arg !== "--check-rest")) throw new Error("Usage: npm run doctor -- [--transport=whatsapp|rest] [--check-rest]");
  const transport = readExecutionTransport(args.find((arg) => arg.startsWith("--transport="))?.split("=")[1]);
  ui.intro("PGN Sawala diagnostics");
  const report = await ui.task(
    "Checking operator prerequisites",
    () => collectDiagnostics({ transport, checkRestAccess: args.includes("--check-rest") }),
    "Diagnostics complete",
  );
  ui.note(formatDiagnosticReport(report), "Diagnostics");
  if (report.ready) {
    ui.outro("Required operator checks passed");
  } else {
    ui.outro("Resolve the reported errors before test execution");
    process.exitCode = 1;
  }
});
