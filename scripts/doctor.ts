import { runCliMain } from "../src/cli-entrypoint";
import {
  collectDiagnostics,
  formatDiagnosticReport,
} from "../src/operator/diagnostics";
import { createClackUi } from "../src/operator/ui";

const ui = createClackUi();
runCliMain(async () => {
  ui.intro("PGN Sawala diagnostics");
  const report = await ui.task(
    "Checking operator prerequisites",
    () => collectDiagnostics(),
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
