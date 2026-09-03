import path from "node:path";
import { loadConfig } from "../config";
import { REPOSITORY_ROOT } from "../environment";
import { runPgnWorkbook } from "../pgn-runner";
import {
  loginWhatsApp,
  recreateWhatsAppAuthentication,
  verifyWhatsApp,
} from "../whatsapp/auth";
import { migrateEvidence } from "../../scripts/evidence-migrate";
import { validateEvidence } from "../../scripts/evidence-validate";
import { prepareFreshPgnWorkbook } from "../../scripts/fresh-pgn";
import { validatePgnWorkbook } from "../../scripts/validate-pgn-workbook";
import { validateRetest } from "../../scripts/validate-retest";
import { runBrowserAction } from "./browser-runtime";
import type { OperatorActions } from "./control-panel";
import { collectDiagnostics, formatDiagnosticReport } from "./diagnostics";
import { inspectPgnExecution } from "./pgn-preflight";
import { runInheritedCommand } from "./process";
import { runSetupWizard } from "./setup";
import type { OperatorUi } from "./ui";

const SAFE_TEST_FILES = [
  "scripts/response-collector.test.ts",
  "scripts/session-reset.test.ts",
  "scripts/workbook-writer.test.ts",
  "scripts/config.test.ts",
  "scripts/retest.test.ts",
  "scripts/evidence-migration.test.ts",
  "scripts/operator.test.ts",
];

function scriptPath(name: string): string {
  return path.join(REPOSITORY_ROOT, "scripts", name);
}

async function browserAction(
  entrypoint: string,
  args: string[],
  direct: () => Promise<void>,
  browserRequired?: () => Promise<boolean>,
): Promise<void> {
  const config = loadConfig();
  await runBrowserAction({
    headless: config.headless,
    projectRoot: config.projectRoot,
    scriptPath: scriptPath(entrypoint),
    args,
    direct,
    browserRequired,
  });
}

export function createDefaultActions(ui: OperatorUi): OperatorActions {
  const login = (): Promise<void> =>
    browserAction("whatsapp-login.ts", [], () => loginWhatsApp());

  return {
    validatePgn: validatePgnWorkbook,
    prepareFresh: prepareFreshPgnWorkbook,
    runPgn: (args) =>
      browserAction(
        "run-pgn.ts",
        args,
        () => runPgnWorkbook(args, "full"),
        async () =>
          (await inspectPgnExecution(args, "full", loadConfig()))
            .browserRequired,
      ),
    validateRetest,
    runRetest: (args) =>
      browserAction(
        "retest-pgn.ts",
        args,
        () => runPgnWorkbook(args, "retest"),
        async () =>
          (await inspectPgnExecution(args, "retest", loadConfig()))
            .browserRequired,
      ),
    validateEvidence,
    migrateEvidence,
    loginWhatsApp: login,
    verifyWhatsApp: () =>
      browserAction("whatsapp-verify.ts", [], () => verifyWhatsApp()),
    recreateWhatsApp: () =>
      browserAction("whatsapp-recreate.ts", [], () =>
        recreateWhatsAppAuthentication(),
      ),
    setup: async () => {
      const result = await runSetupWizard(ui, { loginWhatsApp: login });
      return result.nextAction;
    },
    diagnostics: async () => {
      const report = await ui.task(
        "Checking operator prerequisites",
        () => collectDiagnostics(),
        "Diagnostics complete",
      );
      ui.note(formatDiagnosticReport(report), "Diagnostics");
      if (report.ready) ui.success("Required operator checks passed");
      else ui.warn("Resolve the reported errors before test execution");
    },
    typecheck: () =>
      runInheritedCommand(
        process.execPath,
        [path.join(REPOSITORY_ROOT, "node_modules", "typescript", "bin", "tsc"), "--noEmit"],
        { cwd: REPOSITORY_ROOT },
      ),
    regressionTests: () =>
      runInheritedCommand(
        process.execPath,
        ["--import", "tsx", "--test", ...SAFE_TEST_FILES],
        { cwd: REPOSITORY_ROOT },
      ),
    createTemplate: () =>
      runInheritedCommand(
        process.execPath,
        ["--import", "tsx", scriptPath("create-test-template.ts")],
        { cwd: REPOSITORY_ROOT },
      ),
  };
}
