import path from "node:path";
import { loadConfig } from "../config";
import { REPOSITORY_ROOT } from "../environment";
import { validateDiscordWebhook } from "../notifications/discord";
import { parseCliOptions } from "../pgn-cli";
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
import { validateRest } from "../../scripts/validate-rest";
import { runBrowserAction } from "./browser-runtime";
import type { OperatorActions } from "./control-panel";
import { collectDiagnostics, formatDiagnosticReport } from "./diagnostics";
import { inspectPgnExecution } from "./pgn-preflight";
import { runInheritedCommand } from "./process";
import {
  configureDiscordNotifications,
  runSetupWizard,
} from "./setup";
import type { OperatorUi } from "./ui";
import { ensureReviewedWorkbookMapping, formatWorkbookMappings, inspectWorkbookMappings, reviewWorkbookMapping } from "./workbook-configuration";
import {
  abandonRecoveryRun,
  repairRecoveryProgress,
  skipRecoveryScenario,
  validateRecoveryRun,
} from "../recovery/recovery-service";
import {
  assertRecoveryRunExecutable,
  discoverRecoveryRun,
  readRecoveryRun,
} from "../recovery/run-state";

const SAFE_TEST_FILES = [
  "scripts/response-collector.test.ts",
  "scripts/session-reset.test.ts",
  "scripts/session-mode.test.ts",
  "scripts/rest.test.ts",
  "scripts/workbook-writer.test.ts",
  "scripts/config.test.ts",
  "scripts/retest.test.ts",
  "scripts/evidence-migration.test.ts",
  "scripts/operator.test.ts",
  "scripts/discord.test.ts",
  "scripts/recovery.test.ts",
  "scripts/recovery-demo.test.ts",
  "scripts/workbook-schema.test.ts",
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
    workbookSchema: async (review, redetect = false) => {
      const config = loadConfig();
      if (review) { await reviewWorkbookMapping(ui, config); return; }
      const inspection = await inspectWorkbookMappings(config, redetect);
      ui.note(formatWorkbookMappings(inspection), "Workbook schema");
    },
    inspectRecovery: () => discoverRecoveryRun(REPOSITORY_ROOT),
    validateRecovery: async (runId) => {
      const { state } = await readRecoveryRun(REPOSITORY_ROOT, runId);
      return validateRecoveryRun(loadConfig({ transport: state.transport }), runId);
    },
    resumeRecovery: async (runId, acceptSourceDrift = false) => {
      const { state } = await readRecoveryRun(REPOSITORY_ROOT, runId);
      const config = loadConfig({ transport: state.transport });
      assertRecoveryRunExecutable(state);
      const mode = state.mode;
      const entrypoint = mode === "retest" ? "retest-pgn.ts" : "run-pgn.ts";
      const args = [
        "--resume",
        runId,
        ...(acceptSourceDrift ? ["--accept-source-drift"] : []),
      ];
      if (state.transport === "rest") {
        await runPgnWorkbook(["--transport=rest", ...args], mode);
        return;
      }
      await browserAction(
        entrypoint,
        args,
        () => runPgnWorkbook(args, mode),
        async () => (await inspectPgnExecution(args, mode, config)).browserRequired,
      );
    },
    restartRecovery: async (runId, acceptSourceDrift = false) => {
      const { state } = await readRecoveryRun(REPOSITORY_ROOT, runId);
      const config = loadConfig({ transport: state.transport });
      assertRecoveryRunExecutable(state);
      const mode = state.mode;
      const args = [
        "--restart-run",
        runId,
        ...(acceptSourceDrift ? ["--accept-source-drift"] : []),
      ];
      if (state.transport === "rest") {
        await runPgnWorkbook(["--transport=rest", ...args], mode);
        return;
      }
      await browserAction(
        mode === "retest" ? "retest-pgn.ts" : "run-pgn.ts",
        args,
        () => runPgnWorkbook(args, mode),
        async () => (await inspectPgnExecution(args, mode, config)).browserRequired,
      );
    },
    skipRecoveryScenario: async (runId) => {
      const { state } = await readRecoveryRun(REPOSITORY_ROOT, runId);
      return skipRecoveryScenario(loadConfig({ transport: state.transport }), runId);
    },
    repairRecovery: async (runId, strategy) => {
      const { state } = await readRecoveryRun(REPOSITORY_ROOT, runId);
      return repairRecoveryProgress(loadConfig({ transport: state.transport }), runId, strategy);
    },
    abandonRecovery: async (runId) => {
      const { state } = await readRecoveryRun(REPOSITORY_ROOT, runId);
      return abandonRecoveryRun(loadConfig({ transport: state.transport }), runId);
    },
    validatePgn: validatePgnWorkbook,
    validateRest: (args) => validateRest(loadConfig({ transport: "rest" }), args),
    prepareFresh: prepareFreshPgnWorkbook,
    runPgn: async (args) => {
      if (!(await ensureReviewedWorkbookMapping(ui, loadConfig({ transport: parseCliOptions(args).transport })))) {
        throw new Error("Workbook mapping review cancelled; no testcase was executed.");
      }
      if (parseCliOptions(args).transport === "rest") {
        await runPgnWorkbook(args, "full");
        return;
      }
      await browserAction(
        "run-pgn.ts",
        args,
        () => runPgnWorkbook(args, "full"),
        async () =>
          (await inspectPgnExecution(args, "full", loadConfig()))
            .browserRequired,
      );
    },
    validateRetest,
    runRetest: async (args) => {
      if (!(await ensureReviewedWorkbookMapping(ui, loadConfig({ transport: parseCliOptions(args).transport })))) {
        throw new Error("Workbook mapping review cancelled; no testcase was executed.");
      }
      if (parseCliOptions(args).transport === "rest") {
        await runPgnWorkbook(args, "retest");
        return;
      }
      await browserAction(
        "retest-pgn.ts",
        args,
        () => runPgnWorkbook(args, "retest"),
        async () =>
          (await inspectPgnExecution(args, "retest", loadConfig()))
            .browserRequired,
      );
    },
    validateEvidence,
    migrateEvidence,
    validateDiscord: (sendTest) =>
      validateDiscordWebhook(loadConfig(), { sendTest }),
    configureNotifications: async () => {
      await configureDiscordNotifications(ui);
    },
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
