import assert from "node:assert/strict";
import type { AppConfig } from "../../src/config";
import { runControlPanel, type OperatorActions } from "../../src/operator/control-panel";
import { inspectPgnExecution } from "../../src/operator/pgn-preflight";
import type { OperatorUi } from "../../src/operator/ui";
import { ensureReviewedWorkbookMapping } from "../../src/operator/workbook-configuration";
import { runPgnWorkbook, type PgnExecutionMode } from "../../src/pgn-runner";
import { validateRetest } from "../validate-retest";

// Exercise the menu's public action seam with real workbook handlers and an
// isolated project configuration; transport boundaries are supplied by each test.
export async function runWorkbookMenu(config: AppConfig, answers: unknown[], expectedErrors: RegExp[] = []) {
  const events: string[] = [];
  const record = (message: string) => { events.push(message); };
  const ui: OperatorUi = {
    intro: record, outro: record, cancel: record, info: record,
    success: record, warn: record, error: (message) => {
      const expected = expectedErrors.shift();
      assert(expected, message);
      assert.match(message, expected);
      record(message);
    },
    note: record,
    select: async (prompt) => {
      record(prompt.message);
      assert(answers.length, `Missing answer for ${prompt.message}`);
      return answers.shift() as never;
    },
    confirm: async (prompt) => {
      record(prompt.message);
      assert(answers.length, `Missing answer for ${prompt.message}`);
      return answers.shift() as boolean;
    },
    text: async () => answers.shift() as string,
    secret: async () => { throw new Error("No credentials should be prompted for"); },
    task: async (_, operation) => operation(),
  };
  const execute = async (args: string[], mode: PgnExecutionMode) => {
    if (!(await ensureReviewedWorkbookMapping(ui, config))) throw new Error("Workbook mapping review cancelled");
    await inspectPgnExecution(args, mode, config);
    await runPgnWorkbook(args, mode, config);
  };
  const unexpected = async (): Promise<never> => { throw new Error("Unexpected menu action"); };
  const actions: OperatorActions = {
    runPgn: (args: string[]) => execute(args, "full"),
    runRetest: (args: string[]) => execute(args, "retest"),
    validateRetest: (args: string[]) => validateRetest(args, config),
    validatePgn: unexpected, prepareFresh: unexpected, validateEvidence: unexpected,
    migrateEvidence: unexpected, validateDiscord: unexpected, configureNotifications: unexpected,
    loginWhatsApp: unexpected, verifyWhatsApp: unexpected, recreateWhatsApp: unexpected,
    setup: unexpected, diagnostics: unexpected, typecheck: unexpected,
    regressionTests: unexpected, createTemplate: unexpected,
  };
  await runControlPanel(ui, actions);
  assert.equal(answers.length, 0, "All menu answers must be consumed");
  assert.equal(expectedErrors.length, 0, "All expected menu errors must be reported");
  return events;
}
