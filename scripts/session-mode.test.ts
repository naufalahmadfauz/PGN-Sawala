import assert from "node:assert/strict";
import { mkdtemp, mkdir, readFile, readdir, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import test, { type TestContext } from "node:test";
import ExcelJS from "exceljs";
import JSZip from "jszip";
import { chromium } from "playwright";
import { loadConfig } from "../src/config";
import { GoogleDriveEvidencePublisher } from "../src/evidence/google-drive";
import { getEvidenceFileMetadata, getEvidenceRunMetadata, readEvidenceHyperlink } from "../src/excel/evidence-workbook";
import { KB_HEADERS, NEGATIVE_HEADERS, loadPgnWorkbook } from "../src/excel/pgn-workbook-loader";
import { getRunConfiguration } from "../src/excel/run-configuration";
import { getRetestRunMetadata, upsertRetestRunMetadata } from "../src/excel/retest-workbook";
import { openExecutedPgnWorkbook, saveExecutedPgnWorkbook } from "../src/excel/pgn-workbook-writer";
import { fieldCell } from "../src/excel/workbook-schema";
import { inspectPgnExecution } from "../src/operator/pgn-preflight";
import { assertResumeOptionsCompatible, parseCliOptions } from "../src/pgn-cli";
import { runPgnWorkbook } from "../src/pgn-runner";
import { abandonRecoveryRun, repairRecoveryProgress, skipRecoveryScenario, validateRecoveryRun } from "../src/recovery/recovery-service";
import { discoverRecoveryRun, hashFile, inspectRunProcessLock, openRecoveryCheckpoint, readRecoveryRun, recoveryPaths } from "../src/recovery/run-state";
import { readSessionMode, shouldResetBeforeScenario, type SessionMode } from "../src/session-mode";
import type { MessageSnapshot, ResponseCapture, SentMessage, WhatsAppMessage } from "../src/types";
import { WhatsAppClient } from "../src/whatsapp/client";
import { parsePgnValidationArgs, validatePgnWorkbook } from "./validate-pgn-workbook";
import { validateRetest } from "./validate-retest";
import { runWorkbookMenu } from "./fixtures/workbook-menu";

const inputs = ["Input one", "Input two A", "Input two B", "Input three", "Negative one", "Negative two"];
const ids = ["SESSION-001", "SESSION-002", "SESSION-003", "SESSION-NEG-001"];

async function fixture(context: TestContext, mode: "full" | "retest" = "full", sheet?: "kb" | "negative" | "both") {
  const root = await mkdtemp(path.join(tmpdir(), "pgn-session-mode-"));
  const config = {
    ...loadConfig({ repositoryRoot: root, environment: {
      PGN_WHATSAPP_CHAT: "Synthetic session-mode fixture",
      GOOGLE_DRIVE_EVIDENCE_ENABLED: mode === "retest" ? "true" : "false",
      GOOGLE_DRIVE_EVIDENCE_PARENT_FOLDER: "fixture-parent-folder",
      GOOGLE_SERVICE_ACCOUNT_JSON: JSON.stringify({ type: "service_account", client_email: "fixture@example.invalid", private_key: "not-a-real-private-key" }),
      DISCORD_NOTIFICATIONS_ENABLED: "false", WHATSAPP_HEADLESS: "true",
    } }),
    postResetQuietMs: 1, resetTimeoutMs: 10,
  };
  await Promise.all([
    mkdir(path.dirname(config.pgnSourceWorkbookPath), { recursive: true }),
    mkdir(config.profileDir, { recursive: true }),
  ]);
  await writeFile(path.join(config.profileDir, "fixture-session"), "synthetic profile sentinel");
  const workbook = new ExcelJS.Workbook();
  const kb = workbook.addWorksheet("Test Case Knowledge Base");
  kb.addRow(KB_HEADERS);
  kb.addRow([1, "Fixture", ids[0], "Fixture", "Objective", "Expected", 1, inputs[0], null, null, null, "Ready for Re-test"]);
  kb.addRow([2, "Fixture", ids[1], "Fixture", "Objective", "Expected", 1, inputs[1], null, null, null, "Ready for Re-test"]);
  kb.addRow([null, null, null, null, null, null, 2, inputs[2]]);
  kb.addRow([3, "Fixture", ids[2], "Fixture", "Objective", "Expected", 1, inputs[3], null, null, null, "Ready for Re-test"]);
  const negative = workbook.addWorksheet("Negative Case");
  negative.addRow(NEGATIVE_HEADERS);
  negative.addRow([1, "Fixture", ids[3], "Objective", `Turn 1: ${inputs[4]}\nTurn 2: ${inputs[5]}`, "Negative", "Expected", null, null, null, "Ready for Re-test"]);
  if (sheet && sheet !== "both") workbook.removeWorksheet(sheet === "kb" ? negative.id : kb.id);
  if (sheet) {
    (workbook.getWorksheet(kb.name) ?? negative).addTable({ name: "CategoryReference", ref: "Q1", columns: [{ name: "Fixture reference" }], rows: [["Preserve reference table"]] });
  }
  await workbook.xlsx.writeFile(config.pgnSourceWorkbookPath);
  let sourceHash = await hashFile(config.pgnSourceWorkbookPath);
  const editSource = async (edit: (workbook: ExcelJS.Workbook) => void) => {
    const source = new ExcelJS.Workbook();
    await source.xlsx.readFile(config.pgnSourceWorkbookPath);
    edit(source);
    await source.xlsx.writeFile(config.pgnSourceWorkbookPath);
    sourceHash = await hashFile(config.pgnSourceWorkbookPath);
  };
  const sent: string[] = [];
  const screenshots: string[] = [];
  const uploads: Array<{ folderId: string; fileName: string }> = [];
  const folders: string[] = [];
  const logs: string[] = [];
  const messages: WhatsAppMessage[] = [];
  const behavior = { responsePrefix: "Synthetic response", failReset: false, failUpload: false, interruptAt: "", timeoutAt: "", signal: "SIGINT" as "SIGINT" | "SIGTERM" };
  const realBrowser = context.mock.method(chromium, "launchPersistentContext", async () => { throw new Error("Live browser forbidden"); });
  const realFetch = context.mock.method(globalThis, "fetch", async () => { throw new Error("Live HTTP forbidden"); });
  context.mock.method(console, "log", (...values: unknown[]) => { logs.push(values.join(" ")); });
  context.mock.method(console, "warn", (...values: unknown[]) => { logs.push(values.join(" ")); });
  context.mock.method(console, "error", (...values: unknown[]) => { logs.push(values.join(" ")); });
  context.mock.method(WhatsAppClient.prototype, "open", async () => undefined);
  context.mock.method(WhatsAppClient.prototype, "ensureAuthenticated", async () => undefined);
  context.mock.method(WhatsAppClient.prototype, "openChat", async () => undefined);
  context.mock.method(WhatsAppClient.prototype, "close", async () => undefined);
  context.mock.method(WhatsAppClient.prototype, "isRemoteTyping", async () => false);
  context.mock.method(WhatsAppClient.prototype, "captureMessageState", async (): Promise<MessageSnapshot> => ({ ids: new Set(messages.map((item) => item.id)), messages: [...messages], messageCount: messages.length }));
  context.mock.method(WhatsAppClient.prototype, "getMessages", async () => [...messages]);
  context.mock.method(WhatsAppClient.prototype, "sendMessage", async (message: string): Promise<SentMessage> => {
    sent.push(message);
    const sentAt = new Date();
    const messageId = `outgoing-${messages.length}`;
    messages.push({ id: messageId, text: message, direction: "outgoing", domIndex: messages.length, observedAt: sentAt });
    if (message === "reset" && !behavior.failReset) messages.push({ id: `reset-confirmation-${messages.length}`, text: "Session deleted", direction: "incoming", domIndex: messages.length, observedAt: new Date() });
    return { sentAt, messageId, renderedText: message };
  });
  context.mock.method(WhatsAppClient.prototype, "waitForBotResponse", async (_baseline: MessageSnapshot, message: SentMessage): Promise<ResponseCapture> => {
    if (message.renderedText === behavior.interruptAt) {
      behavior.interruptAt = "";
      process.emit(behavior.signal);
    }
    const response: WhatsAppMessage = { id: `response-${messages.length}`, text: `${behavior.responsePrefix} to ${message.renderedText}`, direction: "incoming", domIndex: messages.length, observedAt: new Date() };
    messages.push(response);
    return { messages: [response], combinedResponse: response.text, sentAt: message.sentAt, completedAt: new Date(), timedOut: message.renderedText === behavior.timeoutAt, firstResponseMs: 1, totalResponseMs: 2 };
  });
  context.mock.method(WhatsAppClient.prototype, "captureScreenshot", async (filePath: string) => { screenshots.push(filePath); });
  context.mock.method(WhatsAppClient.prototype, "saveDebugArtifacts", async (name: string) => ({ screenshotPath: path.join(root, `${name}.png`), diagnosticsPath: path.join(root, `${name}.json`) }));
  context.mock.method(GoogleDriveEvidencePublisher.prototype, "validateParentFolder", async () => ({ id: "fixture-parent-folder", name: "Fixture parent", webViewLink: "https://example.invalid/parent", reused: true }));
  context.mock.method(GoogleDriveEvidencePublisher.prototype, "validateRunFolder", async (id: string, name?: string) => ({ id, name: name ?? "Fixture folder", webViewLink: "https://example.invalid/folder", reused: true }));
  context.mock.method(GoogleDriveEvidencePublisher.prototype, "ensureRunFolder", async (runId: string) => {
    const id = `folder-${runId}`;
    folders.push(id);
    return { id, name: id, webViewLink: `https://example.invalid/${id}`, reused: false };
  });
  context.mock.method(GoogleDriveEvidencePublisher.prototype, "uploadPng", async ({ folderId, fileName }: { folderId: string; fileName: string }) => {
    uploads.push({ folderId, fileName });
    if (behavior.failUpload) throw new Error("Synthetic upload failure");
    return { id: fileName, name: fileName, webViewLink: `https://example.invalid/${fileName}`, reused: false };
  });
  let settleSignal!: () => void;
  const signalSettled = new Promise<void>((resolve) => { settleSignal = resolve; });
  context.mock.method(process, "kill", (_pid: number, signal?: number | NodeJS.Signals) => {
    if (signal === "SIGINT" || signal === "SIGTERM") settleSignal();
    else assert.equal(signal, 0, "Only synthetic signal settlement or liveness checks are permitted");
    return true;
  });
  context.after(async () => {
    try {
      assert.equal(realBrowser.mock.callCount(), 0);
      assert.equal(realFetch.mock.callCount(), 0);
      assert.equal(await hashFile(config.pgnSourceWorkbookPath), sourceHash);
      assert.equal(await readFile(path.join(config.profileDir, "fixture-session"), "utf8"), "synthetic profile sentinel");
    } finally {
      await rm(root, { recursive: true, force: true });
    }
  });
  const latest = async () => {
    const entries = await readdir(recoveryPaths(root).runs);
    const recovered = await Promise.all(entries.map((id) => readRecoveryRun(root, id)));
    recovered.sort((a, b) => b.state.startedAt.localeCompare(a.state.startedAt));
    return recovered[0];
  };
  return { root, config, sent, screenshots, uploads, folders, logs, messages, behavior, latest, signalSettled, editSource };
}

for (const removed of ["Test Case Knowledge Base", "Negative Case"]) {
  for (const empty of [false, true]) {
    for (const mode of ["full", "retest"] as const) {
      test(`removing ${empty ? "header-only" : "populated, table-bearing"} ${removed} rejects old results and allows ${mode} at a new executed-workbook path`, async (context) => {
        const fx = await fixture(context, "retest", "both");
        await fx.editSource((workbook) => {
          workbook.getWorksheet("Negative Case")!.addTable({ name: "NegativeReference", ref: "Q1", columns: [{ name: "Fixture reference" }], rows: [["Keep old table"]] });
          if (empty) {
            workbook.removeWorksheet(workbook.getWorksheet(removed)!.id);
            workbook.addWorksheet(removed).addRow(removed === "Negative Case" ? NEGATIVE_HEADERS : KB_HEADERS);
          }
        });
        await runPgnWorkbook([], "full", fx.config);
        const oldResults = await readFile(fx.config.pgnExecutedWorkbookPath);
        const oldArchive = await JSZip.loadAsync(oldResults);
        assert.equal(oldArchive.file(/^xl\/tables\/[^/]+\.xml$/).length, empty ? 1 : 2);
        if (empty) {
          const old = await loadPgnWorkbook(fx.config.pgnExecutedWorkbookPath);
          assert.equal(old.workbook.getWorksheet(removed)!.rowCount, 1);
        }
        await fx.editSource((workbook) => { workbook.removeWorksheet(workbook.getWorksheet(removed)!.id); });
        const changedSource = await readFile(fx.config.pgnSourceWorkbookPath);
        const sourceArchive = await JSZip.loadAsync(changedSource);
        assert.equal(sourceArchive.file(/^xl\/tables\/[^/]+\.xml$/).length, 1);
        fx.sent.length = 0;
        await assert.rejects(runPgnWorkbook([], mode, fx.config), (error: unknown) => {
          assert(error instanceof Error);
          assert.match(error.message, /test-case sheet.*differ/i);
          assert(error.message.includes(`Only in executed workbook: "${removed}"`));
          assert.match(error.message, /Use a new executed-workbook path.*PGN_EXECUTED_WORKBOOK/);
          return true;
        });
        assert.deepEqual(fx.sent, []);
        assert.deepEqual(await readFile(fx.config.pgnExecutedWorkbookPath), oldResults);
        assert.deepEqual(await readFile(fx.config.pgnSourceWorkbookPath), changedSource);

        const remaining = removed === "Negative Case" ? "kb" : "negative";
        const nextConfig = { ...fx.config, pgnExecutedWorkbookPath: path.join(fx.root, "reports", `${remaining}-only.xlsx`) };
        await runPgnWorkbook([], mode, nextConfig);
        assert.deepEqual(fx.sent.filter((message) => message !== "reset"), remaining === "kb" ? inputs.slice(0, 4) : inputs.slice(4));
        const reopened = await openExecutedPgnWorkbook(nextConfig.pgnSourceWorkbookPath, nextConfig.pgnExecutedWorkbookPath);
        assert.equal(reopened.resumed, true);
        assert.deepEqual(reopened.parsed.availableSheets, [remaining]);
        assert.equal(reopened.workbook.getWorksheet(removed), undefined);
        for (const scenario of reopened.parsed.scenarios) {
          assert.equal(scenario.status, mode === "retest" ? "Pending Evaluation" : "Ready for Re-test");
          const sheet = reopened.workbook.getWorksheet(scenario.sheetName)!;
          if (remaining === "kb") {
            for (const turn of scenario.turns) assert.equal(fieldCell(sheet, turn.rowNumber, "botResponse").text, `Synthetic response to ${turn.userInput}`);
          } else {
            assert.equal(fieldCell(sheet, scenario.sourceRowNumber, "botResponse").text, "Turn 1:\nSynthetic response to Negative one\n\nTurn 2:\nSynthetic response to Negative two");
          }
        }
        const newArchive = await JSZip.loadAsync(await readFile(nextConfig.pgnExecutedWorkbookPath));
        assert.equal(newArchive.file(/^xl\/tables\/[^/]+\.xml$/).length, 1);
        assert.equal(await newArchive.file("xl/tables/table1.xml")!.async("string"), await sourceArchive.file("xl/tables/table1.xml")!.async("string"));
        assert.deepEqual(await readFile(fx.config.pgnExecutedWorkbookPath), oldResults);
        assert.deepEqual(await readFile(fx.config.pgnSourceWorkbookPath), changedSource);
      });
    }
  }
}

for (const removed of ["Test Case Knowledge Base", "Negative Case"]) {
  test(`full and retest menus explain removed ${removed} even when prior results have no eligible scenarios`, async (context) => {
    const fx = await fixture(context, "retest", "both");
    await runPgnWorkbook([], "retest", fx.config);
    await fx.editSource((workbook) => { workbook.removeWorksheet(workbook.getWorksheet(removed)!.id); });
    const sourceBefore = await readFile(fx.config.pgnSourceWorkbookPath);
    const outputBefore = await readFile(fx.config.pgnExecutedWorkbookPath);
    fx.sent.length = 0;
    const mismatch = new RegExp(`Only in executed workbook: "${removed}".*Use a new executed-workbook path.*PGN_EXECUTED_WORKBOOK`);
    const events = await runWorkbookMenu(fx.config, [
      "run", "remaining", "isolated", true, "accept", "back",
      "retest", "ready", "isolated", "back", "exit",
    ], [mismatch, mismatch]);
    assert.equal(events.filter((message) => mismatch.test(message)).length, 2);
    assert.deepEqual(fx.sent, []);
    assert.deepEqual(await readFile(fx.config.pgnSourceWorkbookPath), sourceBefore);
    assert.deepEqual(await readFile(fx.config.pgnExecutedWorkbookPath), outputBefore);
  });
}

for (const layout of ["kb", "negative", "both"] as const) {
  for (const explicit of [false, true]) {
    test(`${layout} workbook full execution and retest preserve results, evidence, history and tables (${explicit ? "explicit category" : "all categories"})`, async (context) => {
      const fx = await fixture(context, "retest", layout);
      const selectedCategory = explicit ? layout === "both" ? "negative" : layout : undefined;
      const args = selectedCategory ? ["--sheet", selectedCategory] : [];
      const selected = selectedCategory ?? layout;
      const expectedIds = selected === "kb" ? ids.slice(0, 3) : selected === "negative" ? ids.slice(3) : ids;
      const expectedInputs = selected === "kb" ? inputs.slice(0, 4) : selected === "negative" ? inputs.slice(4) : inputs;
      const sourceZip = await JSZip.loadAsync(await readFile(fx.config.pgnSourceWorkbookPath));
      const table = await sourceZip.file("xl/tables/table1.xml")!.async("string");
      assert.equal((await inspectPgnExecution(args, "full", fx.config)).selectedCount, expectedIds.length);
      await runPgnWorkbook(args, "full", fx.config);
      assert.deepEqual(fx.sent.filter((message) => message !== "reset"), expectedInputs);
      const fullRun = (await fx.latest()).state;
      assert.equal(fullRun.status, "COMPLETED");
      assert.deepEqual(fullRun.selectedScenarioIds, expectedIds);
      const { workbook, parsed } = await loadPgnWorkbook(fx.config.pgnExecutedWorkbookPath);
      const expectedSheets = layout === "kb" ? ["Test Case Knowledge Base"] : layout === "negative" ? ["Negative Case"] : ["Test Case Knowledge Base", "Negative Case"];
      assert.deepEqual(workbook.worksheets.filter((sheet) => ["Test Case Knowledge Base", "Negative Case"].includes(sheet.name)).map((sheet) => sheet.name), expectedSheets);
      const previous: Array<{ id: string; response: string; evidence?: string }> = [];
      for (const scenario of parsed.scenarios) {
        const sheet = workbook.getWorksheet(scenario.sheetName)!;
        assert.equal(fieldCell(sheet, scenario.sourceRowNumber, "expectedResponse").text, "Expected");
        assert.equal(scenario.status, "Ready for Re-test");
        if (!expectedIds.includes(scenario.testCaseId)) {
          assert.equal(fieldCell(sheet, scenario.sourceRowNumber, "botResponse").text, "");
          continue;
        }
        if (scenario.sheetKind === "negative") {
          assert.equal(fieldCell(sheet, scenario.sourceRowNumber, "botResponse").text, "Turn 1:\nSynthetic response to Negative one\n\nTurn 2:\nSynthetic response to Negative two");
        }
        for (const turn of scenario.turns) {
          assert.equal(fieldCell(sheet, turn.rowNumber, "userInput").text, scenario.sheetKind === "kb" ? turn.userInput : "Turn 1: Negative one\nTurn 2: Negative two");
          if (scenario.sheetKind === "kb") assert.equal(fieldCell(sheet, turn.rowNumber, "botResponse").text, `Synthetic response to ${turn.userInput}`);
          assert.equal(getEvidenceFileMetadata(workbook, `${fullRun.runId}|${scenario.testCaseId}|${turn.turnNumber}`)?.status, "EVIDENCE_SYNCED");
        }
        for (const row of scenario.sheetKind === "kb" ? scenario.turns.map((turn) => turn.rowNumber) : [scenario.sourceRowNumber]) {
          const evidence = readEvidenceHyperlink(fieldCell(sheet, row, "evidence"));
          assert(evidence);
          previous.push({ id: scenario.testCaseId, response: fieldCell(sheet, row, "botResponse").text, evidence });
        }
      }
      const sentBefore = [...fx.sent];
      await runPgnWorkbook(args, "full", fx.config);
      assert.deepEqual(fx.sent, sentBefore, "Already captured scenarios retain ordinary skip eligibility");
      assert.equal((await validateRetest(args, fx.config)).selectedCount, expectedIds.length);
      fx.behavior.responsePrefix = "Retested response";
      fx.sent.length = 0;
      await runPgnWorkbook(args, "retest", fx.config);
      assert.deepEqual(fx.sent.filter((message) => message !== "reset"), expectedInputs);
      const retestRun = (await fx.latest()).state;
      assert.equal(retestRun.status, "COMPLETED");
      const reopened = await openExecutedPgnWorkbook(fx.config.pgnSourceWorkbookPath, fx.config.pgnExecutedWorkbookPath);
      assert.equal(reopened.resumed, true);
      assert.deepEqual(getRetestRunMetadata(reopened.workbook, retestRun.runId)?.finishedIds, expectedIds);
      const history = reopened.workbook.getWorksheet("Retest History")!;
      assert.equal(history.rowCount - 1, previous.length);
      for (const [index, prior] of previous.entries()) {
        assert.equal(fieldCell(history, index + 2, "testCaseId").text, prior.id);
        assert.equal(fieldCell(history, index + 2, "previousBotResponse").text, prior.response);
        assert.equal(fieldCell(history, index + 2, "previousStatus").text, "Ready for Re-test");
        assert.equal(readEvidenceHyperlink(fieldCell(history, index + 2, "previousEvidenceUrl")), prior.evidence);
        assert.equal(fieldCell(history, index + 2, "newBotResponse").text, prior.response.replaceAll("Synthetic response", "Retested response"));
        assert(readEvidenceHyperlink(fieldCell(history, index + 2, "newEvidenceUrl")));
      }
      for (const scenario of reopened.parsed.scenarios.filter((scenario) => expectedIds.includes(scenario.testCaseId))) {
        assert.equal(scenario.status, "Pending Evaluation");
        for (const turn of scenario.turns) assert.equal(getEvidenceFileMetadata(reopened.workbook, `${retestRun.runId}|${scenario.testCaseId}|${turn.turnNumber}`)?.status, "EVIDENCE_SYNCED");
      }
      const transcript = reopened.workbook.getWorksheet("Execution Transcript")!;
      for (const [runId, prefix] of [[fullRun.runId, "Synthetic response"], [retestRun.runId, "Retested response"]]) {
        const botRows = transcript.getRows(2, transcript.rowCount - 1)!.filter((row) => fieldCell(transcript, row.number, "runId").text === runId && fieldCell(transcript, row.number, "role").text === "BOT");
        assert.deepEqual(botRows.map((row) => fieldCell(transcript, row.number, "message").text), expectedInputs.map((input) => `${prefix} to ${input}`));
      }
      const outputZip = await JSZip.loadAsync(await readFile(fx.config.pgnExecutedWorkbookPath));
      assert.equal(await outputZip.file("xl/tables/table1.xml")!.async("string"), table);
    });
  }
}

for (const layout of ["kb", "negative", "both"] as const) {
  test(`WhatsApp menu executes and retests the ${layout} layout after mapping review`, async (context) => {
    const fx = await fixture(context, "retest", layout);
    const events = await runWorkbookMenu(fx.config, [
      "run", ...(layout === "both" ? ["remaining"] : ["sheet", layout]), "isolated", true, "accept", "back",
      "retest", "ready", "isolated", true, "back", "exit",
    ]);
    assert(events.indexOf("Use this mapping?") < events.indexOf("PGN execution finished"));
    assert(events.includes("Retest execution finished"));
    const expected = layout === "kb" ? inputs.slice(0, 4) : layout === "negative" ? inputs.slice(4) : inputs;
    assert.deepEqual(fx.sent.filter((message) => message !== "reset"), [...expected, ...expected]);
    const { workbook, parsed } = await loadPgnWorkbook(fx.config.pgnExecutedWorkbookPath);
    assert(parsed.scenarios.every((scenario) => scenario.status === "Pending Evaluation"));
    assert(workbook.getWorksheet("Retest History")!.rowCount > 1);
  });
}

for (const [args, expected] of [
  [[], "isolated"], [["--session=isolated"], "isolated"], [["--session=continuous"], "continuous"],
  [["--session", "continuous"], "continuous"], [["--fast"], "continuous"],
  [["--fast", "--session=continuous"], "continuous"],
] as Array<[string[], SessionMode]>) {
  test(`session CLI ${JSON.stringify(args)} resolves ${expected}`, () => {
    assert.equal(parseCliOptions(args).sessionMode, expected);
    assert.equal(parsePgnValidationArgs(args), expected);
  });
}
for (const args of [["--session=dirty"], ["--session="], ["--session"], ["--fast", "--session=isolated"], ["--session=isolated", "--fast"], ["--session=continuous", "--session=isolated"], ["--resume", "RUN", "--restart-run", "RUN"]]) {
  test(`invalid or conflicting flags ${JSON.stringify(args)} are rejected`, () => assert.throws(() => parseCliOptions(args)));
}
test("restart cannot change selection or configure a persistent fast default", () => {
  assert.throws(() => assertResumeOptionsCompatible(parseCliOptions(["--restart-run", "RUN", "--test", "SESSION-001"])), /original selection/);
  assert.equal(parseCliOptions([]).sessionModeExplicit, false);
  assert.equal(readSessionMode(undefined), "isolated");
  assert.deepEqual([0, 1, 2].map((index) => shouldResetBeforeScenario("isolated", index)), [true, true, true]);
  assert.deepEqual([0, 1, 2].map((index) => shouldResetBeforeScenario("continuous", index)), [true, false, false]);
});

for (const mode of ["full", "retest"] as const) {
  for (const option of [undefined, "--session=isolated", "--session=continuous", "--fast"] as const) {
    test(`${mode} execution ${option ?? "default"} preserves exact scenario/turn order and reset policy`, async (context) => {
      const fx = await fixture(context, mode);
      const sessionMode = option === "--fast" || option === "--session=continuous" ? "continuous" : "isolated";
      await runPgnWorkbook(option ? [option] : [], mode, fx.config);
      assert.deepEqual(fx.sent, sessionMode === "continuous"
        ? ["reset", ...inputs]
        : ["reset", inputs[0], "reset", inputs[1], inputs[2], "reset", inputs[3], "reset", inputs[4], inputs[5], "reset"]);
      assert.equal(fx.screenshots.length, inputs.length);
      assert.equal(new Set(fx.screenshots).size, inputs.length);
      const { state, manifest } = await fx.latest();
      assert.equal(state.status, "COMPLETED");
      assert.equal(state.sessionMode, sessionMode);
      assert.equal(manifest.sessionMode, sessionMode);
      assert.equal(state.transport, "whatsapp");
      assert.equal(manifest.transport, "whatsapp");
      assert.deepEqual(state.selectedScenarioIds, ids);
      assert.deepEqual(state.completedScenarioIds, ids);
      assert.equal(state.sessionResetAttempts, sessionMode === "continuous" ? 1 : 5);
      assert.equal(state.finalCleanupComplete, true);
      const { workbook } = await loadPgnWorkbook(fx.config.pgnExecutedWorkbookPath);
      assert.equal(getRunConfiguration(workbook, state.runId)?.sessionMode, sessionMode);
      assert.equal(getRunConfiguration(workbook, state.runId)?.sessionResetAttempts, state.sessionResetAttempts);
      assert.equal(getEvidenceRunMetadata(workbook, state.runId)?.sessionMode, sessionMode);
      assert.equal(getEvidenceFileMetadata(workbook, `${state.runId}|${ids[0]}|1`)?.sessionMode, sessionMode);
      const transcript = workbook.getWorksheet("Execution Transcript")!;
      const boundaries: string[] = [];
      for (let row = 2; row <= transcript.rowCount; row += 1) {
        assert.equal(fieldCell(transcript, row, "sessionMode").text, sessionMode);
        assert.equal(fieldCell(transcript, row, "transport").text, "whatsapp");
        const status = fieldCell(transcript, row, "status").text;
        if (status.startsWith("SCENARIO_ATTEMPT_")) boundaries.push(status);
      }
      assert.deepEqual(boundaries, ids.flatMap(() => ["SCENARIO_ATTEMPT_STARTED", "SCENARIO_ATTEMPT_COMPLETED"]));
      if (mode === "retest") {
        assert.equal(fx.uploads.length, inputs.length);
        assert.deepEqual(getRetestRunMetadata(workbook, state.runId)?.selectedIds, ids);
        assert.deepEqual(getRetestRunMetadata(workbook, state.runId)?.finishedIds, ids);
        assert.equal(getRetestRunMetadata(workbook, state.runId)?.sessionMode, sessionMode);
      }
      assert.match(fx.logs.join("\n"), /Session reset attempts:/);
      assert.match(fx.logs.join("\n"), /Duration:/);
      assert.equal((await inspectRunProcessLock(fx.root)).status, "unlocked");
    });
  }
}

for (const sessionMode of ["isolated", "continuous"] as const) {
  test(`${sessionMode} initial reset failure ignores stale confirmation and sends no testcase`, async (context) => {
    const fx = await fixture(context);
    fx.messages.push({ id: "historical-confirmation", direction: "incoming", text: "Session deleted", domIndex: 0, observedAt: new Date() });
    fx.behavior.failReset = true;
    await assert.rejects(runPgnWorkbook([`--session=${sessionMode}`], "full", fx.config), /Unable to confirm clean/);
    assert.deepEqual(fx.sent, ["reset"]);
    const { state } = await fx.latest();
    assert.equal(state.status, "FAILED");
    assert.equal(state.sessionMode, sessionMode);
    assert.equal(state.sessionResetAttempts, 1);
    assert.deepEqual(state.completedScenarioIds, []);
    assert.equal(fx.screenshots.length, 0);
    const { workbook } = await loadPgnWorkbook(fx.config.pgnExecutedWorkbookPath);
    const transcript = workbook.getWorksheet("Execution Transcript")!;
    assert(transcript.getRows(2, transcript.rowCount - 1)!.some((row) => fieldCell(transcript, row.number, "status").text === "RESET_FAILED"));
  });
}

for (const signal of ["SIGINT", "SIGTERM"] as const) {
  test(`continuous ${signal} interruption checkpoints mode and refuses partial resume, skip, or repair`, async (context) => {
    const fx = await fixture(context);
    fx.behavior.signal = signal;
    fx.behavior.interruptAt = inputs[2];
    await assert.rejects(runPgnWorkbook(["--session=continuous"], "full", fx.config), /interrupted/);
    await fx.signalSettled;
    const { state, manifest } = await fx.latest();
    assert.equal(state.status, "INTERRUPTED");
    assert.equal(state.sessionMode, "continuous");
    assert.equal(manifest.sessionMode, "continuous");
    assert.deepEqual(state.completedScenarioIds, [ids[0]]);
    assert.equal(state.activeScenarioId, ids[1]);
    assert.equal(state.sessionResetAttempts, 1);
    const validation = await validateRecoveryRun(fx.config, state.runId);
    assert.equal(validation.ready, false);
    assert.equal(validation.restartReady, true);
    const before = await hashFile(fx.config.pgnExecutedWorkbookPath);
    const sentBefore = [...fx.sent];
    await assert.rejects(runPgnWorkbook(["--resume", state.runId], "full", fx.config), /Mid-stream resume is disabled/);
    await assert.rejects(inspectPgnExecution(["--resume", state.runId], "full", fx.config), /Mid-stream resume is disabled/);
    await assert.rejects(runPgnWorkbook(["--resume", state.runId, "--session=isolated"], "full", fx.config), /Session mode conflicts/);
    await assert.rejects(skipRecoveryScenario(fx.config, state.runId), /Mid-stream resume/);
    await assert.rejects(repairRecoveryProgress(fx.config, state.runId, "rerun"), /Mid-stream resume/);
    const { checkpoint } = await openRecoveryCheckpoint(fx.root, state.runId);
    await assert.rejects(checkpoint.update((draft) => { draft.sessionMode = "isolated"; }), /cannot be changed/);
    await assert.rejects(checkpoint.update((draft) => { delete draft.sessionMode; }), /cannot be changed/);
    assert.equal(await hashFile(fx.config.pgnExecutedWorkbookPath), before);
    assert.deepEqual(fx.sent, sentBefore);
    await abandonRecoveryRun(fx.config, state.runId);
    assert.equal((await discoverRecoveryRun(fx.root)).kind, "none");
    assert.equal((await readRecoveryRun(fx.root, state.runId)).state.sessionMode, "continuous");
  });
}

for (const mode of ["full", "retest"] as const) {
  test(`${mode} continuous restart repeats the entire original selection under a new Run ID`, async (context) => {
    const fx = await fixture(context, mode);
    fx.behavior.interruptAt = inputs[2];
    await assert.rejects(runPgnWorkbook(["--session=continuous"], mode, fx.config), /interrupted/);
    await fx.signalSettled;
    const previous = await fx.latest();
    const oldManifest = await hashFile(recoveryPaths(fx.root, previous.state.runId).manifest!);
    fx.sent.length = 0;
    const preflight = await inspectPgnExecution(["--restart-run", previous.state.runId], mode, fx.config);
    assert.equal(preflight.selectedCount, ids.length);
    await runPgnWorkbook(["--restart-run", previous.state.runId], mode, fx.config);
    assert.deepEqual(fx.sent, ["reset", ...inputs]);
    const restarted = await fx.latest();
    assert.notEqual(restarted.state.runId, previous.state.runId);
    assert.equal(restarted.state.restartedFromRunId, previous.state.runId);
    assert.equal(restarted.state.sessionMode, "continuous");
    assert.deepEqual(restarted.state.selectedScenarioIds, ids);
    assert.equal(restarted.state.status, "COMPLETED");
    const old = await readRecoveryRun(fx.root, previous.state.runId);
    assert.equal(old.state.status, "ABANDONED");
    assert.deepEqual(old.state.completedScenarioIds, previous.state.completedScenarioIds);
    assert.equal(await hashFile(recoveryPaths(fx.root, previous.state.runId).manifest!), oldManifest);
    const { workbook } = await loadPgnWorkbook(fx.config.pgnExecutedWorkbookPath);
    assert.equal(getRunConfiguration(workbook, restarted.state.runId)?.restartedFromRunId, old.state.runId);
    if (mode === "retest") {
      assert.equal(new Set(fx.folders).size, 2);
      assert.deepEqual(getRetestRunMetadata(workbook, restarted.state.runId)?.selectedIds, ids);
      assert.deepEqual(getRetestRunMetadata(workbook, old.state.runId)?.selectedIds, ids);
    }
  });
}

test("isolated recovery resumes only incomplete scenarios and retains hardened reset behavior", async (context) => {
  const fx = await fixture(context);
  fx.behavior.interruptAt = inputs[2];
  await assert.rejects(runPgnWorkbook([], "full", fx.config), /interrupted/);
  await fx.signalSettled;
  const previous = await fx.latest();
  assert.equal((await validateRecoveryRun(fx.config, previous.state.runId)).ready, true);
  fx.sent.length = 0;
  await runPgnWorkbook(["--resume", previous.state.runId], "full", fx.config);
  assert.deepEqual(fx.sent, ["reset", inputs[1], inputs[2], "reset", inputs[3], "reset", inputs[4], inputs[5], "reset"]);
  assert.equal((await fx.latest()).state.runId, previous.state.runId);
});

test("legacy checkpoints without execution context still resume as isolated", async (context) => {
  const fx = await fixture(context);
  fx.behavior.interruptAt = inputs[2];
  await assert.rejects(runPgnWorkbook([], "full", fx.config), /interrupted/);
  await fx.signalSettled;
  const previous = await fx.latest();
  const paths = recoveryPaths(fx.root, previous.state.runId);
  for (const file of [paths.state!, paths.manifest!]) {
    const value = JSON.parse(await readFile(file, "utf8"));
    delete value.sessionMode;
    delete value.transport;
    delete value.sessionResetAttempts;
    await writeFile(file, JSON.stringify(value));
  }
  assert.equal((await validateRecoveryRun(fx.config, previous.state.runId)).ready, true);
  fx.sent.length = 0;
  await runPgnWorkbook(["--resume", previous.state.runId], "full", fx.config);
  assert.deepEqual(fx.sent, ["reset", inputs[1], inputs[2], "reset", inputs[3], "reset", inputs[4], inputs[5], "reset"]);
  assert.equal((await fx.latest()).state.sessionMode, "isolated");
});

test("legacy workbook-only retests cannot be converted to continuous recovery", async (context) => {
  const fx = await fixture(context, "retest");
  const opened = await openExecutedPgnWorkbook(fx.config.pgnSourceWorkbookPath, fx.config.pgnExecutedWorkbookPath);
  upsertRetestRunMetadata(opened.workbook, {
    runId: "RETEST-LEGACY", startedAt: new Date(), updatedAt: new Date(), state: "IN_PROGRESS",
    selectedIds: ids, finishedIds: [],
  });
  await saveExecutedPgnWorkbook(opened.workbook, fx.config.pgnExecutedWorkbookPath);
  const before = await hashFile(fx.config.pgnExecutedWorkbookPath);
  const args = ["--resume", "RETEST-LEGACY", "--session=continuous"];
  await assert.rejects(runPgnWorkbook(args, "retest", fx.config), /Session mode conflicts/);
  await assert.rejects(inspectPgnExecution(args, "retest", fx.config), /Session mode conflicts/);
  await assert.rejects(validateRetest(args, fx.config), /Session mode conflicts/);
  assert.equal(await hashFile(fx.config.pgnExecutedWorkbookPath), before);
  assert.deepEqual(fx.sent, []);
});

test("continuous evidence upload failures remain fail-open without injecting resets", async (context) => {
  const fx = await fixture(context, "retest");
  fx.behavior.failUpload = true;
  await runPgnWorkbook(["--session=continuous"], "retest", fx.config);
  assert.deepEqual(fx.sent, ["reset", ...inputs]);
  const { state } = await fx.latest();
  assert.equal(state.status, "COMPLETED");
  assert.equal(state.metrics.evidenceUploadErrors, inputs.length);
  assert.equal(state.sessionResetAttempts, 1);
});

test("a failed continuous retest requires a full restart, including previously successful cases", async (context) => {
  const fx = await fixture(context, "retest");
  fx.behavior.timeoutAt = inputs[2];
  await runPgnWorkbook(["--session=continuous"], "retest", fx.config);
  const previous = await fx.latest();
  assert.equal(previous.state.status, "RECOVERABLE");
  assert.deepEqual(previous.state.completedScenarioIds, [ids[0], ids[2], ids[3]]);
  const validation = await validateRecoveryRun(fx.config, previous.state.runId);
  assert.equal(validation.restartReady, true);
  assert.equal(validation.ready, false);
  fx.behavior.timeoutAt = "";
  fx.sent.length = 0;
  await runPgnWorkbook(["--restart-run", previous.state.runId], "retest", fx.config);
  assert.deepEqual(fx.sent, ["reset", ...inputs]);
  assert.deepEqual((await fx.latest()).state.completedScenarioIds, ids);
});

test("continuous validation is read-only, describes shared context, and has no environment default", async (context) => {
  const fx = await fixture(context);
  const before = await hashFile(fx.config.pgnSourceWorkbookPath);
  assert.equal(await validatePgnWorkbook(fx.config, "continuous"), true);
  assert.equal(await hashFile(fx.config.pgnSourceWorkbookPath), before);
  assert.deepEqual(fx.sent, []);
  assert.match(fx.logs.join("\n"), /Continuous/);
  assert.match(fx.logs.join("\n"), /initial|Initial/);
  assert.match(fx.logs.join("\n"), /Disabled|DISABLED/);
  const env = loadConfig({ repositoryRoot: fx.root, environment: { FAST_MODE: "true", SESSION_MODE: "continuous" } });
  assert.equal("sessionMode" in env, false);
  assert.equal(parseCliOptions([]).sessionMode, "isolated");
});
