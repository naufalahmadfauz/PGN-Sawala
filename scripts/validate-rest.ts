import { constants, access } from "node:fs/promises";
import path from "node:path";
import { loadConfig, type AppConfig } from "../src/config";
import { isEntrypoint, runCliMain } from "../src/cli-entrypoint";
import { loadPgnWorkbook } from "../src/excel/pgn-workbook-loader";
import { assertPgnWorkbookValid } from "../src/excel/pgn-workbook-validator";
import { assertResumeOptionsCompatible, parseCliOptions } from "../src/pgn-cli";
import { selectScenarios } from "../src/pgn-selection";
import { discoverRecoveryRun } from "../src/recovery/run-state";
import { assertRestConfig } from "../src/rest/config";
import { LivePersonClient, type RestDependencies } from "../src/rest/liveperson-client";
import { safeRestError } from "../src/rest/errors";
import { continuousSessionWarning, sessionModeLabel } from "../src/session-mode";

export async function validateRest(
  config: AppConfig = loadConfig({ transport: "rest" }),
  args: string[] = [],
  dependencies: RestDependencies = {},
): Promise<{ ready: boolean; selectedCount: number }> {
  const options = parseCliOptions(["--transport=rest", ...args]);
  assertResumeOptionsCompatible(options);
  if (options.resumeRunId || options.restartRunId) throw new Error("Use test:pgn:resume:validate for stored run recovery readiness");
  try {
    assertRestConfig(config.livePersonRest);
    const source = await loadPgnWorkbook(config.pgnSourceWorkbookPath);
    assertPgnWorkbookValid(source.parsed);
    const outputExists = await access(config.pgnExecutedWorkbookPath).then(() => true).catch(() => false);
    const executed = outputExists ? await loadPgnWorkbook(config.pgnExecutedWorkbookPath) : source;
    assertPgnWorkbookValid(executed.parsed);
    let directory = path.dirname(config.pgnExecutedWorkbookPath);
    while (!(await access(directory).then(() => true).catch(() => false))) {
      const parent = path.dirname(directory);
      if (parent === directory) throw new Error("Executed workbook parent directory is unavailable");
      directory = parent;
    }
    await access(directory, constants.W_OK);
    if (outputExists) await access(config.pgnExecutedWorkbookPath, constants.W_OK);
    const recovery = await discoverRecoveryRun(config.projectRoot);
    if (recovery.kind !== "none" || recovery.lock.status === "active" || recovery.lock.status === "unreadable") throw new Error("Existing PGN recovery/process state needs attention before starting a new REST run");
    const selectedCount = selectScenarios(executed.parsed.scenarios, options, executed.workbook).runnable.length;
    const client = new LivePersonClient(config.livePersonRest, dependencies);
    await client.validate();
    console.log(`REST transport validation\nTransport: REST\nSession Mode: ${sessionModeLabel(options.sessionMode)}\nWorkbook/schema: OK\nOutput access: OK\nSelected: ${selectedCount}\nApp credentials: configured\nDomain discovery: OK\nApplication and consumer authentication: OK\nSkill: configured\nEvidence: Not applicable for REST transport\nNo conversation created and no testcase sent.\nREADY`);
    if (options.sessionMode === "continuous") console.warn(continuousSessionWarning("rest"));
    return { ready: true, selectedCount };
  } catch (error) {
    throw new Error(safeRestError(error, config.livePersonRest?.clientSecret));
  }
}

if (isEntrypoint(import.meta.url)) runCliMain(() => validateRest(undefined, process.argv.slice(2)).then(() => undefined), safeRestError);
