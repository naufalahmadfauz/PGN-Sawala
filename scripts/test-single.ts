import path from "node:path";
import { loadConfig } from "../src/config";
import { runCliMain } from "../src/cli-entrypoint";
import { runBrowserEntrypoint } from "../src/operator/browser-runtime";
import { runTestCases } from "../src/runner";

export async function runSingleTest(): Promise<void> {
  const config = loadConfig();
  config.reportFilePath = path.join(
    config.projectRoot,
    "reports",
    "PGN_Single_Test_Result.xlsx",
  );
  await runTestCases(
    [
      {
        testId: "SINGLE-HALO",
        category: "Greeting",
        userInput: "Halo",
        expectedBehaviour: "Bot should greet the user",
      },
    ],
    config,
  );
}

const config = loadConfig();
runCliMain(() =>
  runBrowserEntrypoint(runSingleTest, {
    headless: config.headless,
    projectRoot: config.projectRoot,
  }),
);
