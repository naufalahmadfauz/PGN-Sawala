import path from "node:path";
import { loadConfig } from "../src/config";
import { runTestCases } from "../src/runner";

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
).catch((error) => {
  console.error(error instanceof Error ? error.message : error);
  process.exitCode = 1;
});
