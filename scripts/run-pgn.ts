import { loadConfig } from "../src/config";
import { runCliMain } from "../src/cli-entrypoint";
import { runBrowserEntrypoint } from "../src/operator/browser-runtime";
import { inspectPgnExecution } from "../src/operator/pgn-preflight";
import { runPgnWorkbook } from "../src/pgn-runner";
import { parseCliOptions } from "../src/pgn-cli";

const config = loadConfig({ transport: parseCliOptions(process.argv.slice(2)).transport });
runCliMain(() =>
  runBrowserEntrypoint(
    () => runPgnWorkbook(process.argv.slice(2), "full"),
    {
      headless: config.headless,
      projectRoot: config.projectRoot,
      browserRequired: async () =>
        (await inspectPgnExecution(process.argv.slice(2), "full", config))
          .browserRequired,
    },
  ),
);
