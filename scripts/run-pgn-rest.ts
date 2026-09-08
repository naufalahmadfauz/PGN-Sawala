import { runCliMain } from "../src/cli-entrypoint";
import { runPgnWorkbook } from "../src/pgn-runner";
import { safeRestError } from "../src/rest/errors";

runCliMain(() => runPgnWorkbook(["--transport=rest", ...process.argv.slice(2)], "full"), safeRestError);
