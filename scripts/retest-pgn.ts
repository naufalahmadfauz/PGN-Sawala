import { runPgnWorkbook } from "../src/pgn-runner";

runPgnWorkbook(process.argv.slice(2), "retest").catch((error) => {
  console.error(error instanceof Error ? error.message : error);
  process.exitCode = 1;
});
