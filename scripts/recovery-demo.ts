import { isEntrypoint, runCliMain } from "../src/cli-entrypoint";
import { REPOSITORY_ROOT } from "../src/environment";
import { createRecoveryDemo, resetRecoveryDemos, type RecoveryDemoOptions } from "../src/recovery/recovery-demo";
import { formatRecoveryDiscovery } from "../src/recovery/recovery-service";
import { discoverRecoveryRun } from "../src/recovery/run-state";

export function parseRecoveryDemoArgs(args: readonly string[]): RecoveryDemoOptions & { reset?: true } {
  if (args.length === 1 && args[0] === "--reset") return { reset: true };
  const options: RecoveryDemoOptions = {};
  const seen = new Set<string>();
  for (const arg of args) {
    const key = arg.split("=")[0];
    if (seen.has(key)) throw new Error(`Duplicate recovery demo option: ${key}`);
    seen.add(key);
    if (arg === "--mode=full" || arg === "--mode=retest") options.mode = arg === "--mode=full" ? "full" : "retest";
    else if (arg === "--source-drift") options.sourceDrift = true;
    else if (arg === "--mismatch") options.mismatch = true;
    else throw new Error("Usage: npm run recovery:demo -- [--mode=full|retest] [--source-drift] [--mismatch]; or npm run recovery:demo:reset");
  }
  return options;
}

if (isEntrypoint(import.meta.url)) {
  runCliMain(async () => {
    const options = parseRecoveryDemoArgs(process.argv.slice(2));
    if (options.reset) {
      const removed = await resetRecoveryDemos();
      console.log(removed.length ? `Recovery demo cleaned.\nRemoved: ${removed.join(", ")}` : "No recovery demo artifacts found.");
      console.log("Real run data was not modified.");
      return;
    }
    await createRecoveryDemo(REPOSITORY_ROOT, options);
    console.log(formatRecoveryDiscovery(await discoverRecoveryRun(REPOSITORY_ROOT)));
    console.log("Google Drive: skipped in demo mode\nDiscord notifications: suppressed in demo mode\nWhatsApp and Playwright: never launched");
    console.log("Inspect with npm run pgn or npm run test:pgn:resume:validate.\nClean only demo data with npm run recovery:demo:reset.");
  });
}
