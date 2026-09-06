import { loadConfig } from "../src/config";
import { isEntrypoint, runCliMain } from "../src/cli-entrypoint";
import { createClackUi } from "../src/operator/ui";
import { formatWorkbookMappings, inspectWorkbookMappings, reviewWorkbookMapping } from "../src/operator/workbook-configuration";

if (isEntrypoint(import.meta.url)) {
  runCliMain(async () => {
    const args = process.argv.slice(2);
    if (args.length > 1 || (args.length && args[0] !== "--review")) throw new Error("Usage: npm run workbook:schema -- [--review]");
    const config = loadConfig();
    if (args[0] === "--review") { await reviewWorkbookMapping(createClackUi(), config); return; }
    const inspection = await inspectWorkbookMappings(config);
    console.log(formatWorkbookMappings(inspection));
    if (!inspection.ready) process.exitCode = 1;
  });
}
