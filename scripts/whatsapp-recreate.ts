import { loadConfig } from "../src/config";
import { runCliMain } from "../src/cli-entrypoint";
import { runBrowserEntrypoint } from "../src/operator/browser-runtime";
import { recreateWhatsAppAuthentication } from "../src/whatsapp/auth";

const config = loadConfig();
runCliMain(() =>
  runBrowserEntrypoint(() => recreateWhatsAppAuthentication(config), {
    headless: config.headless,
    projectRoot: config.projectRoot,
  }),
);
