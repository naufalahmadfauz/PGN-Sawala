import { loadConfig } from "../src/config";
import { runCliMain } from "../src/cli-entrypoint";
import { runBrowserEntrypoint } from "../src/operator/browser-runtime";
import { verifyWhatsApp } from "../src/whatsapp/auth";

const config = loadConfig();
runCliMain(() =>
  runBrowserEntrypoint(() => verifyWhatsApp(config), {
    headless: config.headless,
    whatsappBrowserChannel: config.whatsappBrowserChannel,
    projectRoot: config.projectRoot,
  }),
);
