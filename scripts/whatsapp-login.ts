import { loadConfig } from "../src/config";
import { runCliMain } from "../src/cli-entrypoint";
import { runBrowserEntrypoint } from "../src/operator/browser-runtime";
import { loginWhatsApp } from "../src/whatsapp/auth";

const config = loadConfig();
runCliMain(() =>
  runBrowserEntrypoint(() => loginWhatsApp(config), {
    headless: config.headless,
    projectRoot: config.projectRoot,
  }),
);
