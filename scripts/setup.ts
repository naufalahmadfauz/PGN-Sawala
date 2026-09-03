import { runCliMain } from "../src/cli-entrypoint";
import { REPOSITORY_ROOT } from "../src/environment";
import { runBrowserAction } from "../src/operator/browser-runtime";
import { runSetupWizard } from "../src/operator/setup";
import { createClackUi } from "../src/operator/ui";
import { loadConfig } from "../src/config";
import { loginWhatsApp } from "../src/whatsapp/auth";
import path from "node:path";

const ui = createClackUi();
runCliMain(async () => {
  await runSetupWizard(ui, {
    loginWhatsApp: async () => {
      const config = loadConfig();
      await runBrowserAction({
        headless: config.headless,
        projectRoot: config.projectRoot,
        scriptPath: path.join(REPOSITORY_ROOT, "scripts", "whatsapp-login.ts"),
        direct: () => loginWhatsApp(config),
      });
    },
  });
});
