import { runCliMain } from "../src/cli-entrypoint";
import {
  runConfirmedFullTest,
  runControlPanel,
} from "../src/operator/control-panel";
import { createDefaultActions } from "../src/operator/default-actions";
import { runSetupWizard } from "../src/operator/setup";
import { createClackUi } from "../src/operator/ui";

const ui = createClackUi();
runCliMain(async () => {
  const actions = createDefaultActions(ui);
  const result = await runSetupWizard(ui, {
    loginWhatsApp: actions.loginWhatsApp,
  });
  if (result.cancelled) return;
  if (result.nextAction === "diagnostics") {
    await actions.diagnostics();
  } else if (result.nextAction === "main-menu") {
    await runControlPanel(ui, actions);
  } else if (result.nextAction === "full-test") {
    await runConfirmedFullTest(ui, actions);
  } else {
    ui.outro("Setup closed");
  }
});
