import { runCliMain } from "../src/cli-entrypoint";
import { runControlPanel } from "../src/operator/control-panel";
import { createDefaultActions } from "../src/operator/default-actions";
import { createClackUi } from "../src/operator/ui";

const ui = createClackUi();
runCliMain(() => runControlPanel(ui, createDefaultActions(ui)));
