import { loadConfig } from "../src/config";
import { WhatsAppClient } from "../src/whatsapp/client";

const config = loadConfig();
const client = new WhatsAppClient(config);
let authenticated = false;

const stop = (signal: "SIGINT" | "SIGTERM"): void => {
  console.log(`\n[WhatsApp] Received ${signal}; closing Chromium`);
  void client.close().finally(() => {
    process.exitCode = signal === "SIGINT" ? 130 : 143;
  });
};
process.once("SIGINT", () => stop("SIGINT"));
process.once("SIGTERM", () => stop("SIGTERM"));

try {
  await client.open();
  await client.ensureAuthenticated({ allowQrLogin: true });
  authenticated = true;
} catch (error) {
  console.error(error instanceof Error ? error.message : error);
  process.exitCode = 1;
} finally {
  await client.close();
}

if (authenticated) {
  console.log("[WhatsApp] Session saved in .whatsapp-profile/");
}
