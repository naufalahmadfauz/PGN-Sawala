import { loadConfig } from "../src/config";
import { WhatsAppClient } from "../src/whatsapp/client";

const config = loadConfig();
const client = new WhatsAppClient(config);

try {
  await client.open();
  await client.ensureAuthenticated({ allowQrLogin: false });
  console.log("[WhatsApp] Persistent session verification PASSED");
} catch (error) {
  console.error(error instanceof Error ? error.message : error);
  process.exitCode = 1;
} finally {
  await client.close();
}
