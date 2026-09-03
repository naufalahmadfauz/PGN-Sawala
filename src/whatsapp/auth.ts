import { loadConfig, type AppConfig } from "../config";
import { rm } from "node:fs/promises";
import { WhatsAppClient } from "./client";

export async function loginWhatsApp(
  config: AppConfig = loadConfig(),
): Promise<void> {
  const client = new WhatsAppClient(config);
  try {
    await client.open();
    await client.ensureAuthenticated({ allowQrLogin: true });
  } finally {
    await client.close();
  }
  console.log("[WhatsApp] Session saved in .whatsapp-profile/");
}

export async function verifyWhatsApp(
  config: AppConfig = loadConfig(),
): Promise<void> {
  const client = new WhatsAppClient(config);
  try {
    await client.open();
    await client.ensureAuthenticated({ allowQrLogin: false });
  } finally {
    await client.close();
  }
  console.log("[WhatsApp] Persistent session verification PASSED");
}

export async function recreateWhatsAppAuthentication(
  config: AppConfig = loadConfig(),
): Promise<void> {
  await rm(config.profileDir, { recursive: true, force: true });
  console.log("[WhatsApp] Saved profile cleared; a new QR login is required");
  await loginWhatsApp(config);
}
