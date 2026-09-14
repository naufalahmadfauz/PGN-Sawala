import assert from "node:assert/strict";
import { mkdir, mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import test from "node:test";
import { chromium } from "playwright";
import { loadConfig } from "../src/config";
import { collectDiagnostics } from "../src/operator/diagnostics";
import { assertWhatsAppBrowserAvailable, CHROME_NOT_INSTALLED, inspectGoogleChrome } from "../src/whatsapp/browser";
import { WhatsAppClient } from "../src/whatsapp/client";

function config(environment: NodeJS.ProcessEnv = {}) {
  return loadConfig({
    repositoryRoot: "/tmp/pgn-whatsapp-browser-fixture",
    environment: {
      PGN_WHATSAPP_CHAT: "Fixture",
      WHATSAPP_HEADLESS: "true",
      GOOGLE_DRIVE_EVIDENCE_ENABLED: "false",
      ...environment,
    },
  });
}

test("WhatsApp browser channel defaults to bundled Chromium and accepts chrome", () => {
  assert.equal(config().whatsappBrowserChannel, undefined);
  assert.equal(config({ WHATSAPP_BROWSER_CHANNEL: "chrome" }).whatsappBrowserChannel, "chrome");
  assert.throws(() => config({ WHATSAPP_BROWSER_CHANNEL: "edge" }), /must be chrome or empty/);
});

test("REST configuration ignores WhatsApp browser channel settings", () => {
  const rest = loadConfig({
    transport: "rest",
    repositoryRoot: "/tmp/pgn-rest-browser-fixture",
    environment: { WHATSAPP_BROWSER_CHANNEL: "not-a-browser-channel" },
  });
  assert.equal(rest.whatsappBrowserChannel, undefined);
});

test("Chrome availability is checked without launching a browser", async () => {
  const installed = await inspectGoogleChrome({
    platform: "linux",
    pathExists: async (file) => file === "/opt/google/chrome/chrome",
    readVersion: async () => "Google Chrome 153.0.8010.36",
  });
  assert.deepEqual(installed, { installed: true, version: "153.0.8010.36" });
  await assertWhatsAppBrowserAvailable({ whatsappBrowserChannel: "chrome" }, {
    platform: "linux",
    pathExists: async () => true,
  });
  await assert.rejects(
    assertWhatsAppBrowserAvailable({ whatsappBrowserChannel: "chrome" }, {
      platform: "linux",
      pathExists: async () => false,
    }),
    new RegExp(CHROME_NOT_INSTALLED.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")),
  );
});

test("doctor reports configured Chrome availability and does not inspect bundled Chromium", async () => {
  const base = {
    LIVEPERSON_REST_ENABLED: "false",
    PGN_WHATSAPP_CHAT: "Fixture",
    WHATSAPP_HEADLESS: "true",
    WHATSAPP_BROWSER_CHANNEL: "chrome",
  };
  const make = (installed: boolean) => collectDiagnostics({
    projectRoot: "/tmp/pgn-whatsapp-doctor-fixture",
    environment: base,
    npmVersion: async () => "11.0.0",
    packageVersion: async () => "1.0.0",
    pathExists: async () => true,
    chromiumExecutablePath: async () => assert.fail("Chrome channel must not inspect bundled Chromium"),
    inspectChrome: async () => installed ? { installed: true, version: "153.0.8010.36" } : { installed: false },
    hasCommand: async () => true,
    checkDriveAccess: false,
    inspectWorkbookSchema: async () => ({ ready: true, detail: "fixture" }),
    inspectRecovery: async () => ({ kind: "none", lock: { status: "unlocked" } }),
  });
  const installed = await make(true);
  assert.equal(installed.checks.find((check) => check.id === "browser-channel")?.detail, "chrome");
  assert.equal(installed.checks.find((check) => check.id === "chrome")?.status, "ok");
  assert.equal(installed.checks.find((check) => check.id === "chrome-version")?.detail, "153.0.8010.36");
  const missing = await make(false);
  assert.equal(missing.ready, false);
  assert.match(missing.checks.find((check) => check.id === "chrome")?.detail ?? "", /Google Chrome is configured/);
});

test("WhatsApp client passes channel only when configured", async (context) => {
  const launches: unknown[] = [];
  const fakeContext = {
    pages: () => [],
    newPage: async () => ({
      setDefaultTimeout: () => undefined,
      goto: async () => undefined,
    }),
    browser: () => ({ version: () => "153.0.8010.36" }),
    close: async () => undefined,
  };
  context.mock.method(chromium, "launchPersistentContext", async (_dir: string, options: unknown) => {
    launches.push(options);
    return fakeContext as never;
  });
  const roots: string[] = [];
  for (const channel of [undefined, "chrome"] as const) {
    const root = await mkdtemp(path.join(tmpdir(), "pgn-browser-channel-"));
    roots.push(root);
    const client = new WhatsAppClient(config({
      WHATSAPP_BROWSER_CHANNEL: channel ?? "",
      PGN_WHATSAPP_CHAT: "Fixture",
    }));
    (client as unknown as { config: { profileDir: string; artifactsDir: string; debugDir: string; evidenceDir: string } }).config.profileDir = path.join(root, "profile");
    await client.open();
    await client.close();
  }
  context.after(async () => { for (const root of roots) await rm(root, { recursive: true, force: true }); });
  assert.equal((launches[0] as { channel?: string }).channel, undefined);
  assert.equal((launches[1] as { channel?: string }).channel, "chrome");
});
