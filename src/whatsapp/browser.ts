import { execFile } from "node:child_process";
import { constants, access } from "node:fs/promises";
import path from "node:path";
import { promisify } from "node:util";
import type { AppConfig, WhatsAppBrowserChannel } from "../config";

export const CHROME_INSTALL_COMMAND = "npx playwright install --with-deps chrome";
export const CHROME_NOT_INSTALLED = `Google Chrome is configured for WhatsApp but is not installed. Install it with: ${CHROME_INSTALL_COMMAND}`;

export interface ChromeInspectionDependencies {
  platform?: NodeJS.Platform;
  environment?: NodeJS.ProcessEnv;
  pathExists?: (filePath: string) => Promise<boolean>;
  readVersion?: (filePath: string, platform: NodeJS.Platform) => Promise<string | undefined>;
}

export interface ChromeInstallation {
  installed: boolean;
  version?: string;
}

export function whatsappBrowserLabel(channel?: WhatsAppBrowserChannel): string {
  return channel === "chrome" ? "Google Chrome" : "bundled Chromium";
}

export function safeBrowserVersion(value: string | undefined): string | undefined {
  return value?.match(/\b\d{1,4}(?:\.\d{1,6}){1,3}\b/)?.[0];
}

// Match Playwright's stable Chrome channel locations. A browser found only on
// PATH is not sufficient: channel: "chrome" must be able to find the same binary.
function chromeLocations(platform: NodeJS.Platform, environment: NodeJS.ProcessEnv): string[] {
  if (platform === "linux") return ["/opt/google/chrome/chrome"];
  if (platform === "darwin") return ["/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"];
  if (platform !== "win32") return [];
  const roots = [
    environment.LOCALAPPDATA, environment.PROGRAMFILES, environment["PROGRAMFILES(X86)"],
    ...(environment.HOMEDRIVE ? [`${environment.HOMEDRIVE}\\Program Files`, `${environment.HOMEDRIVE}\\Program Files (x86)`] : []),
  ].filter((value): value is string => Boolean(value) && path.win32.isAbsolute(value!));
  return [...new Set(roots.map((root) => path.win32.join(root, "Google", "Chrome", "Application", "chrome.exe")))];
}

export async function findGoogleChrome(dependencies: ChromeInspectionDependencies = {}): Promise<string | undefined> {
  const platform = dependencies.platform ?? process.platform;
  const available = dependencies.pathExists ?? (async (file: string) => {
    try { await access(file, platform === "win32" ? constants.F_OK : constants.X_OK); return true; }
    catch { return false; }
  });
  for (const file of chromeLocations(platform, dependencies.environment ?? process.env)) {
    if (await available(file)) return file;
  }
  return undefined;
}

async function readChromeVersion(file: string, platform: NodeJS.Platform): Promise<string | undefined> {
  const execute = promisify(execFile);
  // Windows chrome.exe --version can open a browser. Read its file metadata instead.
  const result = platform === "win32"
    ? await execute("powershell.exe", ["-NoProfile", "-NonInteractive", "-Command", "(Get-Item -LiteralPath $env:PGN_CHROME_VERSION_PATH).VersionInfo.ProductVersion"], {
      env: { ...process.env, PGN_CHROME_VERSION_PATH: file }, timeout: 3000, maxBuffer: 4096, windowsHide: true,
    })
    : await execute(file, ["--version"], { timeout: 3000, maxBuffer: 4096 });
  return result.stdout;
}

export async function inspectGoogleChrome(dependencies: ChromeInspectionDependencies = {}): Promise<ChromeInstallation> {
  const executable = await findGoogleChrome(dependencies);
  if (!executable) return { installed: false };
  try {
    const version = await (dependencies.readVersion ?? readChromeVersion)(executable, dependencies.platform ?? process.platform);
    return { installed: true, version: safeBrowserVersion(version) };
  } catch {
    return { installed: true }; // The actual browser version is also logged after launch.
  }
}

export async function assertWhatsAppBrowserAvailable(
  config: Pick<AppConfig, "whatsappBrowserChannel">,
  dependencies: ChromeInspectionDependencies = {},
): Promise<void> {
  if (config.whatsappBrowserChannel === "chrome" && !(await findGoogleChrome(dependencies))) {
    throw new Error(CHROME_NOT_INSTALLED);
  }
}
