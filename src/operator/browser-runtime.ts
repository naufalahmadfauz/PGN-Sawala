import path from "node:path";
import { commandAvailable, runInheritedCommand } from "./process";

export type BrowserRuntimeMode = "direct" | "xvfb" | "unavailable";

export interface BrowserRuntimePlan {
  mode: BrowserRuntimeMode;
  reason: string;
}

export interface BrowserRuntimeInput {
  platform: NodeJS.Platform;
  display?: string;
  headless: boolean;
  xvfbAvailable: boolean;
}

export function planBrowserRuntime(
  input: BrowserRuntimeInput,
): BrowserRuntimePlan {
  if (input.headless) {
    return {
      mode: "direct",
      reason: "Chromium is configured to run headless",
    };
  }
  if (input.platform !== "linux") {
    return {
      mode: "direct",
      reason: `headed Chromium runs directly on ${input.platform}`,
    };
  }
  if (input.display?.trim()) {
    return {
      mode: "direct",
      reason: "Linux DISPLAY is available",
    };
  }
  if (input.xvfbAvailable) {
    return {
      mode: "xvfb",
      reason: "headed Chromium needs a virtual display",
    };
  }
  return {
    mode: "unavailable",
    reason:
      "Headed Chromium needs DISPLAY or xvfb-run. Install xvfb, provide DISPLAY, or set WHATSAPP_HEADLESS=true.",
  };
}

export interface DetectBrowserRuntimeOptions {
  platform?: NodeJS.Platform;
  environment?: NodeJS.ProcessEnv;
  headless: boolean;
  hasCommand?: (command: string, args?: readonly string[]) => Promise<boolean>;
}

export async function detectBrowserRuntime(
  options: DetectBrowserRuntimeOptions,
): Promise<BrowserRuntimePlan> {
  const platform = options.platform ?? process.platform;
  const environment = options.environment ?? process.env;
  const needsXvfb =
    platform === "linux" &&
    !options.headless &&
    !environment.DISPLAY?.trim();
  if (needsXvfb && environment.PGN_XVFB_CHILD === "1") {
    return {
      mode: "unavailable",
      reason: "xvfb-run started without providing a usable DISPLAY",
    };
  }
  const xvfbAvailable = needsXvfb
    ? await (options.hasCommand ?? commandAvailable)("xvfb-run", ["--help"])
    : false;
  return planBrowserRuntime({
    platform,
    display: environment.DISPLAY,
    headless: options.headless,
    xvfbAvailable,
  });
}

interface BrowserActionOptions extends DetectBrowserRuntimeOptions {
  scriptPath: string;
  args?: readonly string[];
  projectRoot: string;
  direct: () => Promise<void>;
  browserRequired?: () => Promise<boolean>;
  runCommand?: typeof runInheritedCommand;
}

export async function runBrowserAction(
  options: BrowserActionOptions,
): Promise<void> {
  if (options.browserRequired && !(await options.browserRequired())) {
    await options.direct();
    return;
  }
  const runtime = await detectBrowserRuntime(options);
  if (runtime.mode === "direct") {
    await options.direct();
    return;
  }
  if (runtime.mode === "unavailable") {
    throw new Error(runtime.reason);
  }

  const environment = {
    ...(options.environment ?? process.env),
    PGN_XVFB_CHILD: "1",
  };
  await (options.runCommand ?? runInheritedCommand)(
    "xvfb-run",
    [
      "-a",
      process.execPath,
      "--import",
      "tsx",
      path.resolve(options.scriptPath),
      ...(options.args ?? []),
    ],
    { cwd: options.projectRoot, environment },
  );
}

export async function runBrowserEntrypoint(
  direct: () => Promise<void>,
  options: Omit<
    BrowserActionOptions,
    "scriptPath" | "args" | "direct" | "projectRoot"
  > & { headless: boolean; projectRoot: string },
): Promise<void> {
  const scriptPath = process.argv[1];
  if (!scriptPath) {
    throw new Error("Could not determine the browser command entrypoint");
  }
  await runBrowserAction({
    ...options,
    scriptPath,
    args: process.argv.slice(2),
    direct,
  });
}
