import { lstat } from "node:fs/promises";
import path from "node:path";
import type { AppConfig } from "../config";
import { recoveryPaths, type RecoveryRunState } from "./run-state";

export function recoveryDemoPaths(projectRoot: string, runId: string) {
  const paths = recoveryPaths(projectRoot, runId);
  if (!runId.startsWith("DEMO-RECOVERY-")) {
    throw new Error("Only explicitly identified recovery demos may use demo storage");
  }
  const directory = path.join(paths.root, "demos", runId);
  return {
    directory,
    sourceWorkbookPath: path.join(directory, "source.xlsx"),
    executedWorkbookPath: path.join(directory, "executed.xlsx"),
  };
}

export async function assertSafeDemoStorage(
  projectRoot: string,
  runId?: string,
): Promise<void> {
  const paths = recoveryPaths(projectRoot, runId);
  const targets = [
    path.join(projectRoot, ".runtime"),
    paths.root,
    paths.runs,
    paths.active,
    paths.lock,
    path.join(paths.root, "demos"),
  ];
  if (runId) {
    const demo = recoveryDemoPaths(projectRoot, runId);
    targets.push(
      paths.runDirectory!, paths.state!, paths.manifest!,
      demo.directory, demo.sourceWorkbookPath, demo.executedWorkbookPath,
      `${demo.executedWorkbookPath}.lock`,
    );
  }
  // Never follow links into real run data, including links in parent directories.
  for (const target of targets) {
    const info = await lstat(target).catch((error) => {
      if ((error as NodeJS.ErrnoException).code === "ENOENT") return undefined;
      throw error;
    });
    if (info?.isSymbolicLink() || (info?.isFile() && info.nlink !== 1)) {
      throw new Error("Recovery demo storage must not contain symbolic or hard links");
    }
  }
}

export async function recoveryDemoConfig(
  config: AppConfig,
  state: RecoveryRunState,
): Promise<AppConfig> {
  if (!state.isDemo) return config;
  const demo = recoveryDemoPaths(config.projectRoot, state.runId);
  if (
    path.resolve(config.projectRoot, state.sourceWorkbookPath) !== demo.sourceWorkbookPath ||
    path.resolve(config.projectRoot, state.executedWorkbookPath) !== demo.executedWorkbookPath
  ) {
    throw new Error("Recovery demo workbook paths must stay in their dedicated demo directory");
  }
  await assertSafeDemoStorage(config.projectRoot, state.runId);
  return {
    ...config,
    pgnSourceWorkbookPath: demo.sourceWorkbookPath,
    pgnExecutedWorkbookPath: demo.executedWorkbookPath,
    googleDriveEvidenceEnabled: false,
    googleServiceAccount: undefined,
    discordNotificationsEnabled: false,
    discordWebhookUrl: undefined,
    target: undefined,
  };
}
