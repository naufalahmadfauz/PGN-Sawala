import { spawn } from "node:child_process";

export interface ProcessResult {
  exitCode: number;
  stdout: string;
  stderr: string;
}

export interface RunProcessOptions {
  cwd?: string;
  environment?: NodeJS.ProcessEnv;
  inherit?: boolean;
}

export function runProcess(
  command: string,
  args: readonly string[],
  options: RunProcessOptions = {},
): Promise<ProcessResult> {
  return new Promise((resolve, reject) => {
    const child = spawn(command, [...args], {
      cwd: options.cwd,
      env: options.environment ?? process.env,
      shell: false,
      stdio: options.inherit ? "inherit" : ["ignore", "pipe", "pipe"],
    });
    let stdout = "";
    let stderr = "";
    child.stdout?.setEncoding("utf8");
    child.stderr?.setEncoding("utf8");
    child.stdout?.on("data", (chunk: string) => {
      stdout += chunk;
    });
    child.stderr?.on("data", (chunk: string) => {
      stderr += chunk;
    });
    child.once("error", reject);
    child.once("close", (code) => {
      resolve({ exitCode: code ?? 1, stdout, stderr });
    });
  });
}

export async function runInheritedCommand(
  command: string,
  args: readonly string[],
  options: Omit<RunProcessOptions, "inherit"> = {},
): Promise<void> {
  const result = await runProcess(command, args, { ...options, inherit: true });
  if (result.exitCode !== 0) {
    throw new Error(`${command} exited with status ${result.exitCode}`);
  }
}

export async function commandAvailable(
  command: string,
  args: readonly string[] = ["--version"],
): Promise<boolean> {
  try {
    await runProcess(command, args);
    return true;
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code === "ENOENT") {
      return false;
    }
    return false;
  }
}
