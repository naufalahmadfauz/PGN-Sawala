import type { PgnSheetKind } from "./excel/pgn-types";
import { readSessionMode, type SessionMode } from "./session-mode";

export interface CliOptions {
  limit?: number;
  sheet?: PgnSheetKind;
  testIds: Set<string>;
  rerunAll: boolean;
  rerunIds: Set<string>;
  resumeRunId?: string;
  restartRunId?: string;
  sessionMode: SessionMode;
  sessionModeExplicit: boolean;
  acceptSourceDrift: boolean;
}

function parseIdList(value: string): string[] {
  return value
    .split(",")
    .map((item) => item.trim())
    .filter(Boolean);
}

export function parseCliOptions(args: string[]): CliOptions {
  const options: CliOptions = {
    testIds: new Set(),
    rerunAll: false,
    rerunIds: new Set(),
    acceptSourceDrift: false,
    sessionMode: "isolated",
    sessionModeExplicit: false,
  };

  for (let index = 0; index < args.length; index += 1) {
    const argument = args[index];
    if (argument === "--fast" || argument === "--session" || argument.startsWith("--session=")) {
      const value = argument === "--fast" ? "continuous"
        : argument === "--session" ? args[++index] : argument.slice("--session=".length);
      if (!value) throw new Error("--session requires isolated or continuous");
      const sessionMode = readSessionMode(value);
      if (options.sessionModeExplicit && options.sessionMode !== sessionMode) {
        throw new Error("Conflicting session flags: choose either isolated or continuous");
      }
      options.sessionMode = sessionMode;
      options.sessionModeExplicit = true;
    } else if (argument === "--limit") {
      const value = Number(args[++index]);
      if (!Number.isInteger(value) || value < 1) {
        throw new Error("--limit requires a positive integer");
      }
      options.limit = value;
    } else if (argument === "--sheet") {
      const value = args[++index]?.toLowerCase();
      if (value === "kb" || value === "knowledge") {
        options.sheet = "kb";
      } else if (value === "negative" || value === "neg") {
        options.sheet = "negative";
      } else if (value === "all") {
        options.sheet = undefined;
      } else {
        throw new Error("--sheet must be kb, negative, or all");
      }
    } else if (argument === "--test") {
      const value = args[++index];
      if (!value || value.startsWith("--")) {
        throw new Error("--test requires a Test Case ID");
      }
      parseIdList(value).forEach((id) => options.testIds.add(id));
    } else if (argument === "--rerun") {
      const value = args[index + 1];
      if (value && !value.startsWith("--")) {
        index += 1;
        parseIdList(value).forEach((id) => options.rerunIds.add(id));
      } else {
        options.rerunAll = true;
      }
    } else if (argument === "--resume" || argument === "--restart-run") {
      const value = args[++index]?.trim();
      if (!value || value.startsWith("--")) {
        throw new Error(`${argument} requires a Run ID`);
      }
      if (argument === "--resume") options.resumeRunId = value;
      else options.restartRunId = value;
    } else if (argument === "--accept-source-drift") {
      options.acceptSourceDrift = true;
    } else {
      throw new Error(`Unknown argument: ${argument}`);
    }
  }
  if (options.resumeRunId && options.restartRunId) {
    throw new Error("--resume and --restart-run cannot be combined");
  }
  if (options.acceptSourceDrift && !options.resumeRunId && !options.restartRunId) {
    throw new Error("--accept-source-drift requires --resume or --restart-run");
  }
  return options;
}

export function assertResumeOptionsCompatible(options: CliOptions): void {
  if (
    (options.resumeRunId || options.restartRunId) &&
    (options.limit !== undefined ||
      options.sheet !== undefined ||
      options.testIds.size > 0 ||
      options.rerunAll ||
      options.rerunIds.size > 0)
  ) {
    throw new Error(
      "--resume/--restart-run cannot be combined with --limit, --sheet, --test, or --rerun; recovery uses the original selection snapshot",
    );
  }
}
