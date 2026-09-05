import type { PgnSheetKind } from "./excel/pgn-types";

export interface CliOptions {
  limit?: number;
  sheet?: PgnSheetKind;
  testIds: Set<string>;
  rerunAll: boolean;
  rerunIds: Set<string>;
  resumeRunId?: string;
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
  };

  for (let index = 0; index < args.length; index += 1) {
    const argument = args[index];
    if (argument === "--limit") {
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
    } else if (argument === "--resume") {
      const value = args[++index]?.trim();
      if (!value || value.startsWith("--")) {
        throw new Error("--resume requires a Run ID");
      }
      options.resumeRunId = value;
    } else if (argument === "--accept-source-drift") {
      options.acceptSourceDrift = true;
    } else {
      throw new Error(`Unknown argument: ${argument}`);
    }
  }
  if (options.acceptSourceDrift && !options.resumeRunId) {
    throw new Error("--accept-source-drift requires --resume");
  }
  return options;
}

export function assertResumeOptionsCompatible(options: CliOptions): void {
  if (
    options.resumeRunId &&
    (options.limit !== undefined ||
      options.sheet !== undefined ||
      options.testIds.size > 0 ||
      options.rerunAll ||
      options.rerunIds.size > 0)
  ) {
    throw new Error(
      "--resume cannot be combined with --limit, --sheet, --test, or --rerun; recovery uses the original selection snapshot",
    );
  }
}
