import path from "node:path";
import { fileURLToPath } from "node:url";
import dotenv from "dotenv";

export const REPOSITORY_ROOT = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  "..",
);

export type EnvironmentValueSource = "process environment" | ".env" | "unset";
export type EnvironmentConfigurationSource =
  | "process environment"
  | ".env"
  | "process environment + .env"
  | "defaults";

export interface LoadedEnvironment {
  values: NodeJS.ProcessEnv;
  envFilePath: string;
  envFileLoaded: boolean;
  sourceFor(name: string): EnvironmentValueSource;
}

export interface LoadEnvironmentOptions {
  repositoryRoot?: string;
  values?: NodeJS.ProcessEnv;
}

let defaultEnvironment: LoadedEnvironment | undefined;

export function loadEnvironment(
  options: LoadEnvironmentOptions = {},
): LoadedEnvironment {
  const repositoryRoot = path.resolve(
    options.repositoryRoot ?? REPOSITORY_ROOT,
  );
  const values = options.values ?? process.env;
  const useDefaultCache =
    values === process.env && repositoryRoot === REPOSITORY_ROOT;
  if (useDefaultCache && defaultEnvironment) {
    return defaultEnvironment;
  }

  const preexistingKeys = new Set(
    Object.keys(values).filter((name) =>
      Object.prototype.hasOwnProperty.call(values, name),
    ),
  );
  const envFilePath = path.join(repositoryRoot, ".env");
  const result = dotenv.config({
    path: envFilePath,
    processEnv: values,
    override: false,
    quiet: true,
  });
  const error = result.error as NodeJS.ErrnoException | undefined;
  if (error && error.code !== "ENOENT") {
    throw new Error(
      `Could not load repository .env file (${error.code ?? "read error"})`,
    );
  }
  const envFileKeys = new Set(Object.keys(result.parsed ?? {}));
  const loaded: LoadedEnvironment = {
    values,
    envFilePath,
    envFileLoaded: !error,
    sourceFor(name: string): EnvironmentValueSource {
      if (preexistingKeys.has(name)) {
        return "process environment";
      }
      return envFileKeys.has(name) ? ".env" : "unset";
    },
  };
  if (useDefaultCache) {
    defaultEnvironment = loaded;
  }
  return loaded;
}

export function summarizeEnvironmentSources(
  environment: LoadedEnvironment,
  names: readonly string[],
): EnvironmentConfigurationSource {
  const sources = new Set(
    names
      .map((name) => environment.sourceFor(name))
      .filter((source) => source !== "unset"),
  );
  if (sources.has("process environment") && sources.has(".env")) {
    return "process environment + .env";
  }
  if (sources.has("process environment")) {
    return "process environment";
  }
  if (sources.has(".env")) {
    return ".env";
  }
  return "defaults";
}
