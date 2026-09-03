import path from "node:path";
import { pathToFileURL } from "node:url";

export function isEntrypoint(importMetaUrl: string): boolean {
  const entrypoint = process.argv[1];
  return Boolean(
    entrypoint &&
      importMetaUrl === pathToFileURL(path.resolve(entrypoint)).href,
  );
}

export function runCliMain(
  main: () => Promise<void>,
  formatError: (error: unknown) => string = (error) =>
    error instanceof Error ? error.message : String(error),
): void {
  void main().catch((error) => {
    console.error(formatError(error));
    process.exitCode = 1;
  });
}
