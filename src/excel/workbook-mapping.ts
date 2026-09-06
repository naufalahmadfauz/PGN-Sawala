import { readFile } from "node:fs/promises";
import path from "node:path";
import type { Workbook } from "exceljs";
import { atomicWriteJson } from "../recovery/run-state";
import { WORKBOOK_SCHEMAS, setWorkbookSchemaOverrides, type WorkbookMappingOverride } from "./workbook-schema";

export const WORKBOOK_MAPPING_FILE = ".workbook-mappings.json";
export interface SavedWorkbookMapping {
  overrides: WorkbookMappingOverride[];
  sheets: Record<string, { fingerprint: string; fields: Record<string, { header: string; columnLetter: string }> }>;
}
export interface WorkbookMappingStore { version: 1; workbooks: Record<string, SavedWorkbookMapping> }

export async function readWorkbookMappingStore(projectRoot: string): Promise<WorkbookMappingStore> {
  let content: string;
  try { content = await readFile(path.join(projectRoot, WORKBOOK_MAPPING_FILE), "utf8"); }
  catch (error) {
    if ((error as NodeJS.ErrnoException).code === "ENOENT") return { version: 1, workbooks: {} };
    throw error;
  }
  const store = JSON.parse(content) as WorkbookMappingStore;
  if (store.version !== 1 || !store.workbooks || typeof store.workbooks !== "object") throw new Error("Workbook mapping configuration is invalid; review column mapping");
  for (const entry of Object.values(store.workbooks)) {
    if (!Array.isArray(entry.overrides) || !entry.sheets || typeof entry.sheets !== "object") throw new Error("Workbook mapping entry is invalid");
    const used = new Set<string>();
    for (const override of entry.overrides) {
      const schema = WORKBOOK_SCHEMAS.find((item) => item.interactive && item.sheetName === override.sheet);
      const key = `${override.sheet}|${override.field}`;
      if (!schema?.fields.some((item) => item.field === override.field) || typeof override.header !== "string" || !override.header.trim() || used.has(key)) throw new Error("Workbook mapping override is invalid");
      used.add(key);
    }
  }
  return store;
}
export function workbookMappingKey(projectRoot: string, filePath: string): string {
  return path.relative(projectRoot, path.resolve(filePath)).replaceAll(path.sep, "/");
}
export async function saveWorkbookMappingStore(projectRoot: string, store: WorkbookMappingStore): Promise<void> {
  await atomicWriteJson(path.join(projectRoot, WORKBOOK_MAPPING_FILE), store);
}
export async function attachWorkbookMappings(workbook: Workbook, filePath: string, fallbackPath?: string): Promise<void> {
  let directory = path.dirname(path.resolve(filePath));
  while (true) {
    const store = await readWorkbookMappingStore(directory);
    const entry = store.workbooks[workbookMappingKey(directory, filePath)] ??
      (fallbackPath ? store.workbooks[workbookMappingKey(directory, fallbackPath)] : undefined);
    if (entry) { setWorkbookSchemaOverrides(workbook, entry.overrides); return; }
    const parent = path.dirname(directory);
    if (parent === directory) return;
    directory = parent;
  }
}
