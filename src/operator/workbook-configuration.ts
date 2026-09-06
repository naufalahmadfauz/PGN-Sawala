import { access } from "node:fs/promises";
import path from "node:path";
import type { AppConfig } from "../config";
import { loadPgnWorkbook } from "../excel/pgn-workbook-loader";
import {
  KB_SCHEMA, NEGATIVE_SCHEMA, formatWorksheetSchema, resolveWorksheetSchema,
  setWorkbookSchemaOverrides, normalizeWorkbookHeader, type WorkbookField,
} from "../excel/workbook-schema";
import {
  readWorkbookMappingStore, saveWorkbookMappingStore, workbookMappingKey,
  type SavedWorkbookMapping,
} from "../excel/workbook-mapping";
import { discoverRecoveryRun, hashFile } from "../recovery/run-state";
import type { OperatorUi } from "./ui";

export async function inspectWorkbookMappings(config: AppConfig, updateCache = false) {
  const store = await readWorkbookMappingStore(config.projectRoot);
  const paths = [config.pgnSourceWorkbookPath];
  if (await access(config.pgnExecutedWorkbookPath).then(() => true).catch(() => false)) paths.push(config.pgnExecutedWorkbookPath);
  const documents = await Promise.all(paths.map(async (filePath) => {
    const loaded = await loadPgnWorkbook(filePath);
    const key = workbookMappingKey(config.projectRoot, filePath);
    const saved = store.workbooks[key];
    const overrides = structuredClone(saved?.overrides ?? []);
    const schemas = [KB_SCHEMA, NEGATIVE_SCHEMA].map((definition) => {
      const worksheet = loaded.workbook.getWorksheet(definition.sheetName);
      if (!worksheet) throw new Error(`Required sheet "${definition.sheetName}" is missing`);
      return resolveWorksheetSchema(worksheet, definition, overrides);
    });
    return { filePath, key, ...loaded, saved, overrides, schemas, hash: await hashFile(filePath) };
  }));
  const notices: string[] = [];
  let changed = false;
  for (const doc of documents) {
    if (!doc.saved) continue;
    for (const schema of doc.schemas) {
      const previous = doc.saved.sheets[schema.sheetName];
      for (const [field, current] of Object.entries(schema.fields)) {
        const old = previous?.fields[field];
        if (old && old.columnLetter !== current.columnLetter) notices.push(`${path.basename(doc.filePath)} / ${schema.sheetName}: ${field} moved ${old.columnLetter} -> ${current.columnLetter}`);
        else if (old && old.header !== current.header) notices.push(`${schema.sheetName}: ${field} header changed to "${current.header}"`);
      }
      if (previous?.fingerprint !== schema.fingerprint) changed = true;
    }
  }
  const ready = documents.every((doc) => doc.schemas.every((schema) => schema.valid));
  if (updateCache && changed && ready) {
    for (const doc of documents) {
      if (doc.saved) store.workbooks[doc.key] = mappingSnapshot(doc.overrides, doc.schemas);
    }
    await saveWorkbookMappingStore(config.projectRoot, store);
  }
  return { documents, store, ready, notices, approved: documents.every((doc) => Boolean(doc.saved)) };
}

function mappingSnapshot(
  overrides: SavedWorkbookMapping["overrides"],
  schemas: Awaited<ReturnType<typeof inspectWorkbookMappings>>["documents"][number]["schemas"],
): SavedWorkbookMapping {
  return {
    overrides,
    sheets: Object.fromEntries(schemas.map((schema) => [schema.sheetName, {
      fingerprint: schema.fingerprint,
      fields: Object.fromEntries(Object.entries(schema.fields).map(([field, value]) => [field, { header: value.header, columnLetter: value.columnLetter }])),
    }])),
  };
}
export function formatWorkbookMappings(inspection: Awaited<ReturnType<typeof inspectWorkbookMappings>>): string {
  return [
    "Workbook schema",
    ...inspection.documents.flatMap((doc) => [path.basename(doc.filePath), ...doc.schemas.map(formatWorksheetSchema)]),
    ...(inspection.notices.length ? ["NOTICE: Workbook layout changed.", ...inspection.notices, inspection.ready ? "Mapping resolved safely from current headers." : "Mapping needs review."] : []),
    `Mapping overrides: ${inspection.documents.reduce((count, doc) => count + doc.overrides.length, 0)}`,
    `Result: ${inspection.ready ? "READY" : "NEEDS REVIEW"}`,
  ].join("\n");
}

export async function reviewWorkbookMapping(ui: OperatorUi, config: AppConfig): Promise<boolean> {
  const recovery = await discoverRecoveryRun(config.projectRoot);
  if (recovery.kind === "unreadable" || ((recovery.kind === "recoverable" || recovery.kind === "running") && !recovery.state.isDemo)) {
    throw new Error("An active or recoverable run exists. Mapping changes are blocked; preserve its source-drift protections before reviewing a new run.");
  }
  const inspection = await inspectWorkbookMappings(config);
  while (true) {
    for (const doc of inspection.documents) {
      setWorkbookSchemaOverrides(doc.workbook, doc.overrides);
      doc.schemas = [KB_SCHEMA, NEGATIVE_SCHEMA].map((definition) => resolveWorksheetSchema(doc.workbook.getWorksheet(definition.sheetName)!, definition, doc.overrides));
    }
    inspection.ready = inspection.documents.every((doc) => doc.schemas.every((schema) => schema.valid));
    ui.note(formatWorkbookMappings(inspection), "Workbook mapping");
    const choice = await ui.select({
      message: "Use this mapping?",
      options: [
        { value: "accept", label: "Yes, save mapping", disabled: !inspection.ready },
        { value: "change", label: "Review / change" },
        { value: "cancel", label: "Cancel" },
      ],
      initialValue: inspection.ready ? "accept" : "change",
    });
    if (!choice || choice === "cancel") return false;
    if (choice === "accept") {
      if (!inspection.ready) continue;
      const currentRecovery = await discoverRecoveryRun(config.projectRoot);
      if (currentRecovery.kind === "unreadable" || ((currentRecovery.kind === "recoverable" || currentRecovery.kind === "running") && !currentRecovery.state.isDemo)) {
        throw new Error("A real run appeared during mapping review; no mapping changes were saved");
      }
      for (const doc of inspection.documents) {
        if (await hashFile(doc.filePath) !== doc.hash) throw new Error("Workbook changed during review; re-detect columns before saving");
        inspection.store.workbooks[doc.key] = mappingSnapshot(doc.overrides, doc.schemas);
      }
      // A not-yet-created executed copy inherits the reviewed source header identities.
      if (inspection.documents.length === 1) inspection.store.workbooks[workbookMappingKey(config.projectRoot, config.pgnExecutedWorkbookPath)] = structuredClone(inspection.store.workbooks[inspection.documents[0].key]);
      await saveWorkbookMappingStore(config.projectRoot, inspection.store);
      ui.success("Workbook-specific mapping saved. Workbook files were not modified.");
      return true;
    }
    const sheetOptions = inspection.documents.flatMap((doc, index) => doc.schemas.map((schema) => ({ value: `${index}:${schema.definition.id}`, label: `${path.basename(doc.filePath)} / ${schema.sheetName}` })));
    const sheetChoice = await ui.select({ message: "Which worksheet mapping?", options: [...sheetOptions, { value: "cancel", label: "Cancel" }] });
    if (!sheetChoice || sheetChoice === "cancel") continue;
    const [index, id] = sheetChoice.split(":");
    const doc = inspection.documents[Number(index)];
    const schema = doc?.schemas.find((schema) => schema.definition.id === id);
    if (!schema) continue;
    const field = await ui.select({ message: "Which semantic field?", options: [...schema.definition.fields.map((item) => ({ value: item.field, label: item.header })), { value: "cancel", label: "Cancel" }] });
    if (!field || field === "cancel") continue;
    const column = await ui.select({
      message: `Select the header to use for ${field}. No similarity guesses are made.`,
      options: [
        ...schema.headers.map((header) => ({ value: String(header.columnIndex), label: `${header.columnLetter} - ${header.header}` })),
        { value: "automatic", label: "Use automatic detection" }, { value: "cancel", label: "None / cancel" },
      ],
    });
    if (!column || column === "cancel") continue;
    const selected = schema.headers.find((header) => String(header.columnIndex) === column);
    if (selected && schema.headers.filter((header) => normalizeWorkbookHeader(header.header) === normalizeWorkbookHeader(selected.header)).length !== 1) {
      ui.warn("Duplicate identical headers cannot be saved safely by position. Rename one in the workbook.");
      continue;
    }
    doc.overrides = doc.overrides.filter((override) => override.sheet !== schema.sheetName || override.field !== field);
    if (selected) doc.overrides.push({ sheet: schema.sheetName, field: field as WorkbookField, header: selected.header });
  }
}

export async function ensureReviewedWorkbookMapping(ui: OperatorUi, config: AppConfig): Promise<boolean> {
  const inspection = await inspectWorkbookMappings(config, true);
  if (!inspection.approved || !inspection.ready) return reviewWorkbookMapping(ui, config);
  if (inspection.notices.length) ui.note(inspection.notices.join("\n"), "Workbook layout changed; mapping updated safely");
  return true;
}
