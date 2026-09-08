import { randomUUID } from "node:crypto";
import { loadConfig, type AppConfig } from "../src/config";
import { isEntrypoint, runCliMain } from "../src/cli-entrypoint";
import { assertRestConfig } from "../src/rest/config";
import { safeRestError } from "../src/rest/errors";
import type { RestDependencies } from "../src/rest/liveperson-client";
import { RestTransport } from "../src/transports/rest";

export async function restSmoke(config: AppConfig = loadConfig({ transport: "rest" }), text = "Halo", dependencies: RestDependencies = {}): Promise<void> {
  assertRestConfig(config.livePersonRest);
  if (!text.trim()) throw new Error("REST smoke text must not be empty");
  const turn = { sheetName: "REST Smoke", rowNumber: 2, turnNumber: 1, userInput: text };
  const scenario = { testCaseId: "REST-SMOKE", sheetKind: "kb" as const, sheetName: turn.sheetName, sourceRowNumber: 2, category: "Synthetic REST smoke", rawStatus: "", turns: [turn] };
  const transport = new RestTransport(config.livePersonRest, `SMOKE-${randomUUID()}`, "isolated", console.warn, dependencies);
  console.log("REST smoke test (explicit live conversation)");
  try {
    await transport.initializeRun();
    await transport.beginScenario(scenario, 0);
    console.log("Authentication: OK\nConversation: created");
    const response = await transport.sendMessage(scenario, turn);
    console.log(`Message: ${response.sentAt && response.technicalStatus !== "SEND_ERROR" ? "sent" : "not confirmed"}\nBot response: ${response.technicalStatus === "CAPTURED" ? "received" : response.technicalStatus}\nResponse time: ${response.totalResponseMs ?? "n/a"} ms`);
    await transport.endScenario(scenario);
    await transport.finalizeRun();
    console.log(`Conversation: ${transport.pendingConversationCount ? "cleanup failed; operator follow-up required" : "closed"}`);
    if (transport.pendingConversationCount) throw new Error("REST smoke conversation cleanup failed; the response was preserved");
    if (response.technicalStatus !== "CAPTURED") throw new Error("REST smoke did not capture a complete bot response");
  } catch (error) { throw new Error(safeRestError(error, config.livePersonRest.clientSecret)); }
  finally { await transport.close(); }
}

if (isEntrypoint(import.meta.url)) runCliMain(async () => {
  const args = process.argv.slice(2);
  if (args.length && !(args.length === 2 && args[0] === "--text")) throw new Error("Usage: npm run rest:smoke -- [--text <synthetic message>]");
  await restSmoke(undefined, args[1] ?? "Halo");
}, safeRestError);
