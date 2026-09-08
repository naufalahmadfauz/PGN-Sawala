import assert from "node:assert/strict";
import { access, mkdir, mkdtemp, readFile, readdir, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import test, { type TestContext } from "node:test";
import ExcelJS from "exceljs";
import { chromium } from "playwright";
import { loadConfig, type AppConfig } from "../src/config";
import { GoogleDriveEvidencePublisher } from "../src/evidence/google-drive";
import { discoverEvidenceInventory } from "../src/evidence/evidence-migration";
import { KB_HEADERS, NEGATIVE_HEADERS, loadPgnWorkbook } from "../src/excel/pgn-workbook-loader";
import { getRunConfiguration } from "../src/excel/run-configuration";
import { getRetestRunMetadata } from "../src/excel/retest-workbook";
import { fieldCell } from "../src/excel/workbook-schema";
import { inspectPgnExecution } from "../src/operator/pgn-preflight";
import { parseCliOptions } from "../src/pgn-cli";
import { runPgnWorkbook } from "../src/pgn-runner";
import { assertRestConfig, loadRestConfig, type LivePersonRestConfig } from "../src/rest/config";
import { RestApiError, safeRestError } from "../src/rest/errors";
import { LivePersonClient } from "../src/rest/liveperson-client";
import { RestTransport } from "../src/transports/rest";
import { WhatsAppClient } from "../src/whatsapp/client";
import { discoverRecoveryRun, hashFile, openRecoveryCheckpoint, readRecoveryRun, recoveryPaths } from "../src/recovery/run-state";
import { validateRecoveryRun } from "../src/recovery/recovery-service";
import { createDiscordNotifier } from "../src/notifications/discord";
import { validateRest } from "./validate-rest";
import { restSmoke } from "./rest-smoke";

const SECRET = "fixture-liveperson-client-secret-never-log";
const APP_TOKEN = "fixture-AppJWT-private-value";
const CONSUMER_TOKEN = "fixture-ConsumerJWS-private-value";
const environment = {
  LIVEPERSON_REST_ENABLED: "true", LIVEPERSON_ACCOUNT_ID: "123456", LIVEPERSON_CLIENT_ID: "fixture-client", LIVEPERSON_CLIENT_SECRET: SECRET, LIVEPERSON_SKILL_ID: "42",
  LIVEPERSON_SENTINEL_DOMAIN: "sentinel.example.invalid", LIVEPERSON_IDP_DOMAIN: "idp.example.invalid",
  LIVEPERSON_ASYNC_MESSAGING_DOMAIN: "async.example.invalid", LIVEPERSON_MESSAGING_REST_DOMAIN: "messaging.example.invalid",
  REST_RESPONSE_IDLE_MS: "3", REST_RESPONSE_TIMEOUT_MS: "60", REST_POLL_INTERVAL_MS: "1", REST_REQUEST_TIMEOUT_MS: "40",
};
const config = (): LivePersonRestConfig => loadRestConfig({ ...environment });
type Message = { sequence: number | string; id: string; dialogId: string; type: string; content: { text: string }; originator: { role: string }; messageAudience?: string };
type Fault = number | Error | { status?: number; headers?: Record<string, string>; body?: unknown };

class MockLivePerson {
  now = 1_800_000_000_000;
  waits: number[] = [];
  calls: Array<{ route: string; url: URL; headers: Headers; body: any }> = [];
  faults = new Map<string, Fault[]>();
  conversations = new Map<string, { id: string; dialog: string; consumer: string; closed: boolean; messages: Message[] }>();
  pending: Array<{ at: number; conversation: string; text: string }> = [];
  noResponse = new Set<string>();
  failText = new Set<string>();
  responseText: string | undefined;
  sequenceStrings = false;
  omitDialogId = false;
  reorderedReplies = false;
  appCalls = 0;
  consumerCalls = 0;
  pageSize = 100;
  afterPublish?: () => void;
  sleep = async (ms: number) => { this.waits.push(ms); this.now += ms; };
  message(conversation: string, role: string, text: string, type = "PLAIN_TEXT"): Message {
    const item = this.conversations.get(conversation)!;
    const seq = item.messages.length;
    const message: Message = { id: `message-${seq}`, sequence: this.sequenceStrings ? String(seq) : seq, dialogId: item.dialog, type, content: { text }, originator: { role }, messageAudience: "ALL" };
    item.messages.push(message);
    return message;
  }
  fetch: typeof fetch = async (input, init) => {
    assert.equal(init?.redirect, "error");
    const url = new URL(String(input));
    const headers = new Headers(init?.headers);
    const body = typeof init?.body === "string" && headers.get("content-type")?.includes("json") ? JSON.parse(init.body) : init?.body;
    const route = url.pathname.includes("baseURI") ? "domains"
      : url.pathname.endsWith("/token") ? "app"
      : url.pathname.endsWith("/consumer") ? "consumer"
      : url.pathname.endsWith("/conversation") ? "create"
      : url.pathname.endsWith("/send") ? body.type === "ms.PublishEvent" ? "send" : "close"
      : url.pathname.endsWith("/messages") ? "messages" : "lookup";
    this.calls.push({ route, url, headers, body });
    const fault = this.faults.get(route)?.shift();
    if (fault instanceof Error) throw fault;
    if (fault !== undefined) {
      const value = typeof fault === "number" ? { status: fault } : fault;
      return new Response(JSON.stringify(value.body ?? { error: `${SECRET} ${APP_TOKEN} ${CONSUMER_TOKEN}` }), { status: value.status ?? 200, headers: value.headers });
    }
    const response = (value: unknown, status = 200) => new Response(JSON.stringify(value), { status });
    if (route === "domains") return response({ baseURIs: ["sentinel", "idp", "asyncMessagingEnt", "messagingRestApiDomain"].map((service) => ({ service, baseURI: `${service.toLowerCase()}.liveperson.net` })) });
    if (route === "app") {
      this.appCalls += 1;
      const params = new URLSearchParams(String(body));
      assert.equal(params.get("client_id"), environment.LIVEPERSON_CLIENT_ID);
      assert.equal(params.get("client_secret"), SECRET);
      return response({ access_token: `${APP_TOKEN}-${this.appCalls}`, expires_in: 3600 });
    }
    assert.match(headers.get("authorization") ?? "", new RegExp(`^${APP_TOKEN}`));
    if (route === "consumer") {
      this.consumerCalls += 1;
      assert.match(body.ext_consumer_id, /^pgn-sawala-/);
      return response({ token: `${CONSUMER_TOKEN}-${body.ext_consumer_id}`, expires_in: 3600 }, 201);
    }
    assert.match(headers.get(route === "messages" || route === "lookup" ? "LP-ON-BEHALF" : "X-LP-ON-BEHALF") ?? "", new RegExp(`^${CONSUMER_TOKEN}`));
    assert.equal(headers.get("Brand-ID"), "123456");
    assert.equal(headers.get("Client-source"), "pgn-sawala-rest");
    if (route === "create") {
      const create = body.find((item: any) => item.type === "cm.ConsumerRequestConversation");
      assert.equal(create.body.skillId, "42");
      assert.equal(create.body.channelType, "MESSAGING");
      assert.equal(create.body.ttrDefName, "NORMAL");
      const id = `conversation-${this.conversations.size + 1}`;
      const dialog = `dialog-${this.conversations.size + 1}`;
      this.conversations.set(id, { id, dialog, consumer: headers.get("X-LP-ON-BEHALF")!, closed: false, messages: [] });
      this.message(id, "BRAND_BOT", "Old greeting that must not become a testcase answer");
      const replies = body.map((item: any) => ({ code: "OK", reqId: item.id, body: item.type === "cm.ConsumerRequestConversation" ? { conversationId: id, ...(!this.omitDialogId ? { dialogId: dialog } : {}) } : {} }));
      return response(this.reorderedReplies ? replies.reverse() : replies, 201);
    }
    if (route === "lookup") {
      const item = this.conversations.get(url.pathname.split("/").at(-1)!)!;
      return response({ id: item.id, dialogs: [{ id: "closed-dialog", dialogType: "MAIN", state: "CLOSE" }, { id: item.dialog, dialogType: "MAIN", state: "OPEN" }] });
    }
    if (route === "send") {
      const item = [...this.conversations.values()].find((value) => value.dialog === body.body.dialogId)!;
      assert(item);
      const text = body.body.event.message;
      assert.equal(body.body.event.type, "ContentEvent");
      assert.equal(body.body.event.contentType, "text/plain");
      if (this.failText.has(text)) return response({ error: "bad input" }, 400);
      const sent = this.message(item.id, "CONSUMER", text);
      this.message(item.id, "AGENT", "typing", "CHAT_STATE");
      if (!this.noResponse.has(text)) {
        this.message(item.id, "ASSIGNED_AGENT", this.responseText ?? `Answer: ${text}`);
        this.message(item.id, "BRAND_BOT", "Second bot bubble");
      }
      this.afterPublish?.();
      return response({ code: "OK", reqId: body.id, body: { sequence: sent.sequence } });
    }
    if (route === "close") {
      const item = this.conversations.get(body.body.conversationId)!;
      assert.equal(body.body.conversationField.conversationState, "CLOSE");
      item.closed = true;
      return response({ code: "OK", reqId: body.id, body: {} });
    }
    const parts = url.pathname.split("/");
    const item = this.conversations.get(parts[parts.indexOf("conversations") + 1])!;
    for (const event of this.pending.filter((event) => event.at <= this.now && event.conversation === item.id)) {
      this.message(item.id, "AGENT", event.text);
    }
    this.pending = this.pending.filter((event) => event.at > this.now || event.conversation !== item.id);
    const start = Number(url.searchParams.get("newerThanSequence"));
    assert.equal(url.searchParams.get("sortOrder"), "ASC");
    const remaining = item.messages.filter((message) => Number(message.sequence) >= start);
    const data = remaining.slice(0, this.pageSize);
    return response({ data, links: remaining.length > data.length ? { next: "https://untrusted.invalid/do-not-follow" } : {} });
  };
}

async function fixture(context: TestContext, api = new MockLivePerson()) {
  const root = await mkdtemp(path.join(tmpdir(), "pgn-rest-"));
  const cfg = loadConfig({ repositoryRoot: root, environment: { ...environment, DISCORD_NOTIFICATIONS_ENABLED: "false", GOOGLE_DRIVE_EVIDENCE_ENABLED: "true", GOOGLE_SERVICE_ACCOUNT_JSON: "must-not-read" } });
  await mkdir(path.dirname(cfg.pgnSourceWorkbookPath), { recursive: true });
  const source = new ExcelJS.Workbook();
  const kb = source.addWorksheet("Test Case Knowledge Base");
  kb.addRow(KB_HEADERS);
  kb.addRow([1, "Fixture", "REST-001", "Tester", "Objective", "Expected", 1, "First", null, null, null, "Ready for Re-test"]);
  kb.addRow([2, "Fixture", "REST-002", "Tester", "Objective", "Expected", 1, "Second A", null, null, null, "Ready for Re-test"]);
  kb.addRow([null, null, null, null, null, null, 2, "Second B"]);
  kb.spliceColumns(8, 0, ["Reviewer", "Do not overwrite"]);
  const negative = source.addWorksheet("Negative Case");
  negative.addRow(NEGATIVE_HEADERS);
  negative.addRow([1, "Fixture", "REST-NEG-001", "Objective", "Negative", "Condition", "Expected", null, null, null, "Ready for Re-test"]);
  await source.xlsx.writeFile(cfg.pgnSourceWorkbookPath);
  const original = await hashFile(cfg.pgnSourceWorkbookPath);
  const noLive = async (): Promise<never> => { throw new Error("Live boundary forbidden"); };
  const forbidden = [
    context.mock.method(chromium, "launchPersistentContext", noLive),
    context.mock.method(WhatsAppClient.prototype, "open", noLive),
    context.mock.method(WhatsAppClient.prototype, "sendMessage", noLive),
    context.mock.method(WhatsAppClient.prototype, "captureScreenshot", noLive),
    ...(["validateParentFolder", "validateRunFolder", "ensureRunFolder", "uploadPng"] as const).map((name) => context.mock.method(GoogleDriveEvidencePublisher.prototype, name, noLive)),
  ];
  context.mock.method(globalThis, "fetch", api.fetch);
  const logs: string[] = [];
  for (const method of ["log", "warn", "error"] as const) context.mock.method(console, method, (...args: unknown[]) => { logs.push(args.join(" ")); });
  context.after(async () => {
    try {
      for (const guard of forbidden) assert.equal(guard.mock.callCount(), 0);
      assert.equal(await hashFile(cfg.pgnSourceWorkbookPath), original);
      for (const secret of [SECRET, APP_TOKEN, CONSUMER_TOKEN]) assert.equal(logs.join("\n").includes(secret), false);
    } finally { await rm(root, { recursive: true, force: true }); }
  });
  const latest = async () => {
    const entries = await readdir(recoveryPaths(root).runs);
    const values = await Promise.all(entries.map((id) => readRecoveryRun(root, id)));
    values.sort((a, b) => b.state.startedAt.localeCompare(a.state.startedAt));
    return values[0];
  };
  return { root, cfg, api, logs, latest };
}

test("REST config defaults, explicit credentials, timing and safe overrides", () => {
  const defaults = loadRestConfig({});
  assert.equal(defaults.enabled, false);
  assert.deepEqual([defaults.responseIdleMs, defaults.responseTimeoutMs, defaults.pollIntervalMs], [3000, 60000, 750]);
  assertRestConfig(config());
  assert.equal(config().domains.sentinel, "https://sentinel.example.invalid");
  assert.equal(parseCliOptions([]).transport, "whatsapp");
  assert.equal(parseCliOptions(["--transport=rest"]).transport, "rest");
  assert.deepEqual([...parseCliOptions(["--test=REST-001"]).testIds], ["REST-001"]);
  assert.throws(() => parseCliOptions(["--transport=rest", "--transport=whatsapp"]), /Conflicting/);
});
for (const entry of ["LIVEPERSON_ACCOUNT_ID", "LIVEPERSON_CLIENT_ID", "LIVEPERSON_CLIENT_SECRET", "LIVEPERSON_SKILL_ID"] as const) {
  test(`missing ${entry} blocks REST safely`, () => {
    assert.throws(() => assertRestConfig(loadRestConfig({ ...environment, [entry]: "" })), new RegExp(entry));
  });
}
for (const value of ["http://example.invalid", "https://secret@example.invalid", "https://example.invalid/path", "https://example.invalid?token=hidden"]) {
  test("unsafe domain origins are rejected without echoing their value", () => {
    assert.throws(() => loadRestConfig({ ...environment, LIVEPERSON_IDP_DOMAIN: value }), (error: unknown) => {
      assert(error instanceof Error); assert.equal(error.message.includes(value), false); return true;
    });
  });
}
for (const overrides of [{ REST_RESPONSE_IDLE_MS: "60000" }, { REST_POLL_INTERVAL_MS: "0" }, { REST_RESPONSE_TIMEOUT_MS: "bad" }]) {
  test("invalid REST timing is rejected", () => assert.throws(() => loadRestConfig({ ...environment, ...overrides })));
}

test("AppJWT and consumer tokens are cached, refreshed near expiry, and use runtime source credentials", async () => {
  const api = new MockLivePerson();
  const client = new LivePersonClient(config(), { fetch: api.fetch, now: () => api.now, sleep: api.sleep });
  const consumer = client.syntheticConsumer("RUN", "scenario");
  await client.applicationToken(); await client.applicationToken();
  await client.consumerToken(consumer); await client.consumerToken(consumer);
  assert.equal(api.appCalls, 1); assert.equal(api.consumerCalls, 1);
  api.now += 3_580_000;
  await client.applicationToken(); await client.consumerToken(consumer);
  assert.equal(api.appCalls, 2); assert.equal(api.consumerCalls, 2);
  assert.equal(client.syntheticConsumer("RUN", "scenario"), consumer);
  assert.notEqual(client.syntheticConsumer("RUN", "other"), consumer);
});

test("JWT expiration claims shorten the cache lifetime", async () => {
  const api = new MockLivePerson();
  const token = `eyJhbGciOiJub25lIn0.${Buffer.from(JSON.stringify({ exp: (api.now + 20000) / 1000 })).toString("base64url")}.fixture-signature`;
  api.faults.set("app", [{ body: { access_token: token, expires_in: 3600 } }]);
  const client = new LivePersonClient(config(), { fetch: api.fetch, now: () => api.now, sleep: api.sleep });
  await client.applicationToken(); await client.applicationToken();
  assert.equal(api.calls.filter((call) => call.route === "app").length, 2);
});

for (const route of ["headers", "body"]) {
  test(`request timeout bounds stalled ${route} and retries safe authentication only three times`, async () => {
    let calls = 0;
    const fetchMock: typeof fetch = async () => {
      calls += 1;
      if (route === "headers") return new Promise<Response>(() => undefined);
      return { ok: true, status: 200, headers: new Headers(), text: () => new Promise<string>(() => undefined) } as Response;
    };
    const client = new LivePersonClient({ ...config(), requestTimeoutMs: 2 }, { fetch: fetchMock, sleep: async () => undefined });
    await assert.rejects(client.applicationToken(), /network\/timeout/);
    assert.equal(calls, 3);
  });
}

test("domain discovery is cached and explicit domain overrides take precedence", async () => {
  const api = new MockLivePerson();
  const cfg = { ...config(), domains: { sentinel: "https://override.example.invalid" } };
  const client = new LivePersonClient(cfg, { fetch: api.fetch });
  await client.validate(); await client.validate();
  assert.equal(api.calls.filter((call) => call.route === "domains").length, 1);
  assert.equal(api.calls.find((call) => call.route === "app")?.url.hostname, "override.example.invalid");
  assert.equal(api.calls.some((call) => call.route === "create"), false);
});

for (const body of [{ baseURIs: [] }, { baseURIs: [{ service: "sentinel", baseURI: "evil.example.invalid" }] }]) {
  test("missing or untrusted discovered domains fail before credentials are sent", async () => {
    const api = new MockLivePerson(); api.faults.set("domains", [{ body }]);
    const client = new LivePersonClient({ ...config(), domains: {} }, { fetch: api.fetch });
    await assert.rejects(client.applicationToken(), /missing|untrusted/);
    assert.equal(api.calls.length, 1);
  });
}

test("conversation batch responses correlate by request ID and resolve the open MAIN dialog", async () => {
  const api = new MockLivePerson(); api.omitDialogId = true; api.reorderedReplies = true;
  const client = new LivePersonClient(config(), { fetch: api.fetch });
  const conversation = await client.createConversation(client.syntheticConsumer("RUN", "case"));
  assert.equal(conversation.conversationId, "conversation-1"); assert.equal(conversation.dialogId, "dialog-1");
  assert.equal(api.calls.filter((call) => call.route === "lookup").length, 1);
  const create = api.calls.find((call) => call.route === "create")!;
  assert.deepEqual(create.body[0].body.authenticatedData.lp_sdes[0].personal, { firstname: "PGN", lastname: "REST-Test" });
  await client.closeConversation(conversation);
  assert(api.conversations.get(conversation.conversationId)?.closed);
});

test("credential-shaped conversation identifiers are never accepted into metadata", async () => {
  const api = new MockLivePerson();
  api.faults.set("create", [{ body: { code: "OK", body: { conversationId: SECRET, dialogId: "dialog" } } }]);
  const client = new LivePersonClient(config(), { fetch: api.fetch });
  await assert.rejects(client.createConversation(client.syntheticConsumer("RUN", "case")), (error: unknown) => {
    assert(error instanceof Error); assert.match(error.message, /credential material/); assert.equal(error.message.includes(SECRET), false); return true;
  });
});

for (const route of ["app", "consumer", "messages", "close"]) {
  for (const status of [429, 500, 503]) {
    test(`${route} HTTP ${status} retries are bounded and safe`, async () => {
      const api = new MockLivePerson();
      const client = new LivePersonClient(config(), { fetch: api.fetch, now: () => api.now, sleep: api.sleep });
      if (route === "app") {
        api.faults.set(route, [status]); await client.applicationToken();
      } else if (route === "consumer") {
        api.faults.set(route, [status]); await client.consumerToken(client.syntheticConsumer("RUN", "case"));
      } else {
        const conversation = await client.createConversation(client.syntheticConsumer("RUN", "case"));
        api.faults.set(route, [status]);
        if (route === "messages") await client.messages(conversation, -1);
        else await client.closeConversation(conversation);
      }
      assert.equal(api.calls.filter((call) => call.route === route).length, 2);
      assert.equal(api.waits[0], 250);
    });
  }
}

test("Retry-After seconds/date are respected and excessive waits are not silently shortened", async () => {
  for (const header of ["2", new Date(1_800_000_002_000).toUTCString(), "120"]) {
    const api = new MockLivePerson(); api.faults.set("app", [{ status: 429, headers: { "Retry-After": header } }]);
    const client = new LivePersonClient(config(), { fetch: api.fetch, now: () => api.now, sleep: api.sleep });
    if (header === "120") { await assert.rejects(client.applicationToken(), /budget/); assert.deepEqual(api.waits, []); }
    else { await client.applicationToken(); assert.equal(api.waits[0], 2000); }
  }
});

for (const status of [401, 403, 429, 503]) {
  test(`persistent HTTP ${status} becomes a global infrastructure failure`, async () => {
    const api = new MockLivePerson(); api.faults.set("app", [status, status, status]);
    const client = new LivePersonClient(config(), { fetch: api.fetch, now: () => api.now, sleep: api.sleep });
    await assert.rejects(client.applicationToken(), (error: unknown) => {
      assert(error instanceof RestApiError); assert.equal(error.global, true);
      assert.equal(error.message.includes(SECRET), false); return true;
    });
    assert.equal(api.calls.length, status === 429 || status === 503 ? 3 : 1);
  });
}

test("expired request authentication refreshes AppJWT and ConsumerJWS once", async () => {
  const api = new MockLivePerson();
  const client = new LivePersonClient(config(), { fetch: api.fetch, now: () => api.now, sleep: api.sleep });
  const conversation = await client.createConversation(client.syntheticConsumer("RUN", "case"));
  api.faults.set("messages", [401]);
  await client.messages(conversation, -1);
  assert.equal(api.appCalls, 2); assert.equal(api.consumerCalls, 2);
  api.faults.set("messages", [401, 401]);
  await assert.rejects(client.messages(conversation, -1), /HTTP 401/);
  assert.equal(api.appCalls, 3);
});

test("Send API 429 rejection and 401 expiry can retry without replaying an accepted publish", async () => {
  for (const status of [429, 401]) {
    const api = new MockLivePerson();
    const client = new LivePersonClient(config(), { fetch: api.fetch, now: () => api.now, sleep: api.sleep });
    const conversation = await client.createConversation(client.syntheticConsumer("RUN", "case"));
    api.faults.set("send", [status]);
    await client.sendText(conversation, "First");
    assert.equal(api.calls.filter((call) => call.route === "send").length, 2);
    assert.equal(api.conversations.get(conversation.conversationId)!.messages.filter((message) => message.originator.role === "CONSUMER").length, 1);
  }
});

test("safe network failures retry while malformed JSON and ambiguous command replies fail closed", async () => {
  const api = new MockLivePerson(); api.faults.set("app", [new Error("synthetic network failure")]);
  const client = new LivePersonClient(config(), { fetch: api.fetch, now: () => api.now, sleep: api.sleep });
  await client.applicationToken(); assert.equal(api.calls.filter((call) => call.route === "app").length, 2);
  const invalid = new LivePersonClient(config(), { fetch: async () => new Response("not json") });
  await assert.rejects(invalid.applicationToken(), /invalid JSON/);
  api.faults.set("create", [{ body: [{ reqId: "unrelated", code: "OK", body: {} }, { reqId: "also-unrelated", code: "OK", body: {} }] }]);
  await assert.rejects(client.createConversation(client.syntheticConsumer("RUN", "case")), /correlated safely/);
});

for (const route of ["create", "send"]) {
  for (const fault of [500, new Error(`network ${APP_TOKEN}`)]) {
    test(`ambiguous ${route} failures are not retried into duplicate messages/conversations`, async () => {
      const api = new MockLivePerson();
      const client = new LivePersonClient(config(), { fetch: api.fetch });
      if (route === "create") {
        api.faults.set(route, [fault]);
        await assert.rejects(client.createConversation(client.syntheticConsumer("RUN", "case")), /uncertain/);
      } else {
        const conversation = await client.createConversation(client.syntheticConsumer("RUN", "case"));
        api.faults.set(route, [fault]); await assert.rejects(client.sendText(conversation, "Input"), /uncertain/);
      }
      assert.equal(api.calls.filter((call) => call.route === route).length, 1);
    });
  }
}

test("pagination and inclusive sequence filtering never reread old bot answers or follow remote next links", async () => {
  const api = new MockLivePerson(); api.pageSize = 2; api.sequenceStrings = true;
  const client = new LivePersonClient(config(), { fetch: api.fetch });
  const conversation = await client.createConversation(client.syntheticConsumer("RUN", "case"));
  for (let index = 0; index < 7; index += 1) api.message(conversation.conversationId, "AGENT", `Message ${index}`);
  const messages = await client.messages(conversation, 2);
  assert.deepEqual(messages.map((message) => message.sequence), [3, 4, 5, 6, 7]);
  assert(api.calls.filter((call) => call.route === "messages").every((call) => call.url.hostname === "messaging.example.invalid"));
  assert.equal((await client.messages(conversation, 7)).length, 0);
});

test("REST collector combines only new agent text and restarts idle settlement for delayed bubbles", async () => {
  const api = new MockLivePerson();
  const transport = new RestTransport(config(), "RUN", "continuous", () => undefined, { fetch: api.fetch, now: () => api.now, sleep: api.sleep });
  await transport.initializeRun();
  const scenario = { testCaseId: "CASE", sheetKind: "kb" as const, sheetName: "Fixture", sourceRowNumber: 2, category: "Fixture", rawStatus: "", turns: [] };
  const turn = { sheetName: "Fixture", rowNumber: 2, turnNumber: 1, userInput: "First" };
  await transport.beginScenario(scenario, 0);
  api.pending.push({ at: api.now + 2, conversation: "conversation-1", text: "Delayed bubble" });
  const first = await transport.sendMessage(scenario, turn);
  assert.equal(first.technicalStatus, "CAPTURED");
  assert.equal(first.combinedResponse, "Answer: First\n\nSecond bot bubble\n\nDelayed bubble");
  assert.equal(first.totalResponseMs, 2);
  assert(api.waits.reduce((sum, ms) => sum + ms, 0) >= 5);
  const next = await transport.sendMessage(scenario, { ...turn, turnNumber: 2, userInput: "Next" });
  assert.equal(next.combinedResponse, "Answer: Next\n\nSecond bot bubble");
  assert.equal(next.evidenceStatus, "EVIDENCE_NOT_APPLICABLE");
  assert.equal(next.transport, "rest");
  assert.equal(next.conversationId, first.conversationId);
  await transport.finalizeRun(); await transport.close();
});

test("no bot response reaches hard timeout with no invented text", async () => {
  const api = new MockLivePerson(); api.noResponse.add("Silent");
  const transport = new RestTransport(config(), "RUN", "continuous", () => undefined, { fetch: api.fetch, now: () => api.now, sleep: api.sleep });
  await transport.initializeRun();
  const result = await transport.sendMessage({} as never, { sheetName: "Fixture", rowNumber: 2, turnNumber: 1, userInput: "Silent" });
  assert.equal(result.technicalStatus, "TIMEOUT"); assert.equal(result.combinedResponse, ""); assert.deepEqual(result.botMessages, []);
  assert.equal(result.totalResponseMs, undefined);
  await transport.close();
});

test("separate attempts under the same Run ID use different synthetic consumers", async () => {
  const api = new MockLivePerson();
  for (let attempt = 0; attempt < 2; attempt += 1) {
    const transport = new RestTransport(config(), "SAME-RUN", "continuous", () => undefined, { fetch: api.fetch });
    await transport.initializeRun(); await transport.close();
  }
  const consumers = api.calls.filter((call) => call.route === "consumer").map((call) => call.body.ext_consumer_id);
  assert.equal(new Set(consumers).size, 2);
});

test("receipts, consumer text, internal audience and metadata events cannot settle a bot response", async () => {
  const api = new MockLivePerson(); api.noResponse.add("Silent");
  const transport = new RestTransport(config(), "RUN", "continuous", () => undefined, { fetch: api.fetch, now: () => api.now, sleep: api.sleep });
  await transport.initializeRun();
  api.afterPublish = () => {
    api.message("conversation-1", "CONSUMER", "Self echo");
    api.message("conversation-1", "AGENT", "Receipt", "ACCEPT_STATUS");
    api.message("conversation-1", "READER", "Metadata");
    api.message("conversation-1", "AGENT", "Agent-internal").messageAudience = "AGENTS";
  };
  const result = await transport.sendMessage({} as never, { sheetName: "Fixture", rowNumber: 2, turnNumber: 1, userInput: "Silent" });
  assert.equal(result.technicalStatus, "TIMEOUT"); assert.equal(result.combinedResponse, "");
  await transport.close();
});

for (const mode of ["full", "retest"] as const) {
  for (const session of ["isolated", "continuous"] as const) {
    test(`shared engine REST ${mode}/${session} writes results without any browser, reset, or Drive dependency`, async (context) => {
      const fx = await fixture(context);
      await runPgnWorkbook(["--transport=rest", `--session=${session}`], mode, fx.cfg);
      assert.equal(fx.api.conversations.size, session === "isolated" ? 3 : 1);
      assert([...fx.api.conversations.values()].every((item) => item.closed));
      const sends = fx.api.calls.filter((call) => call.route === "send");
      assert.deepEqual(sends.map((call) => call.body.body.event.message), ["First", "Second A", "Second B", "Negative"]);
      assert.equal(sends[1].body.body.dialogId, sends[2].body.body.dialogId);
      const { state, manifest } = await fx.latest();
      assert.equal(state.status, "COMPLETED"); assert.equal(state.transport, "rest"); assert.equal(manifest.transport, "rest");
      assert.equal(state.sessionResetAttempts, 0); assert.equal(state.driveRunFolderId, undefined);
      assert.deepEqual(state.restTarget, { accountId: "123456", skillId: "42" });
      const { workbook, parsed } = await loadPgnWorkbook(fx.cfg.pgnExecutedWorkbookPath);
      assert.equal(getRunConfiguration(workbook, state.runId)?.restPollIntervalMs, 1);
      for (const scenario of parsed.scenarios) {
        const sheet = workbook.getWorksheet(scenario.sheetName)!;
        assert.match(fieldCell(sheet, scenario.sourceRowNumber, "botResponse").text, /Answer:/);
        if (mode === "retest") assert.equal(fieldCell(sheet, scenario.sourceRowNumber, "status").text, "Pending Evaluation");
      }
      assert.equal(workbook.getWorksheet("Test Case Knowledge Base")!.getCell(2, 8).text, "Do not overwrite");
      const transcript = workbook.getWorksheet("Execution Transcript")!;
      for (let row = 2; row <= transcript.rowCount; row += 1) {
        assert.equal(fieldCell(transcript, row, "transport").text, "rest");
        if (fieldCell(transcript, row, "role").text === "USER") {
          assert(fieldCell(transcript, row, "conversationId").text);
          assert(fieldCell(transcript, row, "dialogId").text);
          assert.equal(fieldCell(transcript, row, "evidenceStatus").text, "EVIDENCE_NOT_APPLICABLE");
        }
      }
      const inventory = discoverEvidenceInventory(workbook);
      assert.equal(inventory.records.length, 0); assert.equal(inventory.missingCompletedTurns.length, 0);
      if (mode === "retest") assert.deepEqual(getRetestRunMetadata(workbook, state.runId)?.selectedIds, state.selectedScenarioIds);
      await assert.rejects(access(fx.cfg.profileDir)); await assert.rejects(access(fx.cfg.evidenceDir));
      for (const secret of [SECRET, APP_TOKEN, CONSUMER_TOKEN]) {
        assert.equal(JSON.stringify(state).includes(secret), false); assert.equal(JSON.stringify(manifest).includes(secret), false);
        assert.equal(JSON.stringify(transcript.getSheetValues()).includes(secret), false);
      }
    });
  }
}

test("scenario timeout and rejected message do not abort the isolated REST bulk run", async (context) => {
  const fx = await fixture(context);
  fx.api.noResponse.add("First"); fx.api.failText.add("Second A");
  await runPgnWorkbook(["--transport=rest"], "full", fx.cfg);
  const { state } = await fx.latest();
  assert.equal(state.status, "COMPLETED"); assert.equal(state.metrics.timeouts, 1); assert.equal(state.metrics.technicalErrors, 1); assert.equal(state.metrics.capturedScenarios, 1);
  assert.deepEqual(fx.api.calls.filter((call) => call.route === "send").map((call) => call.body.body.event.message), ["First", "Second A", "Negative"]);
});

test("global authentication failure aborts safely before testcase conversations", async (context) => {
  const fx = await fixture(context); fx.api.faults.set("app", [401]);
  await assert.rejects(runPgnWorkbook(["--transport=rest"], "full", fx.cfg), /HTTP 401/);
  assert.equal(fx.api.conversations.size, 0); assert.equal((await fx.latest()).state.status, "FAILED");
});

test("cleanup failures are reported without erasing captured responses", async (context) => {
  const fx = await fixture(context); fx.api.faults.set("close", [400, 400, 400]);
  await runPgnWorkbook(["--transport=rest"], "full", fx.cfg);
  assert.equal((await fx.latest()).state.status, "COMPLETED");
  assert.match(fx.logs.join("\n"), /REST cleanup/);
  assert.equal((await fx.latest()).state.metrics.capturedScenarios, 3);
});

for (const session of ["isolated", "continuous"] as const) {
  test(`REST ${session} recovery uses a fresh conversation and preserves transport policy`, async (context) => {
    const fx = await fixture(context);
    let settle!: () => void; const settled = new Promise<void>((resolve) => { settle = resolve; });
    context.mock.method(process, "kill", (_pid: number, signal?: number | NodeJS.Signals) => { if (signal === "SIGINT") settle(); return true; });
    fx.api.afterPublish = () => {
      if (fx.api.calls.filter((call) => call.route === "send").length === 3) { fx.api.afterPublish = undefined; process.emit("SIGINT"); }
    };
    await assert.rejects(runPgnWorkbook(["--transport=rest", `--session=${session}`], "full", fx.cfg), /interrupted/);
    await settled;
    const previous = await fx.latest();
    assert.equal(previous.state.status, "INTERRUPTED"); assert.equal(previous.state.transport, "rest");
    const opened = await openRecoveryCheckpoint(fx.root, previous.state.runId);
    await assert.rejects(opened.checkpoint.update((state) => { state.restTarget = { accountId: "123456", skillId: "99" }; }), /REST account and skill cannot be changed/);
    const validation = await validateRecoveryRun(fx.cfg, previous.state.runId);
    assert.equal(validation.ready, session === "isolated"); assert.equal(validation.restartReady, session === "continuous");
    const changedTarget = await validateRecoveryRun({ ...fx.cfg, livePersonRest: { ...fx.cfg.livePersonRest!, skillId: "99" } }, previous.state.runId, { checkRestAccess: false });
    assert.equal(changedTarget.ready, false); assert.equal(changedTarget.restartReady, false);
    await assert.rejects(runPgnWorkbook(["--transport=whatsapp", "--resume", previous.state.runId], "full", fx.cfg), /Transport conflicts/);
    if (session === "continuous") await assert.rejects(runPgnWorkbook(["--resume", previous.state.runId], "full", fx.cfg), /Mid-stream/);
    const existingConversations = fx.api.conversations.size;
    const boundary = fx.api.calls.length;
    await runPgnWorkbook([session === "isolated" ? "--resume" : "--restart-run", previous.state.runId], "full", fx.cfg);
    const sends = fx.api.calls.slice(boundary).filter((call) => call.route === "send");
    assert.equal(sends[0].body.body.event.message, session === "isolated" ? "Second A" : "First");
    assert(fx.api.conversations.size > existingConversations);
    const final = await fx.latest();
    assert.equal(final.state.status, "COMPLETED"); assert.equal(final.state.transport, "rest");
    assert.equal(final.state.runId === previous.state.runId, session === "isolated");
  });
}

test("REST validation and explicit smoke use the real client contracts with mocked HTTP", async (context) => {
  const fx = await fixture(context);
  const validated = await validateRest(fx.cfg, ["--test=REST-001"], { fetch: fx.api.fetch });
  assert.equal(validated.ready, true); assert.equal(validated.selectedCount, 1); assert.equal(fx.api.conversations.size, 0);
  assert.equal((await inspectPgnExecution(["--transport=rest"], "full", fx.cfg)).browserRequired, false);
  await assert.rejects(access(fx.cfg.pgnExecutedWorkbookPath));
  await restSmoke(fx.cfg, "Halo", { fetch: fx.api.fetch, now: () => fx.api.now, sleep: fx.api.sleep });
  assert.equal(fx.api.conversations.size, 1); assert([...fx.api.conversations.values()][0].closed);
});

for (const session of ["isolated", "continuous"] as const) {
  test(`REST retest ${session} recovery preserves its selection and applies the proper restart policy`, async (context) => {
    const fx = await fixture(context); fx.api.noResponse.add("First");
    await runPgnWorkbook(["--transport=rest", `--session=${session}`], "retest", fx.cfg);
    const previous = await fx.latest();
    assert.equal(previous.state.status, "RECOVERABLE");
    assert.deepEqual(previous.state.completedScenarioIds, ["REST-002", "REST-NEG-001"]);
    fx.api.noResponse.clear();
    const boundary = fx.api.calls.length;
    await runPgnWorkbook([session === "isolated" ? "--resume" : "--restart-run", previous.state.runId], "retest", fx.cfg);
    const final = await fx.latest();
    const sent = fx.api.calls.slice(boundary).filter((call) => call.route === "send").map((call) => call.body.body.event.message);
    assert.deepEqual(sent, session === "isolated" ? ["First"] : ["First", "Second A", "Second B", "Negative"]);
    assert.deepEqual(final.state.selectedScenarioIds, previous.state.selectedScenarioIds);
    assert.equal(final.state.status, "COMPLETED");
    assert.equal(final.state.transport, "rest");
    const { workbook } = await loadPgnWorkbook(fx.cfg.pgnExecutedWorkbookPath);
    assert.deepEqual(getRetestRunMetadata(workbook, final.state.runId)?.selectedIds, previous.state.selectedScenarioIds);
  });
}

test("REST-only configuration ignores unrelated invalid WhatsApp and Drive settings", () => {
  const cfg = loadConfig({ repositoryRoot: "/tmp/opencode", transport: "rest", environment: { ...environment, WHATSAPP_PROFILE_DIR: "unrelated", WHATSAPP_RESPONSE_TIMEOUT_MS: "invalid", PGN_WHATSAPP_PHONE: "1", GOOGLE_DRIVE_EVIDENCE_PARENT_FOLDER: "bad" } });
  assertRestConfig(cfg.livePersonRest); assert.equal(cfg.googleDriveEvidenceEnabled, false); assert.equal(cfg.target, undefined);
});

test("REST logs and transcript redact credentials even if a bot echoes them", async (context) => {
  const fx = await fixture(context);
  // All actual generated consumer tokens are tested through their full returned value below.
  fx.api.afterPublish = () => {
    const call = fx.api.calls.filter((call) => call.route === "send").at(-1)!;
    const conversation = [...fx.api.conversations.values()].find((item) => item.dialog === call.body.body.dialogId)!;
    conversation.messages.at(-1)!.content.text = `Authorization: ${call.headers.get("authorization")}\nConsumerJWS: ${call.headers.get("X-LP-ON-BEHALF")}`;
  };
  fx.api.responseText = `Echo ${SECRET} ${APP_TOKEN}-1`;
  await runPgnWorkbook(["--transport=rest", "--test=REST-001"], "full", fx.cfg);
  const { workbook } = await loadPgnWorkbook(fx.cfg.pgnExecutedWorkbookPath);
  const serialized = JSON.stringify(workbook.getWorksheet("Execution Transcript")!.getSheetValues());
  for (const secret of [SECRET, APP_TOKEN, CONSUMER_TOKEN]) assert.equal(serialized.includes(secret), false);
  assert.equal(safeRestError(new Error(`Authorization: ${APP_TOKEN}\nclient_secret=${SECRET}`), SECRET).includes(SECRET), false);
});

test("Discord REST lifecycle labels evidence not applicable and never claims screenshot counts", async () => {
  const bodies: any[] = [];
  const notifier = createDiscordNotifier({ discordNotificationsEnabled: true, discordWebhookUrl: "https://discord.com/api/webhooks/123456789/fixture-token", discordProgressEvery: 1, discordProgressMinutes: 1, discordNotifyStart: true, discordNotifyProgress: true, discordNotifyComplete: true, discordNotifyFailure: true }, {
    fetch: async (_url, init) => { bodies.push(JSON.parse(String(init?.body))); return new Response(JSON.stringify({ id: "123456789012345678" })); }, minimumProgressIntervalMs: 0,
  });
  const startedAt = new Date();
  await notifier.runStarted({ runId: "REST", mode: "full", transport: "rest", sessionMode: "isolated", selectedScenarios: 2, startedAt, googleDriveEvidenceEnabled: false, workbookPath: "fixture.xlsx" });
  const progress = { completedScenarios: 1, totalScenarios: 2, capturedScenarios: 1, timeouts: 0, technicalErrors: 0, evidenceUploaded: 0, evidenceUploadErrors: 0, sessionResetAttempts: 0, updatedAt: new Date(startedAt.getTime() + 1000) };
  await notifier.runProgress(progress); await notifier.runCompleted({ ...progress, completedScenarios: 2, completedAt: progress.updatedAt });
  for (const body of bodies) {
    const fields = body.embeds[0].fields;
    assert.equal(fields.find((field: any) => field.name === "Transport")?.value, "REST");
    assert.equal(fields.find((field: any) => field.name === "Evidence")?.value, "Not applicable for REST");
    assert(!fields.some((field: any) => ["Evidence uploaded", "Evidence upload errors", "Session reset attempts"].includes(field.name)));
  }
});
