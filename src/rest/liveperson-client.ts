import { createHash, randomUUID } from "node:crypto";
import { assertRestConfig, domainOrigin, LIVEPERSON_DOMAINS, type LivePersonDomain, type LivePersonRestConfig } from "./config";
import { RestApiError, safeRestError, redactRestText } from "./errors";

export interface RestDependencies {
  fetch?: typeof fetch;
  now?: () => number;
  sleep?: (milliseconds: number) => Promise<void>;
}
export interface RestConversation { conversationId: string; dialogId: string; consumerId: string }
export interface RestMessage { sequence: number; text?: string; role: string; receivedAt: Date }
type JsonObject = Record<string, unknown>;
function object(value: unknown): JsonObject {
  if (!value || typeof value !== "object" || Array.isArray(value)) throw new RestApiError("LivePerson returned an invalid response object");
  return value as JsonObject;
}
function identifier(value: unknown): string {
  if (typeof value !== "string" || !/^[A-Za-z0-9._:-]{1,256}$/.test(value)) throw new RestApiError("LivePerson returned an invalid operational identifier");
  return value;
}
function sequence(value: unknown): number {
  if (typeof value !== "number" && (typeof value !== "string" || !/^\d+$/.test(value))) throw new RestApiError("LivePerson returned an invalid message sequence");
  const parsed = Number(value);
  if (!Number.isSafeInteger(parsed) || parsed < 0) throw new RestApiError("LivePerson returned an invalid message sequence");
  return parsed;
}

export class LivePersonClient {
  readonly #config: LivePersonRestConfig & { accountId: string; clientId: string; clientSecret: string; skillId: string };
  readonly #fetch: typeof fetch;
  readonly now: () => number;
  readonly sleep: (milliseconds: number) => Promise<void>;
  readonly #domains = new Map<LivePersonDomain, string>();
  readonly #controllers = new Set<AbortController>();
  #app?: { token: string; expires: number };
  #consumer?: { id: string; token: string; expires: number };
  readonly #consumers = new Map<string, { id: string; token: string; expires: number }>();
  readonly #secrets = new Set<string>();
  #cancelled = false;
  #commandSequence = 0;

  constructor(config: LivePersonRestConfig, dependencies: RestDependencies = {}) {
    assertRestConfig(config);
    this.#config = config;
    this.#secrets.add(config.clientSecret);
    this.#fetch = dependencies.fetch ?? globalThis.fetch;
    this.now = dependencies.now ?? Date.now;
    this.sleep = dependencies.sleep ?? ((ms) => new Promise((resolve) => setTimeout(resolve, ms)));
    for (const service of LIVEPERSON_DOMAINS) {
      const override = config.domains[service];
      if (override) this.#domains.set(service, domainOrigin(override));
    }
  }

  redact(error: unknown): string {
    return safeRestError(error, ...this.#secrets);
  }

  cancel(): void {
    this.#cancelled = true;
    for (const controller of this.#controllers) controller.abort();
  }

  async #request(
    url: string, init: RequestInit, operation: string,
    options: { retrySafe?: boolean; scenario?: boolean; deadline?: number; cleanup?: boolean } = {},
  ): Promise<unknown> {
    for (let attempt = 0; attempt < 3; attempt += 1) {
      if (this.#cancelled && !options.cleanup) throw new RestApiError("REST execution interrupted");
      const remaining = options.deadline === undefined ? this.#config.requestTimeoutMs : options.deadline - this.now();
      if (remaining <= 0) throw new RestApiError("REST response deadline exceeded", false, 408);
      const controller = new AbortController();
      this.#controllers.add(controller);
      let expire!: (error: Error) => void;
      const expired = new Promise<never>((_resolve, reject) => { expire = reject; });
      const timer = setTimeout(() => {
        controller.abort();
        expire(new Error("REST request deadline exceeded"));
      }, Math.min(this.#config.requestTimeoutMs, remaining));
      let status: number | undefined;
      let retryAfter: string | null = null;
      try {
        const response = await Promise.race([this.#fetch(url, { ...init, redirect: "error", signal: controller.signal }), expired]);
        status = response.status;
        retryAfter = response.headers.get("retry-after");
        if (response.ok) {
          if (status === 204) return {};
          // Keep the deadline active while consuming the body, not just until headers arrive.
          const text = await Promise.race([response.text(), expired]);
          if (!text.trim()) return {};
          try { return JSON.parse(text) as unknown; }
          catch { throw new RestApiError(`LivePerson ${operation} returned invalid JSON`); }
        }
        await response.body?.cancel();
      } catch (error) {
        if (error instanceof RestApiError) throw error;
        status = undefined;
        if (this.#cancelled && !options.cleanup) throw new RestApiError("REST execution interrupted");
        if (options.deadline !== undefined && this.now() >= options.deadline) throw new RestApiError("REST response deadline exceeded", false, 408);
        if (!options.retrySafe || attempt === 2) {
          throw new RestApiError(`LivePerson ${operation} network/timeout failure${options.retrySafe ? " after bounded retries" : "; delivery is uncertain and was not retried"}`, Boolean(options.retrySafe) || !options.scenario);
        }
      } finally {
        clearTimeout(timer);
        this.#controllers.delete(controller);
      }
      const retriable = status === undefined || status === 429 || (options.retrySafe && [500, 502, 503, 504].includes(status));
      if (!retriable || attempt === 2) {
        const global = status === 401 || status === 403 || status === 429 || Boolean(options.retrySafe && status !== undefined && status >= 500) || !options.scenario;
        throw new RestApiError(`LivePerson ${operation} failed${status ? ` (HTTP ${status})` : ""}${status && status >= 500 && !options.retrySafe ? "; delivery is uncertain and was not retried" : ""}`, global, status);
      }
      const numericDelay = retryAfter === null ? NaN : Number(retryAfter);
      const headerDelay = retryAfter === null ? 0 : Number.isFinite(numericDelay) ? numericDelay * 1000 : Date.parse(retryAfter) - this.now();
      const delay = Math.max(250 * 2 ** attempt, Number.isFinite(headerDelay) ? headerDelay : 0);
      // Refuse waits outside the retry budget rather than violating a long Retry-After.
      if (delay > 30000) throw new RestApiError(`LivePerson ${operation} exceeded the retry wait budget`, true, status);
      if (options.deadline !== undefined && this.now() + delay >= options.deadline) throw new RestApiError("REST response deadline exceeded", false, 408);
      await this.sleep(delay);
    }
    throw new RestApiError("LivePerson request exhausted retries");
  }

  async discoverDomains(): Promise<void> {
    if (LIVEPERSON_DOMAINS.every((name) => this.#domains.has(name))) return;
    const response = object(await this.#request(
      `https://api.liveperson.net/api/account/${encodeURIComponent(this.#config.accountId)}/service/baseURI.json?version=1.0`,
      { method: "GET" }, "domain discovery", { retrySafe: true },
    ));
    if (!Array.isArray(response.baseURIs)) throw new RestApiError("LivePerson domain discovery returned no domain list");
    for (const service of LIVEPERSON_DOMAINS) {
      if (this.#domains.has(service)) continue;
      const matches = response.baseURIs.map(object).filter((item) => item.service === service);
      if (matches.length !== 1 || typeof matches[0].baseURI !== "string") throw new RestApiError(`LivePerson domain discovery is missing or ambiguous for ${service}`);
      const origin = domainOrigin(matches[0].baseURI);
      if (!/\.(?:liveperson\.net|liveperson\.com)$/.test(new URL(origin).hostname)) throw new RestApiError("LivePerson domain discovery returned an untrusted hostname; use an explicit reviewed override if needed");
      this.#domains.set(service, origin);
    }
  }

  #expiry(token: string, expiresIn: unknown, fallbackSeconds: number): number {
    let seconds = typeof expiresIn === "number" || typeof expiresIn === "string" ? Number(expiresIn) : fallbackSeconds;
    if (!Number.isFinite(seconds) || seconds <= 0) seconds = fallbackSeconds;
    let expires = this.now() + seconds * 1000;
    try {
      const payload = JSON.parse(Buffer.from(token.split(".")[1], "base64url").toString("utf8"));
      if (typeof payload.exp === "number" && Number.isFinite(payload.exp)) expires = Math.min(expires, payload.exp * 1000);
    } catch { /* Expiry decoding is only a cache hint, not token verification. */ }
    return expires;
  }

  async applicationToken(deadline?: number): Promise<string> {
    if (this.#app && this.#app.expires > this.now() + 30000) return this.#app.token;
    await this.discoverDomains();
    const response = object(await this.#request(
      `${this.#domains.get("sentinel")}/sentinel/api/account/${encodeURIComponent(this.#config.accountId)}/app/token?v=1.0&grant_type=client_credentials`,
      { method: "POST", headers: { "Content-Type": "application/x-www-form-urlencoded" }, body: new URLSearchParams({ client_id: this.#config.clientId, client_secret: this.#config.clientSecret }).toString() },
      "application authentication", { retrySafe: true, deadline },
    ));
    if (typeof response.access_token !== "string" || !response.access_token) throw new RestApiError("LivePerson application authentication returned no token");
    this.#app = { token: response.access_token, expires: this.#expiry(response.access_token, response.expires_in, 300) };
    this.#secrets.add(response.access_token);
    return this.#app.token;
  }

  async consumerToken(consumerId: string, deadline?: number): Promise<string> {
    const cached = this.#consumers.get(consumerId);
    if (cached && cached.expires > this.now() + 30000) return cached.token;
    for (let refresh = 0; refresh < 2; refresh += 1) {
      try {
        const app = await this.applicationToken(deadline);
        const response = object(await this.#request(
          `${this.#domains.get("idp")}/api/account/${encodeURIComponent(this.#config.accountId)}/consumer?v=1.0`,
          { method: "POST", headers: { Authorization: app, "Content-Type": "application/json" }, body: JSON.stringify({ ext_consumer_id: consumerId }) },
          "consumer authentication", { retrySafe: true, deadline },
        ));
        if (typeof response.token !== "string" || !response.token) throw new RestApiError("LivePerson consumer authentication returned no token");
        this.#consumer = { id: consumerId, token: response.token, expires: this.#expiry(response.token, response.expires_in, 900) };
        this.#consumers.set(consumerId, this.#consumer);
        this.#secrets.add(response.token);
        return this.#consumer.token;
      } catch (error) {
        if (!(error instanceof RestApiError) || error.status !== 401 || refresh) throw error;
        this.#app = undefined;
      }
    }
    throw new RestApiError("LivePerson consumer authentication failed");
  }

  syntheticConsumer(runId: string, scope: string): string {
    return `pgn-sawala-${createHash("sha256").update(`${runId}|${scope}`).digest("hex").slice(0, 40)}`;
  }

  async validate(): Promise<void> {
    await this.discoverDomains();
    await this.applicationToken();
    await this.consumerToken(this.syntheticConsumer("validation", randomUUID()));
  }

  async #authorized(
    service: "asyncMessagingEnt" | "messagingRestApiDomain", suffix: string, consumerId: string,
    init: RequestInit, operation: string,
    options: { retrySafe?: boolean; scenario?: boolean; deadline?: number; cleanup?: boolean } = {},
  ): Promise<unknown> {
    const requestId = randomUUID();
    for (let refresh = 0; refresh < 2; refresh += 1) {
      try {
        const app = options.cleanup && this.#app ? this.#app.token : await this.applicationToken(options.deadline);
        const cachedConsumer = this.#consumers.get(consumerId);
        const consumer = options.cleanup && cachedConsumer ? cachedConsumer.token : await this.consumerToken(consumerId, options.deadline);
        return await this.#request(`${this.#domains.get(service)}${suffix}`, {
          ...init,
          headers: {
            ...init.headers,
            Authorization: app,
            [service === "asyncMessagingEnt" ? "X-LP-ON-BEHALF" : "LP-ON-BEHALF"]: consumer,
            "Brand-ID": this.#config.accountId, "Client-source": "pgn-sawala-rest", "Request-ID": requestId,
            "Content-Type": "application/json",
          },
        }, operation, options);
      } catch (error) {
        if (!(error instanceof RestApiError) || error.status !== 401 || refresh || options.cleanup) throw error;
        this.#app = undefined;
        this.#consumer = undefined;
        this.#consumers.delete(consumerId);
      }
    }
    throw new RestApiError("LivePerson authentication refresh failed");
  }

  #commandBody(response: unknown, requestId: string): JsonObject {
    const candidates = (Array.isArray(response) ? response : [response]).map(object);
    const replies = candidates.filter((item) => String(item.reqId ?? item.id ?? "") === requestId);
    const reply = replies.length === 1 ? replies[0] : candidates.length === 1 && candidates[0].reqId === undefined ? candidates[0] : undefined;
    if (!reply) throw new RestApiError("LivePerson command response could not be correlated safely");
    if (reply.code !== undefined && reply.code !== "OK" && reply.code !== 200 && reply.code !== 201 && reply.code !== "200" && reply.code !== "201") {
      const status = Number(reply.code);
      throw new RestApiError("LivePerson command was rejected", true, Number.isInteger(status) ? status : undefined);
    }
    return object(reply.body ?? reply);
  }

  #identifier(value: unknown): string {
    const id = identifier(value);
    if (this.#secrets.has(id) || /^eyJ[A-Za-z0-9_-]+\./.test(id)) {
      throw new RestApiError("LivePerson returned credential material instead of an operational identifier");
    }
    return id;
  }

  async createConversation(consumerId: string): Promise<RestConversation> {
    const profileRequestId = String(++this.#commandSequence);
    const requestId = String(++this.#commandSequence);
    const response = await this.#authorized("asyncMessagingEnt",
      `/api/account/${encodeURIComponent(this.#config.accountId)}/messaging/consumer/conversation?v=3`, consumerId,
      { method: "POST", body: JSON.stringify([
        { kind: "req", id: profileRequestId, type: "userprofile.SetUserProfile", body: { authenticatedData: { lp_sdes: [{ type: "personal", personal: { firstname: "PGN", lastname: "REST-Test" } }] } } },
        { kind: "req", id: requestId, type: "cm.ConsumerRequestConversation", body: { brandId: this.#config.accountId, skillId: this.#config.skillId, channelType: "MESSAGING", ttrDefName: "NORMAL" } },
      ]) }, "conversation creation",
    );
    const body = this.#commandBody(response, requestId);
    const conversationId = this.#identifier(body.conversationId ?? body.id);
    try {
      if (body.dialogId) return { consumerId, conversationId, dialogId: this.#identifier(body.dialogId) };
      const details = Array.isArray(body.dialogs) ? body : object(await this.#authorized("messagingRestApiDomain", `/messaging/v1/conversations/${encodeURIComponent(conversationId)}`, consumerId, { method: "GET" }, "conversation lookup", { retrySafe: true }));
      const dialogs = Array.isArray(details.dialogs) ? details.dialogs.map(object).filter((item) => item.dialogType === "MAIN" && item.state === "OPEN") : [];
      if (dialogs.length !== 1) throw new RestApiError("LivePerson conversation has no unambiguous open MAIN dialog");
      return { consumerId, conversationId, dialogId: this.#identifier(dialogs[0].id) };
    } catch (error) {
      await this.closeConversation({ consumerId, conversationId, dialogId: conversationId }).catch(() => undefined);
      throw error;
    }
  }

  async sendText(conversation: RestConversation, text: string, deadline?: number): Promise<number> {
    const requestId = String(++this.#commandSequence);
    const response = await this.#authorized("asyncMessagingEnt",
      `/api/account/${encodeURIComponent(this.#config.accountId)}/messaging/consumer/conversation/send?v=3`, conversation.consumerId,
      { method: "POST", body: JSON.stringify({ kind: "req", id: requestId, type: "ms.PublishEvent", body: { dialogId: conversation.dialogId, event: { type: "ContentEvent", contentType: "text/plain", message: text } } }) },
      "text publish", { scenario: true, deadline },
    );
    const body = this.#commandBody(response, requestId);
    return sequence(body.sequence);
  }

  async messages(conversation: RestConversation, after: number, deadline?: number): Promise<RestMessage[]> {
    const messages: RestMessage[] = [];
    let cursor = after;
    for (let page = 0; page < 100; page += 1) {
      const query = new URLSearchParams({ newerThanSequence: String(cursor + 1), sortBy: "sequence", sortOrder: "ASC", limit: "100" });
      const response = object(await this.#authorized("messagingRestApiDomain",
        `/messaging/v1/conversations/${encodeURIComponent(conversation.conversationId)}/dialogs/${encodeURIComponent(conversation.dialogId)}/messages?${query}`,
        conversation.consumerId, { method: "GET" }, "message polling", { retrySafe: true, scenario: true, deadline },
      ));
      if (!Array.isArray(response.data)) throw new RestApiError("LivePerson message response is missing its data array");
      let highWater = cursor;
      for (const raw of response.data) {
        const item = object(raw);
        const seq = sequence(item.sequence);
        if (seq <= cursor) continue;
        if (item.dialogId !== undefined && item.dialogId !== conversation.dialogId) throw new RestApiError("LivePerson returned a message from a different dialog");
        const originator = item.originator && typeof item.originator === "object" ? object(item.originator) : {};
        const content = item.content && typeof item.content === "object" ? object(item.content) : {};
        const role = typeof originator.role === "string" ? originator.role : "";
        const text = item.type === "PLAIN_TEXT" && typeof content.text === "string" && (!item.messageAudience || item.messageAudience === "ALL") ? redactRestText(content.text, ...this.#secrets) : undefined;
        const receivedAt = typeof item.createdTs === "string" ? new Date(item.createdTs) : new Date(this.now());
        messages.push({ sequence: seq, role, text, receivedAt: Number.isNaN(receivedAt.getTime()) ? new Date(this.now()) : receivedAt });
        highWater = Math.max(highWater, seq);
      }
      const hasNext = response.links && typeof response.links === "object" && typeof object(response.links).next === "string";
      if (response.data.length < 100 && !hasNext) return messages.sort((a, b) => a.sequence - b.sequence);
      if (highWater === cursor) throw new RestApiError("LivePerson message pagination did not advance");
      cursor = highWater;
    }
    throw new RestApiError("LivePerson message pagination exceeded its bounded page limit");
  }

  async closeConversation(conversation: RestConversation, cleanup = false): Promise<void> {
    const requestId = String(++this.#commandSequence);
    const response = await this.#authorized("asyncMessagingEnt",
      `/api/account/${encodeURIComponent(this.#config.accountId)}/messaging/consumer/conversation/send?v=3`, conversation.consumerId,
      { method: "POST", body: JSON.stringify({ kind: "req", id: requestId, type: "cm.UpdateConversationField", body: { conversationId: conversation.conversationId, conversationField: { field: "ConversationStateField", conversationState: "CLOSE" } } }) },
      "conversation close", { retrySafe: true, cleanup },
    );
    if (Array.isArray(response) || Object.keys(object(response)).length) this.#commandBody(response, requestId);
  }
}
