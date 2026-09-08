import type { PgnTestScenario, PgnTestTurn } from "../excel/pgn-types";
import { randomUUID } from "node:crypto";
import type { SessionMode } from "../session-mode";
import type { LivePersonRestConfig } from "../rest/config";
import { RestApiError } from "../rest/errors";
import { LivePersonClient, type RestConversation, type RestDependencies, type RestMessage } from "../rest/liveperson-client";
import type { TestTransport, TransportResponse } from "./test-transport";

const BOT_ROLES = new Set(["AGENT", "ASSIGNED_AGENT", "BRAND_BOT", "BOT", "MANAGER"]);

export class RestTransport implements TestTransport {
  readonly type = "rest";
  readonly client: LivePersonClient;
  #conversation?: RestConversation;
  #cursor = -1;
  #stopped = false;
  readonly #attemptIdentity = randomUUID();
  readonly #pendingCleanup = new Map<string, RestConversation>();

  get pendingConversationCount(): number { return this.#pendingCleanup.size; }

  constructor(
    private readonly config: LivePersonRestConfig,
    private readonly runId: string,
    private readonly sessionMode: SessionMode,
    private readonly warn: (message: string) => void = console.warn,
    dependencies: RestDependencies = {},
  ) { this.client = new LivePersonClient(config, dependencies); }

  async initializeRun(): Promise<void> {
    await this.client.discoverDomains();
    await this.client.applicationToken();
    if (this.sessionMode === "continuous") await this.#create("continuous");
  }

  async #create(scope: string): Promise<void> {
    if (this.#stopped) throw new RestApiError("REST execution interrupted");
    // A new process/recovery attempt must not reattach an existing consumer conversation.
    const consumer = this.client.syntheticConsumer(this.runId, `${this.#attemptIdentity}|${scope}`);
    const conversation = await this.client.createConversation(consumer);
    if (this.#stopped) {
      await this.#close(conversation, true);
      throw new RestApiError("REST execution interrupted");
    }
    this.#conversation = conversation;
    this.#pendingCleanup.set(this.#conversation.conversationId, this.#conversation);
    this.#cursor = -1;
  }

  async beginScenario(scenario: PgnTestScenario, index: number): Promise<void> {
    if (this.sessionMode === "isolated") await this.#create(`${index}:${scenario.testCaseId}`);
    if (!this.#conversation || this.#stopped) throw new RestApiError("No active REST conversation is available");
  }

  async sendMessage(_scenario: PgnTestScenario, turn: PgnTestTurn): Promise<TransportResponse> {
    const conversation = this.#conversation;
    if (!conversation || this.#stopped) throw new RestApiError("No active REST conversation is available");
    const start = this.client.now();
    let sentAt: Date | undefined;
    let sending = false;
    let firstAt: number | undefined;
    let lastAt: number | undefined;
    const collected = new Map<number, RestMessage>();
    const result = (technicalStatus: TransportResponse["technicalStatus"], error?: string): TransportResponse => ({
      transport: "rest", technicalStatus, sentAt, completedAt: new Date(this.client.now()),
      conversationId: conversation.conversationId, dialogId: conversation.dialogId,
      botMessages: [...collected.values()].sort((a, b) => a.sequence - b.sequence).map((message, index) => ({
        sequence: index + 1, message: message.text!, timestamp: message.receivedAt,
      })),
      combinedResponse: [...collected.values()].sort((a, b) => a.sequence - b.sequence).map((message) => message.text).join("\n\n"),
      firstResponseMs: firstAt === undefined || !sentAt ? undefined : firstAt - sentAt.getTime(),
      totalResponseMs: lastAt === undefined || !sentAt ? undefined : lastAt - sentAt.getTime(),
      error, evidenceStatus: "EVIDENCE_NOT_APPLICABLE",
    });
    try {
      // Drain all prior events before the publish boundary, including greeting/consumer events.
      const previous = await this.client.messages(conversation, this.#cursor, start + this.config.responseTimeoutMs);
      for (const message of previous) this.#cursor = Math.max(this.#cursor, message.sequence);
      const boundary = this.#cursor;
      sentAt = new Date(this.client.now());
      sending = true;
      const publishedSequence = await this.client.sendText(conversation, turn.userInput, sentAt.getTime() + this.config.responseTimeoutMs);
      sending = false;
      this.#cursor = Math.max(boundary, publishedSequence);
      const deadline = sentAt.getTime() + this.config.responseTimeoutMs;
      while (this.client.now() < deadline) {
        if (this.#stopped) throw new RestApiError("REST execution interrupted");
        const messages = await this.client.messages(conversation, this.#cursor, deadline);
        for (const message of messages) {
          if (message.sequence <= this.#cursor) continue;
          if (BOT_ROLES.has(message.role) && message.text?.trim() && !collected.has(message.sequence)) {
            collected.set(message.sequence, message);
            firstAt ??= this.client.now();
            lastAt = this.client.now();
          }
        }
        for (const message of messages) this.#cursor = Math.max(this.#cursor, message.sequence);
        if (lastAt !== undefined && this.client.now() - lastAt >= this.config.responseIdleMs) return result("CAPTURED");
        await this.client.sleep(Math.min(this.config.pollIntervalMs, Math.max(0, deadline - this.client.now())));
      }
      return result("TIMEOUT", `REST response timed out after ${this.config.responseTimeoutMs} ms`);
    } catch (error) {
      if (this.#stopped || !(error instanceof RestApiError) || error.global) throw error;
      if (error.status === 408) return result("TIMEOUT", `REST response timed out after ${this.config.responseTimeoutMs} ms`);
      return result(sending ? "SEND_ERROR" : "CHAT_ERROR", this.client.redact(error));
    }
  }

  async #close(conversation: RestConversation, cleanup = false): Promise<void> {
    try {
      await this.client.closeConversation(conversation, cleanup);
      this.#pendingCleanup.delete(conversation.conversationId);
    } catch (error) {
      this.warn(`[REST cleanup] ${this.client.redact(error)}; captured results remain saved.`);
    }
  }

  async endScenario(_scenario: PgnTestScenario): Promise<void> {
    if (this.sessionMode === "isolated" && this.#conversation) {
      await this.#close(this.#conversation);
      this.#conversation = undefined;
    }
  }

  async finalizeRun(): Promise<void> {
    for (const conversation of this.#pendingCleanup.values()) await this.#close(conversation);
    this.#conversation = undefined;
  }

  async close(): Promise<void> {
    if (this.#stopped) return;
    this.#stopped = true;
    this.client.cancel();
    for (const conversation of this.#pendingCleanup.values()) await this.#close(conversation, true);
    this.#conversation = undefined;
  }
}
