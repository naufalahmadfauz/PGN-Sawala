import assert from "node:assert/strict";
import test from "node:test";
import type { MessageSnapshot, WhatsAppMessage } from "../src/types";
import {
  collectBotResponseWithEnvironment,
  type ResponseCollectorEnvironment,
} from "../src/whatsapp/response-collector";

interface ScheduledMessage {
  arrivalMs: number;
  message: WhatsAppMessage;
}

interface TypingWindow {
  startMs: number;
  endMs: number;
}

function message(
  id: string,
  direction: "incoming" | "outgoing",
  text: string,
  domIndex: number,
  observedAtMs = 0,
): WhatsAppMessage {
  return {
    id,
    direction,
    text,
    domIndex,
    observedAt: new Date(observedAtMs),
  };
}

class VirtualResponseEnvironment implements ResponseCollectorEnvironment {
  currentTimeMs = 0;
  readonly logs: string[] = [];

  constructor(
    private readonly fixedMessages: WhatsAppMessage[],
    private readonly scheduled: ScheduledMessage[],
    private readonly typingWindows: TypingWindow[] = [],
  ) {}

  async readMessages(): Promise<WhatsAppMessage[]> {
    const latestScheduled = new Map<string, WhatsAppMessage>();
    for (const entry of this.scheduled) {
      if (entry.arrivalMs <= this.currentTimeMs) {
        latestScheduled.set(entry.message.id, entry.message);
      }
    }
    return [
      ...this.fixedMessages,
      ...[...latestScheduled.values()].sort(
        (left, right) => left.domIndex - right.domIndex,
      ),
    ];
  }

  async isRemoteTyping(): Promise<boolean> {
    return this.typingWindows.some(
      ({ startMs, endMs }) =>
        this.currentTimeMs >= startMs && this.currentTimeMs < endMs,
    );
  }

  now(): number {
    return this.currentTimeMs;
  }

  async wait(milliseconds: number): Promise<void> {
    this.currentTimeMs += milliseconds;
  }

  log(messageText: string): void {
    this.logs.push(messageText);
  }
}

const baseline: MessageSnapshot = {
  ids: new Set(["historical"]),
  messageCount: 1,
  messages: [message("historical", "incoming", "Old response", 0)],
};
const fixedMessages = [
  ...baseline.messages,
  message("outgoing-test", "outgoing", "Question", 1),
];

test("collector restarts idle timing for every delayed bot bubble", async () => {
  const environment = new VirtualResponseEnvironment(fixedMessages, [
    { arrivalMs: 2_000, message: message("bot-a", "incoming", "A", 2) },
    { arrivalMs: 6_000, message: message("bot-b", "incoming", "B", 3) },
    { arrivalMs: 11_000, message: message("bot-c", "incoming", "C", 4) },
  ]);

  const response = await collectBotResponseWithEnvironment(
    {
      baseline,
      sentAt: new Date(0),
      outgoingMessageId: "outgoing-test",
      timeoutMs: 30_000,
      idleMs: 10_000,
      pollIntervalMs: 1_000,
      context: "SIMULATED",
    },
    environment,
  );

  assert.equal(response.timedOut, false);
  assert.deepEqual(
    response.messages.map((item) => item.text),
    ["A", "B", "C"],
  );
  assert.equal(response.firstResponseMs, 2_000);
  assert.equal(response.totalResponseMs, 11_000);
  assert.equal(response.completedAt.getTime(), 21_000);
  assert.equal(
    response.combinedResponse,
    "Message 1:\nA\n\nMessage 2:\nB\n\nMessage 3:\nC",
  );
  assert.equal(
    environment.logs.filter((line) => line.includes("Idle timer reset")).length,
    3,
  );
});

test("collector enforces hard timeout without adding idle wait to total time", async () => {
  const environment = new VirtualResponseEnvironment(fixedMessages, [
    { arrivalMs: 2_000, message: message("bot-a", "incoming", "A", 2) },
    { arrivalMs: 9_000, message: message("bot-b", "incoming", "B", 3) },
  ]);

  const response = await collectBotResponseWithEnvironment(
    {
      baseline,
      sentAt: new Date(0),
      outgoingMessageId: "outgoing-test",
      timeoutMs: 15_000,
      idleMs: 10_000,
      pollIntervalMs: 1_000,
    },
    environment,
  );

  assert.equal(response.timedOut, true);
  assert.deepEqual(
    response.messages.map((item) => item.text),
    ["A", "B"],
  );
  assert.equal(response.firstResponseMs, 2_000);
  assert.equal(response.totalResponseMs, 9_000);
  assert.equal(response.completedAt.getTime(), 15_000);
});

test("collector holds completion while typing and restarts quiet timing when typing stops", async () => {
  const environment = new VirtualResponseEnvironment(
    fixedMessages,
    [{ arrivalMs: 2_000, message: message("bot-a", "incoming", "A", 2) }],
    [{ startMs: 11_000, endMs: 14_000 }],
  );

  const response = await collectBotResponseWithEnvironment(
    {
      baseline,
      sentAt: new Date(0),
      outgoingMessageId: "outgoing-test",
      timeoutMs: 30_000,
      idleMs: 10_000,
      pollIntervalMs: 1_000,
    },
    environment,
  );

  assert.equal(response.timedOut, false);
  assert.equal(response.totalResponseMs, 2_000);
  assert.equal(response.completedAt.getTime(), 24_000);
  assert(environment.logs.some((line) => line.includes("Bot is typing")));
  assert(environment.logs.some((line) => line.includes("Typing stopped")));
});

test("collector permanently excludes a control bubble even when its text changes", async () => {
  const environment = new VirtualResponseEnvironment(fixedMessages, [
    {
      arrivalMs: 2_000,
      message: message("control", "incoming", "Session", 2),
    },
    {
      arrivalMs: 3_000,
      message: message("control", "incoming", "Session deleted", 2),
    },
    {
      arrivalMs: 4_000,
      message: message("control", "incoming", "Changed control text", 2),
    },
    { arrivalMs: 6_000, message: message("bot-a", "incoming", "A", 3) },
  ]);

  const response = await collectBotResponseWithEnvironment(
    {
      baseline,
      sentAt: new Date(0),
      outgoingMessageId: "outgoing-test",
      timeoutMs: 30_000,
      idleMs: 10_000,
      pollIntervalMs: 1_000,
      excludedIncomingTexts: ["Session deleted"],
    },
    environment,
  );

  assert.equal(response.timedOut, false);
  assert.deepEqual(response.messages.map((item) => item.text), ["A"]);
  assert.equal(response.firstResponseMs, 6_000);
  assert.equal(response.totalResponseMs, 6_000);
  assert.equal(response.completedAt.getTime(), 16_000);
});

test("collector leaves response timing empty when no bot bubble arrives", async () => {
  const environment = new VirtualResponseEnvironment(fixedMessages, []);

  const response = await collectBotResponseWithEnvironment(
    {
      baseline,
      sentAt: new Date(0),
      outgoingMessageId: "outgoing-test",
      timeoutMs: 15_000,
      idleMs: 10_000,
      pollIntervalMs: 1_000,
    },
    environment,
  );

  assert.equal(response.timedOut, true);
  assert.equal(response.firstResponseMs, undefined);
  assert.equal(response.totalResponseMs, undefined);
  assert.equal(response.completedAt.getTime(), 15_000);
});
