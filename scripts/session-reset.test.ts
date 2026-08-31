import assert from "node:assert/strict";
import test from "node:test";
import type {
  MessageSnapshot,
  SentMessage,
  WhatsAppMessage,
} from "../src/types";
import {
  BotSessionResetError,
  resetBotSession,
  type BotSessionClient,
} from "../src/whatsapp/session-reset";

function message(
  id: string,
  direction: "incoming" | "outgoing",
  text: string,
  domIndex: number,
): WhatsAppMessage {
  return { id, direction, text, domIndex, observedAt: new Date() };
}

class FakeResetClient implements BotSessionClient {
  readonly sentCommands: string[] = [];
  debugCaptures = 0;
  private readonly messages: WhatsAppMessage[];
  private responseAdded = false;

  constructor(
    initialMessages: WhatsAppMessage[],
    private readonly newResponse?: string,
  ) {
    this.messages = [...initialMessages];
  }

  async captureMessageState(): Promise<MessageSnapshot> {
    return {
      ids: new Set(this.messages.map((item) => item.id)),
      messageCount: this.messages.length,
      messages: [...this.messages],
    };
  }

  async sendMessage(
    command: string,
    _baseline: MessageSnapshot,
  ): Promise<SentMessage> {
    this.sentCommands.push(command);
    const sentAt = new Date();
    this.messages.push(message("outgoing-reset", "outgoing", command, 10));
    return {
      sentAt,
      messageId: "outgoing-reset",
      renderedText: command,
    };
  }

  async getMessages(): Promise<WhatsAppMessage[]> {
    if (this.newResponse && !this.responseAdded) {
      this.responseAdded = true;
      this.messages.push(
        message("incoming-reset", "incoming", this.newResponse, 11),
      );
    }
    return [...this.messages];
  }

  async saveDebugArtifacts(): Promise<{
    screenshotPath: string;
    diagnosticsPath: string;
  }> {
    this.debugCaptures += 1;
    return {
      screenshotPath: "/tmp/reset-failure.png",
      diagnosticsPath: "/tmp/reset-failure.json",
    };
  }
}

test("reset sends the exact command and tolerates confirmation case and whitespace", async () => {
  const client = new FakeResetClient(
    [message("historical", "incoming", "Session deleted", 1)],
    "  SESSION   deleted  ",
  );

  const attempt = await resetBotSession(client, {
    command: "reset",
    confirmation: "Session deleted",
    timeoutMs: 100,
    pollIntervalMs: 1,
    failureArtifactName: "reset-success-test",
  });

  assert.deepEqual(client.sentCommands, ["reset"]);
  assert.equal(attempt.status, "RESET_CONFIRMED");
  assert.equal(attempt.responseMessages.length, 1);
  assert.equal(attempt.responseMessages[0].id, "incoming-reset");
  assert.equal(client.debugCaptures, 0);
});

test("reset ignores historical confirmations and captures timeout diagnostics", async () => {
  const client = new FakeResetClient([
    message("historical", "incoming", "Session deleted", 1),
  ]);

  let failure: unknown;
  try {
    await resetBotSession(client, {
      command: "reset",
      confirmation: "Session deleted",
      timeoutMs: 20,
      pollIntervalMs: 1,
      failureArtifactName: "reset-failure-test",
    });
  } catch (error) {
    failure = error;
  }

  assert(failure instanceof BotSessionResetError);
  assert.equal(failure.attempt.status, "RESET_FAILED");
  assert.equal(failure.attempt.responseMessages.length, 0);
  assert.match(failure.message, /timed out after 20 ms/);
  assert.equal(failure.attempt.evidencePath, "/tmp/reset-failure.png");
  assert.equal(client.debugCaptures, 1);
});
