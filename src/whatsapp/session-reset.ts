import type {
  BotSessionResetAttempt,
  MessageSnapshot,
  SentMessage,
  WhatsAppMessage,
} from "../types";

export interface BotSessionClient {
  captureMessageState(): Promise<MessageSnapshot>;
  sendMessage(message: string, baseline: MessageSnapshot): Promise<SentMessage>;
  getMessages(): Promise<WhatsAppMessage[]>;
  saveDebugArtifacts(
    name: string,
  ): Promise<{ screenshotPath: string; diagnosticsPath: string }>;
}

export interface ResetBotSessionOptions {
  command: string;
  confirmation: string;
  timeoutMs: number;
  failureArtifactName: string;
  pollIntervalMs?: number;
}

function normalizeMessage(value: string): string {
  return value.replace(/\s+/g, " ").trim().toLowerCase();
}

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

function wait(milliseconds: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, milliseconds));
}

export class BotSessionResetError extends Error {
  constructor(public readonly attempt: BotSessionResetAttempt) {
    super(attempt.error ?? "PGN bot session reset failed");
    this.name = "BotSessionResetError";
  }
}

export async function resetBotSession(
  client: BotSessionClient,
  options: ResetBotSessionOptions,
): Promise<BotSessionResetAttempt> {
  const command = options.command.trim();
  const confirmation = options.confirmation.trim();
  if (!command) {
    throw new Error("Reset command must not be empty");
  }
  if (!confirmation) {
    throw new Error("Reset confirmation must not be empty");
  }
  if (!Number.isInteger(options.timeoutMs) || options.timeoutMs < 1) {
    throw new Error("Reset timeout must be a positive integer");
  }
  const pollIntervalMs = options.pollIntervalMs ?? 250;
  if (!Number.isInteger(pollIntervalMs) || pollIntervalMs < 1) {
    throw new Error("Reset poll interval must be a positive integer");
  }

  const startedAt = new Date();
  const captured = new Map<string, WhatsAppMessage>();
  let firstResponseAt: Date | undefined;
  let sentMessage: SentMessage | undefined;

  const fail = async (message: string): Promise<never> => {
    const completedAt = new Date();
    let evidencePath: string | undefined;
    let diagnosticsPath: string | undefined;
    try {
      const debug = await client.saveDebugArtifacts(options.failureArtifactName);
      evidencePath = debug.screenshotPath;
      diagnosticsPath = debug.diagnosticsPath;
    } catch (error) {
      console.error(
        `[Session] Could not save reset failure diagnostics: ${errorMessage(error)}`,
      );
    }

    const attempt: BotSessionResetAttempt = {
      command,
      expectedConfirmation: confirmation,
      status: "RESET_FAILED",
      startedAt,
      sentAt: sentMessage?.sentAt,
      completedAt,
      responseMessages: [...captured.values()].sort(
        (left, right) => left.domIndex - right.domIndex,
      ),
      firstResponseMs:
        sentMessage && firstResponseAt
          ? firstResponseAt.getTime() - sentMessage.sentAt.getTime()
          : undefined,
      totalResponseMs: sentMessage
        ? completedAt.getTime() - sentMessage.sentAt.getTime()
        : undefined,
      error: message,
      evidencePath,
      diagnosticsPath,
    };
    console.error(`[Session] ERROR: ${message}`);
    throw new BotSessionResetError(attempt);
  };

  console.log("[Session] Resetting bot...");
  let baseline: MessageSnapshot;
  try {
    baseline = await client.captureMessageState();
  } catch (error) {
    return fail(`Could not snapshot WhatsApp messages: ${errorMessage(error)}`);
  }

  console.log(`[Session] Sending: ${command}`);
  try {
    sentMessage = await client.sendMessage(command, baseline);
  } catch (error) {
    return fail(`Could not send reset command: ${errorMessage(error)}`);
  }

  const expected = normalizeMessage(confirmation);
  const deadline = Date.now() + options.timeoutMs;

  do {
    let messages: WhatsAppMessage[];
    try {
      messages = await client.getMessages();
    } catch (error) {
      return fail(`Could not read reset response: ${errorMessage(error)}`);
    }
    const outgoingAnchor = messages.find(
      (message) =>
        message.direction === "outgoing" &&
        message.id === sentMessage!.messageId,
    );
    const newIncoming = messages.filter(
      (message) =>
        Boolean(outgoingAnchor) &&
        message.direction === "incoming" &&
        message.text.length > 0 &&
        !baseline.ids.has(message.id) &&
        message.domIndex > outgoingAnchor!.domIndex,
    );

    for (const message of newIncoming) {
      const existing = captured.get(message.id);
      if (!existing) {
        const observedAt = new Date();
        captured.set(message.id, { ...message, observedAt });
        firstResponseAt ??= observedAt;
        console.log(`[Session] Response: ${message.text}`);
      } else if (existing.text !== message.text) {
        captured.set(message.id, {
          ...message,
          observedAt: existing.observedAt,
        });
      }
    }

    const responseMessages = [...captured.values()].sort(
      (left, right) => left.domIndex - right.domIndex,
    );
    const combinedResponse = responseMessages
      .map((message) => message.text)
      .join("\n");
    if (normalizeMessage(combinedResponse).includes(expected)) {
      const completedAt = new Date();
      console.log("[Session] Reset confirmed");
      return {
        command,
        expectedConfirmation: confirmation,
        status: "RESET_CONFIRMED",
        startedAt,
        sentAt: sentMessage.sentAt,
        completedAt,
        responseMessages,
        firstResponseMs: firstResponseAt
          ? firstResponseAt.getTime() - sentMessage.sentAt.getTime()
          : undefined,
        totalResponseMs: completedAt.getTime() - sentMessage.sentAt.getTime(),
      };
    }

    const remainingMs = deadline - Date.now();
    if (remainingMs > 0) {
      await wait(Math.min(pollIntervalMs, remainingMs));
    }
  } while (Date.now() < deadline);

  return fail(
    `Reset confirmation timed out after ${options.timeoutMs} ms; expected "${confirmation}"`,
  );
}
