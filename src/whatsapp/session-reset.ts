import type {
  BotSessionResetAttempt,
  MessageSnapshot,
  PostResetDrainResult,
  SentMessage,
  WhatsAppMessage,
} from "../types";

export interface BotSessionClient {
  captureMessageState(): Promise<MessageSnapshot>;
  sendMessage(message: string, baseline: MessageSnapshot): Promise<SentMessage>;
  getMessages(): Promise<WhatsAppMessage[]>;
  isRemoteTyping?(): Promise<boolean>;
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

export interface PostResetQuietOptions {
  baselineMessages: WhatsAppMessage[];
  quietMs: number;
  pollIntervalMs?: number;
  now?: () => number;
  wait?: (milliseconds: number) => Promise<void>;
  log?: (message: string) => void;
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
      messageStateAtCompletion: [...latestMessages],
    };
    console.error(`[Session] ERROR: ${message}`);
    throw new BotSessionResetError(attempt);
  };

  console.log("[Session] Resetting bot...");
  let baseline: MessageSnapshot;
  let latestMessages: WhatsAppMessage[] = [];
  try {
    baseline = await client.captureMessageState();
    latestMessages = [...baseline.messages];
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
      latestMessages = messages;
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
        messageStateAtCompletion: [...messages],
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

export async function waitForPostResetQuiet(
  client: Pick<BotSessionClient, "getMessages" | "isRemoteTyping">,
  options: PostResetQuietOptions,
): Promise<PostResetDrainResult> {
  if (!Number.isInteger(options.quietMs) || options.quietMs < 1) {
    throw new Error("Post-reset quiet window must be a positive integer");
  }
  const pollIntervalMs = options.pollIntervalMs ?? 250;
  if (!Number.isInteger(pollIntervalMs) || pollIntervalMs < 1) {
    throw new Error("Post-reset poll interval must be a positive integer");
  }

  const now = options.now ?? (() => Date.now());
  const pause = options.wait ?? wait;
  const log = options.log ?? ((message: string) => console.log(message));
  const startedAtMs = now();
  let quietSinceMs = startedAtMs;
  const knownMessages = new Map(
    options.baselineMessages.map((message) => [message.id, message.text]),
  );
  const staleMessages: WhatsAppMessage[] = [];
  let typingWasVisible = false;

  log(
    `[Session] Waiting for post-reset quiet period: ${options.quietMs} ms`,
  );
  while (true) {
    const messages = await client.getMessages();
    for (const message of messages) {
      const knownText = knownMessages.get(message.id);
      knownMessages.set(message.id, message.text);
      if (
        message.direction !== "incoming" ||
        !message.text ||
        knownText === message.text
      ) {
        continue;
      }

      const observedAtMs = now();
      const staleMessage = {
        ...message,
        observedAt: new Date(observedAtMs),
      };
      staleMessages.push(staleMessage);
      quietSinceMs = observedAtMs;
      log(`[Session] STALE BOT MESSAGE:\n"${message.text}"`);
      log(`[Session] Quiet timer restarted: ${options.quietMs} ms`);
    }

    const typing = (await client.isRemoteTyping?.()) ?? false;
    const currentTime = now();
    if (typing) {
      quietSinceMs = currentTime;
      if (!typingWasVisible) {
        log("[Session] Bot is typing; post-reset quiet timer held");
      }
    } else if (typingWasVisible) {
      quietSinceMs = currentTime;
      log(
        `[Session] Typing stopped; quiet timer restarted: ${options.quietMs} ms`,
      );
    }
    typingWasVisible = typing;

    if (!typing && currentTime - quietSinceMs >= options.quietMs) {
      log("[Session] Post-reset quiet period complete");
      return {
        startedAt: new Date(startedAtMs),
        completedAt: new Date(currentTime),
        quietMs: options.quietMs,
        staleMessages,
      };
    }
    await pause(
      Math.min(pollIntervalMs, options.quietMs - (currentTime - quietSinceMs)),
    );
  }
}
