import type { Page } from "playwright";
import type {
  MessageSnapshot,
  ResponseCapture,
  WhatsAppMessage,
} from "../types";
import {
  firstVisibleLocator,
  whatsappSelectors,
} from "./locators";

export interface CollectResponseOptions {
  baseline: MessageSnapshot;
  sentAt: Date;
  outgoingMessageId: string;
  timeoutMs: number;
  idleMs: number;
  pollIntervalMs?: number;
  context?: string;
  excludedIncomingTexts?: string[];
}

export interface ResponseCollectorEnvironment {
  readMessages(): Promise<WhatsAppMessage[]>;
  isRemoteTyping(): Promise<boolean>;
  now(): number;
  wait(milliseconds: number): Promise<void>;
  log?(message: string): void;
}

interface RawMessage {
  id: string;
  direction: "incoming" | "outgoing";
  text: string;
  domIndex: number;
  deliveryStatus?: "pending" | "sent" | "delivered" | "failed";
}

export async function isRemoteTyping(page: Page): Promise<boolean> {
  for (const selector of whatsappSelectors.typingIndicator) {
    const indicator = page.locator(selector).first();
    if (!(await indicator.isVisible().catch(() => false))) {
      continue;
    }
    const text = await indicator.innerText().catch(() => "");
    if (/typing|mengetik/i.test(text)) {
      return true;
    }
  }
  return false;
}

function normalizeMessage(value: string): string {
  return value.replace(/\s+/g, " ").trim().toLowerCase();
}

export function combineBotMessages(messages: WhatsAppMessage[]): string {
  if (messages.length === 1) {
    return messages[0].text;
  }
  return messages
    .map((message, index) => `Message ${index + 1}:\n${message.text}`)
    .join("\n\n");
}

export async function readMessages(page: Page): Promise<WhatsAppMessage[]> {
  const conversation = await firstVisibleLocator(
    page,
    whatsappSelectors.conversationRoot,
  );
  if (!conversation) {
    return [];
  }

  let messageLocator = conversation.locator(whatsappSelectors.message[0]);
  for (const selector of whatsappSelectors.message) {
    const candidate = conversation.locator(selector);
    if ((await candidate.count()) > 0) {
      messageLocator = candidate;
      break;
    }
  }

  const rawMessages = await messageLocator.evaluateAll(
    (elements): RawMessage[] => {
      return elements.flatMap((node, domIndex): RawMessage[] => {
        const element = node as HTMLElement;
        const hasOutgoingMarker =
          Boolean(element.querySelector('[data-testid="tail-out"]')) ||
          element.classList.contains("message-out") ||
          Boolean(element.querySelector(".message-out"));
        const hasIncomingMarker =
          Boolean(element.querySelector('[data-testid="tail-in"]')) ||
          element.classList.contains("message-in") ||
          Boolean(element.querySelector(".message-in"));
        if (!hasOutgoingMarker && !hasIncomingMarker) {
          return [];
        }
        const direction = hasOutgoingMarker ? "outgoing" : "incoming";
        const idElement =
          element.closest("[data-id]") ?? element.querySelector("[data-id]");
        const metadataElement = element.querySelector("[data-pre-plain-text]");
        const metadata = metadataElement?.getAttribute("data-pre-plain-text") ?? "";
        let deliveryStatus: RawMessage["deliveryStatus"];
        if (direction === "outgoing") {
          if (
            element.querySelector(
              '[data-icon="msg-error"], [data-testid="msg-error"]',
            )
          ) {
            deliveryStatus = "failed";
          } else if (
            element.querySelector(
              '[data-icon="msg-time"], [data-testid="msg-time"]',
            )
          ) {
            deliveryStatus = "pending";
          } else if (
            element.querySelector(
              '[data-icon="msg-dblcheck"], [data-testid="msg-dblcheck"]',
            )
          ) {
            deliveryStatus = "delivered";
          } else if (
            element.querySelector(
              '[data-icon="msg-check"], [data-testid="msg-check"]',
            )
          ) {
            deliveryStatus = "sent";
          }
        }
        const textElements = [
          ...element.querySelectorAll(
            '[data-pre-plain-text] [data-testid="selectable-text"], [data-pre-plain-text] .selectable-text, [data-testid="msg-text"]',
          ),
        ].filter(
          (candidate, index, candidates) =>
            !candidates.some(
              (other, otherIndex) =>
                otherIndex !== index && other.contains(candidate),
            ),
        ) as HTMLElement[];
        let text = (textElements.length > 0
          ? textElements
              .map((textElement) => textElement.innerText)
              .filter(Boolean)
              .join("\n")
          : (metadataElement as HTMLElement | null)?.innerText ?? ""
        )
          .replace(/\u200e/g, "")
          .trim();
        const messageContainer =
          element.matches('[data-testid="msg-container"]')
            ? element
            : element.querySelector('[data-testid="msg-container"]');
        const interactiveLabels = messageContainer
          ? [...messageContainer.querySelectorAll('[role="button"]')]
              .map((button) => (button as HTMLElement).innerText.trim())
              .filter(
                (label, index, labels) =>
                  Boolean(label) &&
                  labels.indexOf(label) === index &&
                  !text.includes(label),
              )
          : [];
        if (interactiveLabels.length > 0) {
          text = [text, ...interactiveLabels].filter(Boolean).join("\n");
        }
        if (!text && messageContainer) {
          const testIds = [...messageContainer.querySelectorAll("[data-testid]")]
            .map((candidate) => candidate.getAttribute("data-testid") ?? "")
            .join(" ");
          if (/sticker/i.test(testIds)) {
            text = "[Sticker]";
          } else if (/image|photo/i.test(testIds)) {
            text = "[Image]";
          } else if (/video/i.test(testIds)) {
            text = "[Video]";
          } else if (/audio|voice|ptt/i.test(testIds)) {
            text = "[Audio]";
          } else if (/document|file/i.test(testIds)) {
            text = "[Document]";
          } else if (/location|map/i.test(testIds)) {
            text = "[Location]";
          }
        }
        const rawId = idElement?.getAttribute("data-id");
        const fallbackIdSource = `${metadata}|${direction}|${domIndex}`;
        let fallbackHash = 2166136261;
        for (let index = 0; index < fallbackIdSource.length; index += 1) {
          fallbackHash ^= fallbackIdSource.charCodeAt(index);
          fallbackHash = Math.imul(fallbackHash, 16777619);
        }

        return [{
          id:
            rawId ||
            `fallback-${direction}-${(fallbackHash >>> 0).toString(16)}-${domIndex}`,
          direction,
          text,
          domIndex,
          deliveryStatus,
        }];
      });
    },
  );

  const observedAt = new Date();
  return rawMessages.map((message) => ({ ...message, observedAt }));
}

export async function snapshotMessages(page: Page): Promise<MessageSnapshot> {
  const messages = await readMessages(page);
  return {
    ids: new Set(messages.map(({ id }) => id)),
    messageCount: messages.length,
    messages,
  };
}

export async function collectBotResponse(
  page: Page,
  options: CollectResponseOptions,
): Promise<ResponseCapture> {
  return collectBotResponseWithEnvironment(options, {
    readMessages: () => readMessages(page),
    isRemoteTyping: () => isRemoteTyping(page),
    now: () => Date.now(),
    wait: (milliseconds) => page.waitForTimeout(milliseconds),
    log: (message) => console.log(message),
  });
}

export async function collectBotResponseWithEnvironment(
  options: CollectResponseOptions,
  environment: ResponseCollectorEnvironment,
): Promise<ResponseCapture> {
  const pollIntervalMs = options.pollIntervalMs ?? 250;
  const deadline = options.sentAt.getTime() + options.timeoutMs;
  const captured = new Map<string, WhatsAppMessage>();
  const responseActivityAt = new Map<string, number>();
  const ignored = new Map<string, string>();
  const excludedTexts = new Set(
    (options.excludedIncomingTexts ?? []).map(normalizeMessage),
  );
  const prefix = options.context ? `[Test ${options.context}] ` : "";
  const log = (message: string): void => {
    const formatted = `${prefix}${message}`;
    if (environment.log) {
      environment.log(formatted);
    } else {
      console.log(formatted);
    }
  };
  let firstResponseAtMs: number | undefined;
  let lastResponseAtMs: number | undefined;
  let quietSinceMs: number | undefined;
  let typingWasVisible = false;

  const refreshResponseTiming = (): void => {
    const messages = [...captured.values()];
    firstResponseAtMs = messages.length
      ? Math.min(...messages.map((message) => message.observedAt.getTime()))
      : undefined;
    lastResponseAtMs = responseActivityAt.size
      ? Math.max(...responseActivityAt.values())
      : undefined;
    quietSinceMs = lastResponseAtMs;
  };

  if (!Number.isInteger(options.idleMs) || options.idleMs < 1) {
    throw new Error("Response idle window must be a positive integer");
  }
  if (
    !Number.isInteger(options.timeoutMs) ||
    options.timeoutMs <= options.idleMs
  ) {
    throw new Error("Response timeout must be greater than the idle window");
  }
  if (!Number.isInteger(pollIntervalMs) || pollIntervalMs < 1) {
    throw new Error("Response poll interval must be a positive integer");
  }

  const result = (timedOut: boolean, completedAtMs: number): ResponseCapture => {
    const messages = [...captured.values()].sort(
      (left, right) => left.domIndex - right.domIndex,
    );
    const firstResponseAt =
      firstResponseAtMs === undefined
        ? undefined
        : new Date(firstResponseAtMs);
    return {
      messages,
      combinedResponse: combineBotMessages(messages),
      sentAt: options.sentAt,
      firstResponseAt,
      completedAt: new Date(completedAtMs),
      firstResponseMs:
        firstResponseAtMs === undefined
          ? undefined
          : firstResponseAtMs - options.sentAt.getTime(),
      totalResponseMs:
        lastResponseAtMs === undefined
          ? undefined
          : lastResponseAtMs - options.sentAt.getTime(),
      timedOut,
    };
  };

  while (environment.now() < deadline) {
    const messages = await environment.readMessages();
    if (environment.now() >= deadline) {
      break;
    }
    const outgoingAnchor = messages.find(
      (message) =>
        message.direction === "outgoing" &&
        message.id === options.outgoingMessageId,
    );

    const newIncoming = messages.filter(
      (message) =>
        Boolean(outgoingAnchor) &&
        message.direction === "incoming" &&
        message.text.length > 0 &&
        !options.baseline.ids.has(message.id) &&
        message.domIndex > outgoingAnchor!.domIndex,
    );

    const observedAtMs = environment.now();
    for (const message of newIncoming) {
      const normalized = normalizeMessage(message.text);
      if (excludedTexts.has(normalized)) {
        const removed = captured.delete(message.id);
        responseActivityAt.delete(message.id);
        if (removed) {
          refreshResponseTiming();
        }
        if (ignored.get(message.id) !== message.text) {
          ignored.set(message.id, message.text);
          if (firstResponseAtMs !== undefined) {
            quietSinceMs = observedAtMs;
          }
          log(`[Collector] Ignored control response: ${message.text}`);
        }
        continue;
      }
      if (ignored.has(message.id)) {
        if (ignored.get(message.id) !== message.text) {
          ignored.set(message.id, message.text);
          if (firstResponseAtMs !== undefined) {
            quietSinceMs = observedAtMs;
          }
          log(`[Collector] Ignored control response update: ${message.text}`);
        }
        continue;
      }

      const existing = captured.get(message.id);
      if (!existing) {
        const observedAt = new Date(observedAtMs);
        const capturedMessage = { ...message, observedAt };
        captured.set(message.id, capturedMessage);
        responseActivityAt.set(message.id, observedAtMs);
        firstResponseAtMs ??= observedAtMs;
        lastResponseAtMs = observedAtMs;
        quietSinceMs = observedAtMs;
        log(
          `[Bot] Message ${captured.size} received at +${observedAtMs - options.sentAt.getTime()} ms`,
        );
        log(`[Collector] Idle timer reset: ${options.idleMs} ms`);
      } else if (existing.text !== message.text) {
        captured.set(message.id, { ...message, observedAt: existing.observedAt });
        responseActivityAt.set(message.id, observedAtMs);
        lastResponseAtMs = Math.max(...responseActivityAt.values());
        quietSinceMs = observedAtMs;
        const sequence = [...captured.keys()].indexOf(message.id) + 1;
        log(
          `[Bot] Message ${sequence} updated at +${observedAtMs - options.sentAt.getTime()} ms`,
        );
        log(`[Collector] Idle timer reset: ${options.idleMs} ms`);
      }
    }

    if (firstResponseAtMs !== undefined) {
      const typing = await environment.isRemoteTyping();
      const currentTimeMs = environment.now();
      if (currentTimeMs >= deadline) {
        break;
      }
      if (typing) {
        quietSinceMs = currentTimeMs;
        if (!typingWasVisible) {
          log("[Collector] Bot is typing; quiet timer held");
        }
      } else if (typingWasVisible) {
        quietSinceMs = currentTimeMs;
        log(`[Collector] Typing stopped; idle timer reset: ${options.idleMs} ms`);
      }
      typingWasVisible = typing;

      if (
        !typing &&
        quietSinceMs !== undefined &&
        currentTimeMs - quietSinceMs >= options.idleMs
      ) {
        log(`[Collector] No incoming messages for ${options.idleMs} ms`);
        log("[Collector] Response complete");
        log(`[Collector] ${captured.size} bot messages captured`);
        log(
          `[Collector] First response: ${firstResponseAtMs - options.sentAt.getTime()} ms`,
        );
        log(
          `[Collector] Total response: ${lastResponseAtMs! - options.sentAt.getTime()} ms`,
        );
        return result(false, currentTimeMs);
      }
    }

    const remainingMs = deadline - environment.now();
    if (remainingMs > 0) {
      await environment.wait(Math.min(pollIntervalMs, remainingMs));
    }
  }

  const completedAtMs = Math.min(environment.now(), deadline);
  log(`[Collector] Hard timeout after ${options.timeoutMs} ms`);
  log(`[Collector] ${captured.size} bot messages captured before timeout`);
  return result(true, completedAtMs);
}
