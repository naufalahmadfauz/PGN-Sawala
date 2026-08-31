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

interface CollectResponseOptions {
  baseline: MessageSnapshot;
  sentAt: Date;
  outgoingMessageId: string;
  timeoutMs: number;
  idleMs: number;
  pollIntervalMs?: number;
}

interface RawMessage {
  id: string;
  direction: "incoming" | "outgoing";
  text: string;
  domIndex: number;
  deliveryStatus?: "pending" | "sent" | "delivered" | "failed";
}

async function isRemoteTyping(page: Page): Promise<boolean> {
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
        const fallbackIdSource = `${metadata}|${text}`;
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
  const pollIntervalMs = options.pollIntervalMs ?? 250;
  const deadline = options.sentAt.getTime() + options.timeoutMs;
  const captured = new Map<string, WhatsAppMessage>();
  let firstResponseAt: Date | undefined;
  let lastIncomingAtMs: number | undefined;

  while (Date.now() < deadline) {
    const messages = await readMessages(page);
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

    for (const message of newIncoming) {
      const existing = captured.get(message.id);
      if (!existing) {
        const observedAt = new Date();
        const capturedMessage = { ...message, observedAt };
        captured.set(message.id, capturedMessage);
        firstResponseAt ??= observedAt;
        lastIncomingAtMs = observedAt.getTime();
      } else if (existing.text !== message.text) {
        captured.set(message.id, { ...message, observedAt: existing.observedAt });
        lastIncomingAtMs = Date.now();
      }
    }

    if (
      firstResponseAt &&
      lastIncomingAtMs !== undefined &&
      Date.now() - lastIncomingAtMs >= options.idleMs &&
      !(await isRemoteTyping(page))
    ) {
      const completedAt = new Date();
      const responseMessages = [...captured.values()].sort(
        (left, right) => left.domIndex - right.domIndex,
      );
      return {
        messages: responseMessages,
        combinedResponse: responseMessages.map(({ text }) => text).join("\n"),
        sentAt: options.sentAt,
        firstResponseAt,
        completedAt,
        firstResponseMs:
          firstResponseAt.getTime() - options.sentAt.getTime(),
        totalResponseMs: completedAt.getTime() - options.sentAt.getTime(),
        timedOut: false,
      };
    }

    await page.waitForTimeout(
      Math.min(pollIntervalMs, Math.max(0, deadline - Date.now())),
    );
  }

  const completedAt = new Date();
  const responseMessages = [...captured.values()].sort(
    (left, right) => left.domIndex - right.domIndex,
  );
  return {
    messages: responseMessages,
    combinedResponse: responseMessages.map(({ text }) => text).join("\n"),
    sentAt: options.sentAt,
    firstResponseAt,
    completedAt,
    firstResponseMs: firstResponseAt
      ? firstResponseAt.getTime() - options.sentAt.getTime()
      : undefined,
    totalResponseMs: completedAt.getTime() - options.sentAt.getTime(),
    timedOut: true,
  };
}
