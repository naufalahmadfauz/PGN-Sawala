import type { Locator, Page } from "playwright";

export type LocatorRoot = Page | Locator;

export const whatsappSelectors = {
  qrCanvas: [
    'canvas[aria-label*="QR code" i]',
    'canvas[aria-label*="QR" i]',
  ],
  authenticatedShell: [
    '#app [aria-label="Chat list"]',
    '#app [data-testid="chat-list"]',
    "#pane-side",
    '#app div[contenteditable="true"][role="textbox"][aria-label*="search" i]',
  ],
  sidePanel: ["#pane-side", '#app [aria-label="Chat list"]'],
  searchBox: [
    '#side div[contenteditable="true"][role="textbox"]',
    'div[contenteditable="true"][role="textbox"][aria-label*="search" i]',
    'div[contenteditable="true"][data-tab="3"]',
  ],
  conversationRoot: [
    "#main",
    '[data-testid="conversation-panel-wrapper"]',
  ],
  conversationTitle: [
    '#main [data-testid="conversation-info-header-chat-title"]',
    '#main header span[title]',
  ],
  typingIndicator: ['#main [data-testid="chat-subtitle"]'],
  composer: [
    '#main footer div[contenteditable="true"][role="textbox"]',
    '#main [data-testid="conversation-compose-box-input"]',
    '#main footer div[contenteditable="true"][data-tab]',
  ],
  message: [
    '[role="row"]:has([data-testid="msg-container"])',
    '[data-testid^="conv-msg-"]:has([data-testid="msg-container"])',
    ".message-in, .message-out",
  ],
} as const;

export async function firstVisibleLocator(
  root: LocatorRoot,
  selectors: readonly string[],
): Promise<Locator | undefined> {
  for (const selector of selectors) {
    const locator = root.locator(selector).first();
    if (await locator.isVisible().catch(() => false)) {
      return locator;
    }
  }
  return undefined;
}

export async function waitForFirstVisibleLocator(
  root: LocatorRoot,
  selectors: readonly string[],
  timeoutMs: number,
  pollIntervalMs = 250,
): Promise<Locator> {
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    const locator = await firstVisibleLocator(root, selectors);
    if (locator) {
      return locator;
    }
    await new Promise((resolve) => setTimeout(resolve, pollIntervalMs));
  }

  throw new Error(
    `No visible element matched the expected WhatsApp selectors within ${timeoutMs} ms`,
  );
}
