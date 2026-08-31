import { mkdir, writeFile } from "node:fs/promises";
import path from "node:path";
import { chromium, type BrowserContext, type Locator, type Page } from "playwright";
import type { AppConfig, WhatsAppTarget } from "../config";
import type {
  MessageSnapshot,
  ResponseCapture,
  SentMessage,
  WhatsAppMessage,
} from "../types";
import {
  collectBotResponse,
  readMessages,
  snapshotMessages,
} from "./response-collector";
import {
  firstVisibleLocator,
  waitForFirstVisibleLocator,
  whatsappSelectors,
} from "./locators";

interface AuthenticationOptions {
  allowQrLogin: boolean;
  timeoutMs?: number;
}

interface QrInspection {
  ready: boolean;
  fingerprint: string;
  darkRatio: number;
  transitions: number;
}

export class AuthenticationRequiredError extends Error {
  constructor(message = "WhatsApp authentication is required") {
    super(message);
    this.name = "AuthenticationRequiredError";
  }
}

export class WhatsAppClient {
  private context?: BrowserContext;
  private page?: Page;

  constructor(private readonly config: AppConfig) {}

  async open(): Promise<void> {
    await Promise.all([
      mkdir(this.config.profileDir, { recursive: true }),
      mkdir(this.config.artifactsDir, { recursive: true }),
      mkdir(this.config.debugDir, { recursive: true }),
      mkdir(this.config.evidenceDir, { recursive: true }),
    ]);

    console.log("[WhatsApp] Starting Chromium");
    const launchOptions: NonNullable<
      Parameters<typeof chromium.launchPersistentContext>[1]
    > = {
      headless: this.config.headless,
      viewport: { width: 1440, height: 1000 },
      locale: "en-US",
      args: ["--disable-dev-shm-usage"],
    };
    if (this.config.browserChannel) {
      launchOptions.channel = this.config.browserChannel;
    }

    try {
      this.context = await chromium.launchPersistentContext(
        this.config.profileDir,
        launchOptions,
      );
    } catch (error) {
      const detail = error instanceof Error ? error.message : String(error);
      if (/ProcessSingleton|profile.*in use|SingletonLock/i.test(detail)) {
        throw new Error(
          "The .whatsapp-profile browser is already running. Stop the other WhatsApp harness process and retry.",
          { cause: error },
        );
      }
      throw error;
    }

    const pages = this.context.pages();
    this.page =
      pages.find((candidate) =>
        candidate.url().startsWith(this.config.whatsappUrl),
      ) ??
      pages[0] ??
      (await this.context.newPage());
    this.page.setDefaultTimeout(10_000);

    await this.page.goto(this.config.whatsappUrl, {
      waitUntil: "domcontentloaded",
      timeout: 60_000,
    });
  }

  async ensureAuthenticated(options: AuthenticationOptions): Promise<void> {
    const page = this.requirePage();
    const timeoutMs =
      options.timeoutMs ??
      (options.allowQrLogin
        ? this.config.loginTimeoutMs
        : this.config.authTimeoutMs);
    const deadline = Date.now() + timeoutMs;
    let announcedQr = false;
    let lastQrFingerprint: string | undefined;
    let lastReloadAttempt = 0;

    while (Date.now() < deadline) {
      if (await this.isAuthenticated()) {
        console.log("[WhatsApp] Session authenticated");
        return;
      }

      if (await this.hasUnsupportedBrowserMessage()) {
        await this.saveDebugArtifacts("unsupported-browser");
        throw new Error(
          "WhatsApp rejected this Chromium build as unsupported. See artifacts/debug/unsupported-browser.png.",
        );
      }

      if (
        options.allowQrLogin &&
        Date.now() - lastReloadAttempt >= 5_000
      ) {
        const reloadCandidates = [
          page.getByRole("button", { name: /reload.*QR/i }).first(),
          page.getByText(/(?:click|tap).*reload.*QR code/i).first(),
        ];
        const reloadQr = await this.firstVisibleCandidate(reloadCandidates);
        if (reloadQr) {
          try {
            await reloadQr.click();
            lastReloadAttempt = Date.now();
            lastQrFingerprint = undefined;
            console.log("[WhatsApp] Requested a fresh QR code");
            await page.waitForTimeout(500);
            continue;
          } catch {
            lastReloadAttempt = Date.now();
          }
        }
      }

      const qr = await this.inspectQrCanvas();
      if (qr?.ready) {
        if (!options.allowQrLogin) {
          await page.screenshot({
            path: this.config.loginScreenshotPath,
            fullPage: true,
          });
          throw new AuthenticationRequiredError(
            "The saved WhatsApp session is not authenticated. Run npm run whatsapp:login.",
          );
        }

        if (qr.fingerprint !== lastQrFingerprint) {
          await page.screenshot({
            path: this.config.loginScreenshotPath,
            fullPage: true,
          });
          lastQrFingerprint = qr.fingerprint;

          if (!announcedQr) {
            console.log("");
            console.log("==================================================");
            console.log("ACTION REQUIRED");
            console.log("Open artifacts/whatsapp-login.png");
            console.log("Scan this QR using WhatsApp -> Linked Devices");
            console.log("==================================================");
            console.log("");
            console.log("[WhatsApp] Browser will remain open while login completes");
            announcedQr = true;
          } else {
            console.log(
              "[WhatsApp] QR refreshed; artifacts/whatsapp-login.png has been updated",
            );
          }
        }
      }

      await page.waitForTimeout(500);
    }

    await this.saveDebugArtifacts("login-failure");
    throw new Error(
      `WhatsApp authentication was not detected within ${timeoutMs} ms. See artifacts/debug/login-failure.png.`,
    );
  }

  async openChat(target: WhatsAppTarget): Promise<void> {
    const page = this.requirePage();
    console.log("[WhatsApp] Opening PGN chat");

    try {
      if (target.kind === "phone") {
        const targetUrl = new URL("/send", this.config.whatsappUrl);
        targetUrl.searchParams.set("phone", target.value);
        await page.goto(targetUrl.href, {
          waitUntil: "domcontentloaded",
          timeout: 60_000,
        });
      } else {
        await this.openChatByName(target.value);
      }

      await this.waitForConversationReady(target, 45_000);
      console.log("[WhatsApp] PGN chat ready");
    } catch (error) {
      await this.saveDebugArtifacts("chat-not-found");
      throw error;
    }
  }

  async getMessages() {
    return readMessages(this.requirePage());
  }

  async captureMessageState(): Promise<MessageSnapshot> {
    await this.waitForMessageHistorySettled(30_000);
    return snapshotMessages(this.requirePage());
  }

  async sendMessage(
    message: string,
    baseline: MessageSnapshot,
  ): Promise<SentMessage> {
    if (!message.trim()) {
      throw new Error("Cannot send an empty WhatsApp message");
    }

    const composer = await waitForFirstVisibleLocator(
      this.requirePage(),
      whatsappSelectors.composer,
      15_000,
    );
    await composer.click();
    await composer.fill(message);
    const sentAt = new Date();
    await composer.press("Enter");

    const deadline = Date.now() + 15_000;
    let pendingMessage: WhatsAppMessage | undefined;
    while (Date.now() < deadline) {
      const messages = await readMessages(this.requirePage());
      const newOutgoing = messages.filter(
        (candidate) =>
          candidate.direction === "outgoing" &&
          !baseline.ids.has(candidate.id),
      );
      const normalizedInput = this.normalizeMessageText(message);
      const exactMatch = newOutgoing
        .filter(
          (candidate) =>
            this.normalizeMessageText(candidate.text) === normalizedInput,
        )
        .at(-1);
      const sentMessage =
        exactMatch ?? (newOutgoing.length === 1 ? newOutgoing[0] : undefined);

      if (sentMessage) {
        if (sentMessage.deliveryStatus === "failed") {
          throw new Error("WhatsApp reported that the outgoing message failed");
        }
        if (sentMessage.deliveryStatus !== "pending") {
          console.log("[WhatsApp] Outgoing message confirmed");
          return {
            sentAt,
            messageId: sentMessage.id,
            renderedText: sentMessage.text,
          };
        }
        pendingMessage = sentMessage;
      }

      await this.requirePage().waitForTimeout(200);
    }

    if (pendingMessage) {
      throw new Error(
        "The outgoing WhatsApp message remained pending for 15 seconds",
      );
    }
    throw new Error(
      "WhatsApp did not render a new outgoing message after Enter was pressed",
    );
  }

  async waitForBotResponse(
    baseline: MessageSnapshot,
    sentMessage: SentMessage,
  ): Promise<ResponseCapture> {
    return collectBotResponse(this.requirePage(), {
      baseline,
      sentAt: sentMessage.sentAt,
      outgoingMessageId: sentMessage.messageId,
      timeoutMs: this.config.responseTimeoutMs,
      idleMs: this.config.responseIdleMs,
    });
  }

  async captureScreenshot(absolutePath: string): Promise<void> {
    await mkdir(path.dirname(absolutePath), { recursive: true });
    await this.requirePage().screenshot({ path: absolutePath, fullPage: true });
  }

  async saveDebugArtifacts(
    name: string,
  ): Promise<{ screenshotPath: string; diagnosticsPath: string }> {
    const page = this.requirePage();
    const safeName = name.replace(/[^a-zA-Z0-9_-]+/g, "-");
    const screenshotPath = path.join(this.config.debugDir, `${safeName}.png`);
    const diagnosticsPath = path.join(this.config.debugDir, `${safeName}.json`);
    await mkdir(this.config.debugDir, { recursive: true });

    await page
      .screenshot({ path: screenshotPath, fullPage: true })
      .catch(() => undefined);
    const diagnostics = await page
      .evaluate(() => ({
        capturedAt: new Date().toISOString(),
        title: document.title,
        url: location.href,
        visibleText: document.body?.innerText.slice(0, 8_000) ?? "",
        canvases: [...document.querySelectorAll("canvas")].map((canvas) => ({
          ariaLabel: canvas.getAttribute("aria-label"),
          height: canvas.height,
          width: canvas.width,
        })),
      }))
      .catch((error) => ({
        capturedAt: new Date().toISOString(),
        diagnosticsError: error instanceof Error ? error.message : String(error),
      }));
    await writeFile(diagnosticsPath, JSON.stringify(diagnostics, null, 2));

    return { screenshotPath, diagnosticsPath };
  }

  async close(): Promise<void> {
    const context = this.context;
    this.context = undefined;
    this.page = undefined;
    if (context) {
      await context.close().catch(() => undefined);
    }
  }

  private requirePage(): Page {
    if (!this.page) {
      throw new Error("WhatsAppClient.open() must be called first");
    }
    return this.page;
  }

  private async isAuthenticated(): Promise<boolean> {
    return Boolean(
      await firstVisibleLocator(
        this.requirePage(),
        whatsappSelectors.authenticatedShell,
      ),
    );
  }

  private async hasUnsupportedBrowserMessage(): Promise<boolean> {
    const message = this.requirePage()
      .getByText(/WhatsApp works with (?:Google )?Chrome/i)
      .first();
    return message.isVisible().catch(() => false);
  }

  private async inspectQrCanvas(): Promise<QrInspection | undefined> {
    const canvasLocator = await firstVisibleLocator(
      this.requirePage(),
      whatsappSelectors.qrCanvas,
    );
    if (!canvasLocator) {
      return undefined;
    }

    return canvasLocator.evaluate((node) => {
      const canvas = node as HTMLCanvasElement;
      if (canvas.width < 160 || canvas.height < 160) {
        return {
          ready: false,
          fingerprint: "",
          darkRatio: 0,
          transitions: 0,
        };
      }

      const context = canvas.getContext("2d", { willReadFrequently: true });
      if (!context) {
        return {
          ready: false,
          fingerprint: "",
          darkRatio: 0,
          transitions: 0,
        };
      }

      const pixels = context.getImageData(0, 0, canvas.width, canvas.height).data;
      const step = 2;
      let darkPixels = 0;
      let samples = 0;
      let transitions = 0;
      let fingerprint = 2166136261;

      for (let y = 0; y < canvas.height; y += step) {
        let previousDark: boolean | undefined;
        for (let x = 0; x < canvas.width; x += step) {
          const offset = (y * canvas.width + x) * 4;
          const alpha = pixels[offset + 3];
          const luminance =
            alpha < 20
              ? 255
              : pixels[offset] * 0.2126 +
                pixels[offset + 1] * 0.7152 +
                pixels[offset + 2] * 0.0722;
          const dark = luminance < 128;
          darkPixels += dark ? 1 : 0;
          samples += 1;
          if (previousDark !== undefined && previousDark !== dark) {
            transitions += 1;
          }
          previousDark = dark;
          fingerprint ^= dark ? 1 : 0;
          fingerprint = Math.imul(fingerprint, 16777619);
        }
      }

      const darkRatio = darkPixels / samples;
      return {
        ready:
          darkRatio >= 0.08 && darkRatio <= 0.85 && transitions >= 300,
        fingerprint: (fingerprint >>> 0).toString(16),
        darkRatio,
        transitions,
      };
    });
  }

  private async openChatByName(chatName: string): Promise<void> {
    const page = this.requirePage();
    const searchBox = await waitForFirstVisibleLocator(
      page,
      whatsappSelectors.searchBox,
      30_000,
    );
    await searchBox.click();
    await searchBox.fill("");
    await searchBox.fill(chatName);

    const sidePanel =
      (await firstVisibleLocator(page, whatsappSelectors.sidePanel)) ?? page;
    const titleMatch = sidePanel.getByTitle(chatName, { exact: true }).first();
    const result = await this.waitForVisibleCandidate([titleMatch], 30_000);
    if (!result) {
      throw new Error(`WhatsApp chat named "${chatName}" was not found`);
    }
    await result.click();
  }

  private async waitForConversationReady(
    target: WhatsAppTarget,
    timeoutMs: number,
  ): Promise<void> {
    const page = this.requirePage();
    const deadline = Date.now() + timeoutMs;
    const invalidPhoneMessage = page
      .getByText(
        /phone number shared via url is invalid|phone number isn't on WhatsApp|couldn't find.*WhatsApp/i,
      )
      .first();

    while (Date.now() < deadline) {
      if (await invalidPhoneMessage.isVisible().catch(() => false)) {
        throw new Error("The configured PGN phone number is not available on WhatsApp");
      }
      if (await firstVisibleLocator(page, whatsappSelectors.composer)) {
        await this.waitForMessageHistorySettled(30_000);
        await this.verifyOpenConversation(target);
        return;
      }
      await page.waitForTimeout(250);
    }

    throw new Error(`PGN conversation did not become ready within ${timeoutMs} ms`);
  }

  private async waitForMessageHistorySettled(timeoutMs: number): Promise<void> {
    const page = this.requirePage();
    const conversation = await firstVisibleLocator(
      page,
      whatsappSelectors.conversationRoot,
    );
    if (!conversation) {
      return;
    }

    await page.waitForTimeout(1_500);
    const deadline = Date.now() + timeoutMs;
    let previousCount = -1;
    let stableSince: number | undefined;
    while (Date.now() < deadline) {
      const currentCount = (await readMessages(page)).length;

      if (currentCount === previousCount) {
        stableSince ??= Date.now();
        if (Date.now() - stableSince >= 1_000) {
          return;
        }
      } else {
        stableSince = undefined;
      }
      previousCount = currentCount;
      await page.waitForTimeout(250);
    }

    throw new Error(
      `WhatsApp message history did not finish loading within ${timeoutMs} ms`,
    );
  }

  private async waitForVisibleCandidate(
    candidates: Locator[],
    timeoutMs: number,
  ): Promise<Locator | undefined> {
    const deadline = Date.now() + timeoutMs;
    while (Date.now() < deadline) {
      for (const candidate of candidates) {
        if (await candidate.isVisible().catch(() => false)) {
          return candidate;
        }
      }
      await this.requirePage().waitForTimeout(250);
    }
    return undefined;
  }

  private async firstVisibleCandidate(
    candidates: Locator[],
  ): Promise<Locator | undefined> {
    for (const candidate of candidates) {
      if (await candidate.isVisible().catch(() => false)) {
        return candidate;
      }
    }
    return undefined;
  }

  private async verifyOpenConversation(target: WhatsAppTarget): Promise<void> {
    const title = await waitForFirstVisibleLocator(
      this.requirePage(),
      whatsappSelectors.conversationTitle,
      15_000,
    );
    const visibleTitle = (
      (await title.innerText().catch(() => "")) ||
      (await title.getAttribute("title")) ||
      ""
    ).trim();

    if (target.kind === "phone") {
      const visiblePhone = visibleTitle.replace(/\D/g, "");
      if (visiblePhone !== target.value) {
        throw new Error(
          `Refusing to send: the open WhatsApp header does not match the configured PGN phone number`,
        );
      }
    } else if (visibleTitle.localeCompare(target.value, undefined, {
      sensitivity: "accent",
    }) !== 0) {
      throw new Error(
        `Refusing to send: the open WhatsApp header is not "${target.value}"`,
      );
    }

    console.log("[WhatsApp] PGN recipient verified");
  }

  private normalizeMessageText(value: string): string {
    return value.replace(/\u200e/g, "").replace(/\s+/g, " ").trim();
  }
}
