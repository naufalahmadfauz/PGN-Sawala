import path from "node:path";

const SAFE_ALLOWED_MENTIONS = {
  parse: [] as string[],
  users: [] as string[],
  roles: [] as string[],
  replied_user: false,
};
const DEFAULT_TIMEOUT_MS = 3_000;
const DEFAULT_MAX_ATTEMPTS = 3;
const DEFAULT_MAX_RETRY_DELAY_MS = 2_000;
const DEFAULT_MIN_PROGRESS_INTERVAL_MS = 10_000;
const INTERRUPTION_REQUEST_TIMEOUT_MS = 700;

export type DiscordRunMode = "full" | "retest";

export interface DiscordRunStartedEvent {
  runId: string;
  mode: DiscordRunMode;
  selectedScenarios: number;
  startedAt: Date;
  googleDriveEvidenceEnabled: boolean;
  workbookPath: string;
}

export interface DiscordRunProgressEvent {
  completedScenarios: number;
  totalScenarios: number;
  currentScenarioId?: string;
  capturedScenarios: number;
  timeouts: number;
  technicalErrors: number;
  evidenceUploaded: number;
  evidenceUploadErrors: number;
  updatedAt: Date;
}

export interface DiscordRunCompletedEvent extends DiscordRunProgressEvent {
  completedAt: Date;
  checkpoint?: boolean;
}

export interface DiscordRunFailedEvent extends DiscordRunProgressEvent {
  failedAt: Date;
  reason: string;
  workbookProgress: string;
  evidenceProgress: string;
}

export interface DiscordNotificationSettings {
  discordNotificationsEnabled: boolean;
  discordWebhookUrl?: string;
  discordProgressEvery: number;
  discordProgressMinutes: number;
  discordNotifyStart: boolean;
  discordNotifyProgress: boolean;
  discordNotifyComplete: boolean;
  discordNotifyFailure: boolean;
  discordConfigurationIssues?: readonly string[];
}

export interface DiscordNotifier {
  runStarted(event: DiscordRunStartedEvent): Promise<void>;
  runProgress(event: DiscordRunProgressEvent): Promise<void>;
  runCompleted(event: DiscordRunCompletedEvent): Promise<void>;
  runFailed(event: DiscordRunFailedEvent): Promise<void>;
  runInterrupted(
    signal: "SIGINT" | "SIGTERM",
    event: DiscordRunProgressEvent,
  ): Promise<void>;
}

export interface DiscordNotifierDependencies {
  fetch?: typeof fetch;
  now?: () => Date;
  sleep?: (milliseconds: number) => Promise<void>;
  warn?: (message: string) => void;
  environment?: NodeJS.ProcessEnv;
  platform?: NodeJS.Platform;
  timeoutMs?: number;
  maxAttempts?: number;
  maxRetryDelayMs?: number;
  minimumProgressIntervalMs?: number;
}

export interface DiscordValidationResult {
  enabled: boolean;
  configured: boolean;
  valid: boolean;
  connectivity: "ok" | "failed" | "not-tested";
  testNotificationSent: boolean;
  testNotificationDeliveryUncertain?: boolean;
  reason?: string;
}

export interface DiscordValidationOptions extends DiscordNotifierDependencies {
  sendTest?: boolean;
}

interface DiscordEmbedField {
  name: string;
  value: string;
  inline?: boolean;
}

interface DiscordEmbed {
  title: string;
  color: number;
  fields: DiscordEmbedField[];
  timestamp: string;
}

interface DiscordPayload {
  content?: string;
  embeds?: DiscordEmbed[];
  allowed_mentions: typeof SAFE_ALLOWED_MENTIONS;
}

class DiscordRequestError extends Error {
  constructor(
    message: string,
    readonly retryable: boolean,
    readonly retryAfterMs?: number,
    readonly deliveryUncertain = false,
  ) {
    super(message);
    this.name = "DiscordRequestError";
  }
}

function defaultSleep(milliseconds: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, milliseconds));
}

export function validateDiscordWebhookUrl(value: string | undefined):
  | { valid: true; url: string }
  | { valid: false; reason: string } {
  const trimmed = value?.trim();
  if (!trimmed) {
    return { valid: false, reason: "Discord webhook URL is not configured" };
  }
  let parsed: URL;
  try {
    parsed = new URL(trimmed);
  } catch {
    return { valid: false, reason: "Discord webhook URL is invalid" };
  }
  const allowedHosts = new Set([
    "discord.com",
    "canary.discord.com",
    "ptb.discord.com",
    "discordapp.com",
  ]);
  const webhookPath = /^\/api(?:\/v\d+)?\/webhooks\/\d+\/[^/]+\/?$/;
  if (
    parsed.protocol !== "https:" ||
    !allowedHosts.has(parsed.hostname.toLowerCase()) ||
    Boolean(parsed.username || parsed.password || parsed.port) ||
    !webhookPath.test(parsed.pathname)
  ) {
    return {
      valid: false,
      reason: "Use a valid Discord Incoming Webhook URL",
    };
  }
  parsed.hash = "";
  return { valid: true, url: parsed.toString() };
}

export function safeDiscordError(
  error: unknown,
  configuredWebhookUrl?: string,
): string {
  if (error instanceof DOMException && error.name === "AbortError") {
    return "request timed out";
  }
  const message =
    error instanceof Error ? error.message : "Unknown Discord error";
  return redactSensitiveText(message, configuredWebhookUrl);
}

function redactSensitiveText(
  value: string,
  configuredWebhookUrl?: string,
): string {
  let redacted = value.split(/\r?\n/, 1)[0]?.trim() || "Discord request failed";
  if (configuredWebhookUrl) {
    redacted = redacted
      .split(configuredWebhookUrl)
      .join("[REDACTED]");
  }
  redacted = redacted
    .replace(
      /https?:\/\/(?:canary\.|ptb\.)?(?:discord(?:app)?\.com)\/api(?:\/v\d+)?\/webhooks\/[^/\s]+\/[^\s"'`]+/giu,
      "[REDACTED]",
    )
    .replace(/https?:\/\/[^\s]+/giu, "[REDACTED URL]")
    .replace(/\+?\d(?:[\s().-]*\d){7,}/gu, "[REDACTED PHONE NUMBER]")
    .replace(
      /\bauthorization\s*[:=]\s*(?:bearer\s+)?[^\s,;]+/giu,
      "[REDACTED CREDENTIAL]",
    )
    .replace(
      /\b(?:password|passwd|secret|token|api[_-]?key)\s*[:=]\s*[^\s,;]+/giu,
      "[REDACTED CREDENTIAL]",
    );
  return redacted.length <= 300 ? redacted : `${redacted.slice(0, 297)}...`;
}

function safeOperationalReason(reason: string): string {
  return redactSensitiveText(reason || "Technical error");
}

function retryAfterMilliseconds(response: Response, body: unknown): number | undefined {
  const headerValue = response.headers.get("retry-after");
  if (headerValue !== null && headerValue.trim() !== "") {
    const header = Number(headerValue);
    if (Number.isFinite(header) && header >= 0) return header * 1_000;
  }
  if (body && typeof body === "object" && "retry_after" in body) {
    const seconds = Number((body as { retry_after?: unknown }).retry_after);
    if (Number.isFinite(seconds) && seconds >= 0) return seconds * 1_000;
  }
  return undefined;
}

class DiscordWebhookTransport {
  private readonly fetchImplementation: typeof fetch;
  private readonly sleep: (milliseconds: number) => Promise<void>;
  private readonly timeoutMs: number;
  private readonly maxAttempts: number;
  private readonly maxRetryDelayMs: number;

  constructor(
    private readonly webhookUrl: string,
    dependencies: DiscordNotifierDependencies,
  ) {
    this.fetchImplementation = dependencies.fetch ?? fetch;
    this.sleep = dependencies.sleep ?? defaultSleep;
    this.timeoutMs = dependencies.timeoutMs ?? DEFAULT_TIMEOUT_MS;
    this.maxAttempts = dependencies.maxAttempts ?? DEFAULT_MAX_ATTEMPTS;
    this.maxRetryDelayMs =
      dependencies.maxRetryDelayMs ?? DEFAULT_MAX_RETRY_DELAY_MS;
  }

  async inspect(
    options: { maxAttempts?: number; timeoutMs?: number } = {},
  ): Promise<void> {
    await this.request(
      "GET",
      new URL(this.webhookUrl),
      undefined,
      options,
      async (response) => {
        let webhook: unknown;
        try {
          webhook = await response.json();
        } catch {
          throw new DiscordRequestError(
            "Discord returned an invalid webhook response",
            false,
          );
        }
        const id =
          webhook && typeof webhook === "object" && "id" in webhook
            ? (webhook as { id?: unknown }).id
            : undefined;
        const type =
          webhook && typeof webhook === "object" && "type" in webhook
            ? (webhook as { type?: unknown }).type
            : undefined;
        if (typeof id !== "string" || !/^\d+$/.test(id) || type !== 1) {
          throw new DiscordRequestError(
            "Discord webhook is not a valid Incoming Webhook",
            false,
          );
        }
      },
    );
  }

  async execute(
    payload: DiscordPayload,
    options: { maxAttempts?: number; timeoutMs?: number } = {},
  ): Promise<string> {
    const url = new URL(this.webhookUrl);
    url.searchParams.set("wait", "true");
    return this.request(
      "POST",
      url,
      payload,
      options,
      async (response) => {
        let message: unknown;
        try {
          message = await response.json();
        } catch {
          throw new DiscordRequestError(
            "Discord returned an invalid message response",
            false,
            undefined,
            true,
          );
        }
        const id =
          message && typeof message === "object" && "id" in message
            ? (message as { id?: unknown }).id
            : undefined;
        if (typeof id !== "string" || !/^\d+$/.test(id)) {
          throw new DiscordRequestError(
            "Discord returned an invalid message response",
            false,
            undefined,
            true,
          );
        }
        return id;
      },
    );
  }

  async edit(
    messageId: string,
    payload: DiscordPayload,
    options: { maxAttempts?: number; timeoutMs?: number } = {},
  ): Promise<void> {
    const url = new URL(this.webhookUrl);
    url.pathname = `${url.pathname.replace(/\/$/, "")}/messages/${encodeURIComponent(messageId)}`;
    await this.request("PATCH", url, payload, options);
  }

  private async request<Value = void>(
    method: "GET" | "POST" | "PATCH",
    url: URL,
    payload?: DiscordPayload,
    options: { maxAttempts?: number; timeoutMs?: number } = {},
    readResponse?: (response: Response) => Promise<Value>,
  ): Promise<Value> {
    const attempts = Math.max(1, options.maxAttempts ?? this.maxAttempts);
    const timeoutMs = Math.max(1, options.timeoutMs ?? this.timeoutMs);
    let lastError: unknown;
    for (let attempt = 1; attempt <= attempts; attempt += 1) {
      const controller = new AbortController();
      let timeout: NodeJS.Timeout | undefined;
      try {
        const request = async (): Promise<Value> => {
          const response = await this.fetchImplementation(url, {
            method,
            headers: payload
              ? { "content-type": "application/json" }
              : undefined,
            body: payload ? JSON.stringify(payload) : undefined,
            signal: controller.signal,
          });
          if (response.ok) {
            return readResponse
              ? readResponse(response)
              : (undefined as Value);
          }

          let rateLimitBody: unknown;
          if (response.status === 429) {
            try {
              rateLimitBody = await response.json();
            } catch {
              rateLimitBody = undefined;
            }
          }
          const transient =
            response.status === 408 ||
            response.status === 429 ||
            response.status >= 500;
          const safeToRepeat = method !== "POST" || response.status === 429;
          throw new DiscordRequestError(
            response.status === 429
              ? "Discord rate limit was reached"
              : `Discord returned HTTP ${response.status}`,
            transient && safeToRepeat,
            retryAfterMilliseconds(response, rateLimitBody),
            method === "POST" &&
              (response.status === 408 || response.status >= 500),
          );
        };
        const deadline = new Promise<never>((_resolve, reject) => {
          timeout = setTimeout(() => {
            controller.abort();
            reject(
              new DiscordRequestError(
                "request timed out",
                method !== "POST",
                undefined,
                method === "POST",
              ),
            );
          }, timeoutMs);
        });
        return await Promise.race([request(), deadline]);
      } catch (error) {
        if (timeout) {
          clearTimeout(timeout);
          timeout = undefined;
        }
        const requestError =
          error instanceof DiscordRequestError
            ? error
            : new DiscordRequestError(
                error instanceof DOMException && error.name === "AbortError"
                  ? "request timed out"
                  : "Discord network request failed",
                method !== "POST",
                undefined,
                method === "POST",
              );
        lastError = requestError;
        if (!requestError.retryable || attempt === attempts) throw requestError;
        if (
          requestError.retryAfterMs !== undefined &&
          requestError.retryAfterMs > this.maxRetryDelayMs
        ) {
          throw new DiscordRequestError(
            `${requestError.message}; retry window exceeds the notification deadline`,
            false,
          );
        }
        const delay =
          requestError.retryAfterMs ??
          Math.min(250 * 2 ** (attempt - 1), this.maxRetryDelayMs);
        await this.sleep(delay);
      } finally {
        if (timeout) clearTimeout(timeout);
      }
    }
    throw lastError ?? new DiscordRequestError("Discord request failed", false);
  }
}

function localDateTime(value: Date): string {
  const formatter = new Intl.DateTimeFormat("en-GB", {
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
    hourCycle: "h23",
    timeZoneName: "short",
  });
  const parts = new Map(
    formatter
      .formatToParts(value)
      .filter((part) => part.type !== "literal")
      .map((part) => [part.type, part.value]),
  );
  return `${parts.get("year")}-${parts.get("month")}-${parts.get("day")} ${parts.get("hour")}:${parts.get("minute")} ${parts.get("timeZoneName") ?? "local"}`;
}

function localTime(value: Date): string {
  return new Intl.DateTimeFormat("en-GB", {
    hour: "2-digit",
    minute: "2-digit",
    hourCycle: "h23",
    timeZoneName: "short",
  }).format(value);
}

function duration(startedAt: Date, completedAt: Date): string {
  const totalSeconds = Math.max(
    0,
    Math.floor((completedAt.getTime() - startedAt.getTime()) / 1_000),
  );
  const hours = Math.floor(totalSeconds / 3_600);
  const minutes = Math.floor((totalSeconds % 3_600) / 60);
  const seconds = totalSeconds % 60;
  return [
    ...(hours ? [`${hours}h`] : []),
    ...(minutes || hours ? [`${minutes}m`] : []),
    `${seconds}s`,
  ].join(" ");
}

function environmentLabel(
  environment: NodeJS.ProcessEnv,
  platform: NodeJS.Platform,
): string {
  if (/^(?:true|1)$/i.test(environment.CODESPACES ?? "")) {
    return "GitHub Codespaces";
  }
  if (/^(?:true|1)$/i.test(environment.GITHUB_ACTIONS ?? "")) {
    return "GitHub Actions";
  }
  if (platform === "win32") return "Windows";
  if (platform === "darwin") return "macOS";
  if (platform === "linux") return "Linux";
  return platform;
}

function modeLabel(mode: DiscordRunMode): string {
  return mode === "retest" ? "Ready-for-Retest" : "Full Test";
}

function embed(
  title: string,
  color: number,
  fields: DiscordEmbedField[],
  timestamp: Date,
): DiscordPayload {
  const bounded = (value: string, maximum: number): string =>
    value.length <= maximum
      ? value
      : `${value.slice(0, Math.max(0, maximum - 3))}...`;
  return {
    embeds: [
      {
        title: bounded(title, 256),
        color,
        fields: fields.map((field) => ({
          ...field,
          name: bounded(field.name, 256),
          value: bounded(field.value || "Not available", 1_024),
        })),
        timestamp: timestamp.toISOString(),
      },
    ],
    allowed_mentions: SAFE_ALLOWED_MENTIONS,
  };
}

class FailOpenDiscordNotifier implements DiscordNotifier {
  private readonly transport?: DiscordWebhookTransport;
  private readonly warn: (message: string) => void;
  private readonly now: () => Date;
  private readonly environment: NodeJS.ProcessEnv;
  private readonly platform: NodeJS.Platform;
  private readonly minimumProgressIntervalMs: number;
  private readonly progressMayCreateLiveMessage: boolean;
  private readonly configurationIssue?: string;
  private warningIssued = false;
  private run?: DiscordRunStartedEvent;
  private liveMessageId?: string;
  private lastProgressCompleted = 0;
  private lastProgressAt = 0;
  private terminal = false;
  private terminalPayload?: DiscordPayload;
  private liveMessageCreationAttempted = false;
  private configurationWarningIssued = false;

  constructor(
    private readonly settings: DiscordNotificationSettings,
    dependencies: DiscordNotifierDependencies,
  ) {
    this.warn = dependencies.warn ?? console.warn;
    this.now = dependencies.now ?? (() => new Date());
    this.environment = dependencies.environment ?? process.env;
    this.platform = dependencies.platform ?? process.platform;
    this.minimumProgressIntervalMs =
      dependencies.minimumProgressIntervalMs ??
      DEFAULT_MIN_PROGRESS_INTERVAL_MS;
    this.progressMayCreateLiveMessage = !settings.discordNotifyStart;
    const validation = validateDiscordWebhookUrl(settings.discordWebhookUrl);
    if (settings.discordNotificationsEnabled && !validation.valid) {
      this.configurationIssue = validation.reason;
    } else if (validation.valid) {
      this.transport = new DiscordWebhookTransport(validation.url, dependencies);
    }
  }

  async runStarted(event: DiscordRunStartedEvent): Promise<void> {
    this.run = event;
    this.lastProgressAt = event.startedAt.getTime();
    if (!this.settings.discordNotifyStart) return;
    const transport = this.availableTransport();
    if (!transport) return;
    const title =
      event.mode === "retest" ? "PGN Retest Started" : "PGN Full Test Started";
    const payload = embed(
      title,
      0x3498db,
      [
        { name: "Run ID", value: event.runId, inline: true },
        { name: "Mode", value: modeLabel(event.mode), inline: true },
        {
          name:
            event.mode === "retest"
              ? "Ready-for-Retest scenarios"
              : "Selected scenarios",
          value: String(event.selectedScenarios),
          inline: true,
        },
        {
          name: "Environment",
          value: environmentLabel(this.environment, this.platform),
          inline: true,
        },
        {
          name: "Google Drive Evidence",
          value: event.googleDriveEvidenceEnabled ? "Enabled" : "Disabled",
          inline: true,
        },
        { name: "Started", value: localDateTime(event.startedAt), inline: true },
      ],
      event.startedAt,
    );
    this.liveMessageCreationAttempted = true;
    const messageId = await this.failOpen(() => transport.execute(payload));
    if (messageId) {
      this.liveMessageId = messageId;
      if (this.terminalPayload) {
        await this.failOpen(() =>
          transport.edit(messageId, this.terminalPayload!, {
            maxAttempts: 1,
            timeoutMs: INTERRUPTION_REQUEST_TIMEOUT_MS,
          }),
        );
      }
    }
  }

  async runProgress(event: DiscordRunProgressEvent): Promise<void> {
    if (
      !this.run ||
      this.terminal ||
      !this.settings.discordNotifyProgress ||
      event.completedScenarios >= event.totalScenarios
    ) {
      return;
    }
    const now = event.updatedAt.getTime();
    const countDue =
      event.completedScenarios - this.lastProgressCompleted >=
      this.settings.discordProgressEvery;
    const timeDue =
      now - this.lastProgressAt >=
      this.settings.discordProgressMinutes * 60_000;
    if (!countDue && !timeDue) return;
    if (now - this.lastProgressAt < this.minimumProgressIntervalMs) return;
    this.lastProgressCompleted = event.completedScenarios;
    this.lastProgressAt = now;
    const transport = this.availableTransport();
    if (!transport) return;
    const payload = this.runningPayload(event);
    if (this.liveMessageId) {
      await this.failOpen(() => transport.edit(this.liveMessageId!, payload));
      if (this.terminalPayload) {
        await this.failOpen(() =>
          transport.edit(this.liveMessageId!, this.terminalPayload!),
        );
      }
    } else if (
      this.progressMayCreateLiveMessage &&
      !this.liveMessageCreationAttempted
    ) {
      this.liveMessageCreationAttempted = true;
      const messageId = await this.failOpen(() => transport.execute(payload));
      if (messageId) {
        this.liveMessageId = messageId;
        if (this.terminalPayload) {
          await this.failOpen(() =>
            transport.edit(messageId, this.terminalPayload!),
          );
        }
      }
    }
  }

  async runCompleted(event: DiscordRunCompletedEvent): Promise<void> {
    if (!this.run || this.terminal) return;
    this.terminal = true;
    const transport = this.availableTransport();
    if (!transport) return;
    const payload = this.completedPayload(event);
    this.terminalPayload = payload;
    if (this.liveMessageId) {
      await this.failOpen(() => transport.edit(this.liveMessageId!, payload));
    }
    if (this.settings.discordNotifyComplete) {
      await this.failOpen(() => transport.execute(payload));
    }
  }

  async runFailed(event: DiscordRunFailedEvent): Promise<void> {
    if (!this.run || this.terminal) return;
    this.terminal = true;
    const transport = this.availableTransport();
    if (!transport) return;
    const payload = this.failedPayload(
      event,
      this.run.mode === "retest" ? "PGN Retest Aborted" : "PGN Test Aborted",
    );
    this.terminalPayload = payload;
    if (this.liveMessageId) {
      await this.failOpen(() => transport.edit(this.liveMessageId!, payload));
    }
    if (this.settings.discordNotifyFailure) {
      await this.failOpen(() => transport.execute(payload));
    }
  }

  async runInterrupted(
    signal: "SIGINT" | "SIGTERM",
    event: DiscordRunProgressEvent,
  ): Promise<void> {
    if (!this.run || this.terminal) return;
    this.terminal = true;
    const transport = this.availableTransport();
    if (!transport) return;
    const interruptedAt = this.now();
    const payload = this.failedPayload(
      {
        ...event,
        failedAt: interruptedAt,
        reason: `Process received ${signal}`,
        workbookProgress: "Saved progressively",
        evidenceProgress: `${event.evidenceUploaded} uploaded`,
      },
      this.run.mode === "retest"
        ? "PGN Retest Interrupted"
        : "PGN Test Interrupted",
    );
    this.terminalPayload = payload;
    const requests: Promise<unknown>[] = [];
    if (this.liveMessageId) {
      requests.push(
        this.failOpen(() =>
          transport.edit(this.liveMessageId!, payload, {
            maxAttempts: 1,
            timeoutMs: INTERRUPTION_REQUEST_TIMEOUT_MS,
          }),
        ),
      );
    }
    if (this.settings.discordNotifyFailure) {
      requests.push(
        this.failOpen(() =>
          transport.execute(payload, {
            maxAttempts: 1,
            timeoutMs: INTERRUPTION_REQUEST_TIMEOUT_MS,
          }),
        ),
      );
    }
    await Promise.all(requests);
  }

  private runningPayload(event: DiscordRunProgressEvent): DiscordPayload {
    const run = this.run!;
    return embed(
      run.mode === "retest" ? "PGN Retest Running" : "PGN Test Running",
      0xf1c40f,
      [
        { name: "Run ID", value: run.runId, inline: true },
        {
          name: "Progress",
          value: `${event.completedScenarios} / ${event.totalScenarios}`,
          inline: true,
        },
        {
          name: "Current scenario",
          value: event.currentScenarioId ?? "Preparing run",
          inline: true,
        },
        {
          name: "Completed scenarios",
          value: String(event.completedScenarios),
          inline: true,
        },
        {
          name: "Technical errors",
          value: String(event.technicalErrors),
          inline: true,
        },
        {
          name: "Evidence uploaded",
          value: String(event.evidenceUploaded),
          inline: true,
        },
        {
          name: "Elapsed",
          value: duration(run.startedAt, event.updatedAt),
          inline: true,
        },
        { name: "Last update", value: localTime(event.updatedAt), inline: true },
      ],
      event.updatedAt,
    );
  }

  private completedPayload(event: DiscordRunCompletedEvent): DiscordPayload {
    const run = this.run!;
    const checkpoint = run.mode === "retest" && event.checkpoint;
    return embed(
      checkpoint
        ? "PGN Retest Checkpoint Saved"
        : run.mode === "retest"
          ? "PGN Retest Completed"
          : "PGN Test Completed",
      checkpoint ? 0xf1c40f : 0x2ecc71,
      [
        { name: "Run ID", value: run.runId, inline: true },
        { name: "Mode", value: modeLabel(run.mode), inline: true },
        {
          name: "Selected",
          value: String(run.selectedScenarios),
          inline: true,
        },
        {
          name: "Executed",
          value: String(event.completedScenarios),
          inline: true,
        },
        {
          name: "Captured",
          value: String(event.capturedScenarios),
          inline: true,
        },
        { name: "Timeouts", value: String(event.timeouts), inline: true },
        {
          name: "Technical errors",
          value: String(event.technicalErrors),
          inline: true,
        },
        {
          name: "Evidence uploaded",
          value: String(event.evidenceUploaded),
          inline: true,
        },
        {
          name: "Evidence upload errors",
          value: String(event.evidenceUploadErrors),
          inline: true,
        },
        {
          name: "Duration",
          value: duration(run.startedAt, event.completedAt),
          inline: true,
        },
        {
          name: "Workbook",
          value: path.basename(run.workbookPath),
          inline: false,
        },
      ],
      event.completedAt,
    );
  }

  private failedPayload(
    event: DiscordRunFailedEvent,
    title: string,
  ): DiscordPayload {
    const run = this.run!;
    return embed(
      title,
      0xe74c3c,
      [
        { name: "Run ID", value: run.runId, inline: true },
        { name: "Mode", value: modeLabel(run.mode), inline: true },
        {
          name: "Last scenario",
          value: event.currentScenarioId ?? "Not started",
          inline: true,
        },
        {
          name: "Completed",
          value: `${event.completedScenarios} / ${event.totalScenarios}`,
          inline: true,
        },
        {
          name: "Reason",
          value: safeOperationalReason(event.reason),
          inline: false,
        },
        {
          name: "Workbook progress",
          value: event.workbookProgress,
          inline: true,
        },
        {
          name: "Evidence progress",
          value: event.evidenceProgress,
          inline: true,
        },
        {
          name: "Duration",
          value: duration(run.startedAt, event.failedAt),
          inline: true,
        },
      ],
      event.failedAt,
    );
  }

  private availableTransport(): DiscordWebhookTransport | undefined {
    if (!this.settings.discordNotificationsEnabled) return undefined;
    if (
      !this.configurationWarningIssued &&
      this.settings.discordConfigurationIssues?.length
    ) {
      this.configurationWarningIssued = true;
      this.warn(
        `[Discord] Notification failed: invalid setting(s): ${this.settings.discordConfigurationIssues.join("; ")}; affected notifications were disabled`,
      );
    }
    if (this.transport) return this.transport;
    if (!this.warningIssued) {
      this.warningIssued = true;
      this.warn(
        `[Discord] Notification failed: ${this.configurationIssue ?? "webhook is unavailable"}`,
      );
    }
    return undefined;
  }

  private async failOpen<Value>(
    operation: () => Promise<Value>,
  ): Promise<Value | undefined> {
    try {
      return await operation();
    } catch (error) {
      this.warn(
        `[Discord] Notification failed: ${safeDiscordError(error, this.settings.discordWebhookUrl)}`,
      );
      return undefined;
    }
  }
}

export function createDiscordNotifier(
  config: DiscordNotificationSettings,
  dependencies: DiscordNotifierDependencies = {},
): DiscordNotifier {
  const notifier = new FailOpenDiscordNotifier(config, dependencies);
  const warn = dependencies.warn ?? console.warn;
  const guard = async (operation: () => Promise<void>): Promise<void> => {
    try {
      await operation();
    } catch (error) {
      try {
        warn(
          `[Discord] Notification failed: ${safeDiscordError(error, config.discordWebhookUrl)}`,
        );
      } catch {
        // Optional telemetry must remain fail-open even if a custom logger fails.
      }
    }
  };
  return {
    runStarted: (event) => guard(() => notifier.runStarted(event)),
    runProgress: (event) => guard(() => notifier.runProgress(event)),
    runCompleted: (event) => guard(() => notifier.runCompleted(event)),
    runFailed: (event) => guard(() => notifier.runFailed(event)),
    runInterrupted: (signal, event) =>
      guard(() => notifier.runInterrupted(signal, event)),
  };
}

export async function validateDiscordWebhook(
  config: Pick<
    DiscordNotificationSettings,
    "discordNotificationsEnabled" | "discordWebhookUrl"
  >,
  options: DiscordValidationOptions = {},
): Promise<DiscordValidationResult> {
  const validation = validateDiscordWebhookUrl(config.discordWebhookUrl);
  if (!validation.valid) {
    return {
      enabled: config.discordNotificationsEnabled,
      configured: Boolean(config.discordWebhookUrl?.trim()),
      valid: false,
      connectivity: "not-tested",
      testNotificationSent: false,
      testNotificationDeliveryUncertain: false,
      reason: validation.reason,
    };
  }
  const transport = new DiscordWebhookTransport(validation.url, options);
  try {
    if (options.sendTest) {
      await transport.execute({
        content: "PGN Sawala Discord notification test successful.",
        allowed_mentions: SAFE_ALLOWED_MENTIONS,
      });
    } else {
      await transport.inspect();
    }
    return {
      enabled: config.discordNotificationsEnabled,
      configured: true,
      valid: true,
      connectivity: "ok",
      testNotificationSent: Boolean(options.sendTest),
      testNotificationDeliveryUncertain: false,
    };
  } catch (error) {
    return {
      enabled: config.discordNotificationsEnabled,
      configured: true,
      valid: true,
      connectivity: "failed",
      testNotificationSent: false,
      testNotificationDeliveryUncertain:
        Boolean(options.sendTest) &&
        error instanceof DiscordRequestError &&
        error.deliveryUncertain,
      reason: safeDiscordError(error, config.discordWebhookUrl),
    };
  }
}

export function discordStatusLines(
  config: DiscordNotificationSettings,
  validation?: DiscordValidationResult,
): string {
  const webhook = validateDiscordWebhookUrl(config.discordWebhookUrl);
  const cadence = config.discordNotifyProgress
    ? `every ${config.discordProgressEvery} scenarios or ${config.discordProgressMinutes} minutes`
    : "disabled";
  const connectivity = validation
    ? validation.connectivity === "ok"
      ? "OK"
      : validation.connectivity === "failed"
        ? "FAILED"
        : "not tested"
    : "not tested";
  return [
    `Enabled ........ ${config.discordNotificationsEnabled ? "YES" : "NO"}`,
    `Webhook ........ ${webhook.valid ? "configured" : "not configured"}`,
    `Events ......... start ${config.discordNotifyStart ? "YES" : "NO"}, progress ${config.discordNotifyProgress ? "YES" : "NO"}, complete ${config.discordNotifyComplete ? "YES" : "NO"}, failure ${config.discordNotifyFailure ? "YES" : "NO"}`,
    `Cadence ........ ${cadence}`,
    ...(config.discordConfigurationIssues?.length
      ? [
          `Configuration . WARNING: ${config.discordConfigurationIssues.join("; ")}`,
        ]
      : []),
    `Connectivity ... ${connectivity}`,
  ].join("\n");
}

export interface InterruptionSignalSource {
  once(
    event: "SIGINT" | "SIGTERM",
    listener: () => void,
  ): unknown;
  removeListener(
    event: "SIGINT" | "SIGTERM",
    listener: () => void,
  ): unknown;
}

export function registerDiscordInterruptionHandlers(options: {
  notifier: DiscordNotifier;
  progress: () => DiscordRunProgressEvent;
  settle?: () => Promise<void> | void;
  cleanup?: () => Promise<void> | void;
  signalSource?: InterruptionSignalSource;
  terminate?: (signal: "SIGINT" | "SIGTERM") => void;
  notificationTimeoutMs?: number;
}): () => void {
  const signalSource = options.signalSource ?? process;
  let handling = false;
  const listeners = new Map<"SIGINT" | "SIGTERM", () => void>();
  const remove = (): void => {
    for (const [signal, listener] of listeners) {
      signalSource.removeListener(signal, listener);
    }
  };
  const terminate =
    options.terminate ??
    ((signal: "SIGINT" | "SIGTERM") => {
      process.kill(process.pid, signal);
    });
  const notificationTimeoutMs = Math.max(
    1,
    options.notificationTimeoutMs ?? 1_500,
  );
  for (const signal of ["SIGINT", "SIGTERM"] as const) {
    const listener = (): void => {
      if (handling) return;
      handling = true;
      remove();
      let timeout: NodeJS.Timeout | undefined;
      const notification = Promise.resolve()
        .then(() => options.notifier.runInterrupted(signal, options.progress()))
        .catch(() => undefined);
      const settlement = Promise.resolve()
        .then(() => options.settle?.())
        .catch(() => undefined);
      const cleanup = Promise.resolve()
        .then(() => options.cleanup?.())
        .catch(() => undefined);
      const deadline = new Promise<void>((resolve) => {
        timeout = setTimeout(resolve, notificationTimeoutMs);
      });
      const shutdown = Promise.all([notification, settlement, cleanup]).then(
        () => undefined,
      );
      void Promise.race([shutdown, deadline])
        .finally(() => {
          if (timeout) clearTimeout(timeout);
          terminate(signal);
        });
    };
    listeners.set(signal, listener);
    signalSource.once(signal, listener);
  }
  return remove;
}
