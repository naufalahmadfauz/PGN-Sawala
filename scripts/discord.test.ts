import assert from "node:assert/strict";
import test from "node:test";
import { parseDiscordDemoArgs, runDiscordDemo } from "./discord-demo";
import { parseDiscordValidationArgs } from "./discord-validate";
import {
  createDiscordNotifier,
  discordStatusLines,
  registerDiscordInterruptionHandlers,
  safeDiscordError,
  validateDiscordWebhook,
  validateDiscordWebhookUrl,
  type DiscordNotificationSettings,
  type DiscordNotifier,
  type DiscordRunProgressEvent,
  type InterruptionSignalSource,
} from "../src/notifications/discord";

const webhookUrl = "https://discord.com/api/webhooks/123456789/test-token";

function settings(
  overrides: Partial<DiscordNotificationSettings> = {},
): DiscordNotificationSettings {
  return {
    discordNotificationsEnabled: true,
    discordWebhookUrl: webhookUrl,
    discordProgressEvery: 5,
    discordProgressMinutes: 2,
    discordNotifyStart: true,
    discordNotifyProgress: true,
    discordNotifyComplete: true,
    discordNotifyFailure: true,
    ...overrides,
  };
}

function response(
  status = 200,
  body: unknown = { id: "987654321" },
  headers: Record<string, string> = {},
): Response {
  return new Response(JSON.stringify(body), { status, headers });
}

interface RequestRecord {
  url: URL;
  init?: RequestInit;
}

function recordingFetch(
  implementation: (call: RequestRecord, index: number) => Promise<Response> =
    async (call) =>
      call.init?.method === "GET"
        ? response(200, { id: "123456789", type: 1 })
        : response(),
): { calls: RequestRecord[]; fetch: typeof fetch } {
  const calls: RequestRecord[] = [];
  return {
    calls,
    fetch: (async (input: URL | RequestInfo, init?: RequestInit) => {
      const url = input instanceof URL ? input : new URL(String(input));
      const call = { url, init };
      calls.push(call);
      return implementation(call, calls.length - 1);
    }) as typeof fetch,
  };
}

function payload(call: RequestRecord): Record<string, unknown> {
  const body = call.init?.body;
  assert.equal(typeof body, "string");
  return JSON.parse(body as string) as Record<string, unknown>;
}

function startEvent(mode: "full" | "retest" = "full") {
  return {
    runId: "PGN-20260903-120000",
    mode,
    selectedScenarios: 10,
    startedAt: new Date("2026-09-03T12:00:00.000Z"),
    googleDriveEvidenceEnabled: true,
    workbookPath: "/private/reports/PGN Executed.xlsx",
  } as const;
}

function progressEvent(
  overrides: Partial<DiscordRunProgressEvent> = {},
): DiscordRunProgressEvent {
  return {
    completedScenarios: 5,
    totalScenarios: 10,
    currentScenarioId: "TC-005",
    capturedScenarios: 4,
    timeouts: 1,
    technicalErrors: 0,
    evidenceUploaded: 4,
    evidenceUploadErrors: 0,
    updatedAt: new Date("2026-09-03T12:01:00.000Z"),
    ...overrides,
  };
}

function embedTitle(body: Record<string, unknown>): string {
  const embeds = body.embeds as Array<{ title: string }>;
  return embeds[0]?.title ?? "";
}

function embedField(body: Record<string, unknown>, name: string): string {
  const embeds = body.embeds as Array<{
    fields: Array<{ name: string; value: string }>;
  }>;
  return embeds[0]?.fields.find((field) => field.name === name)?.value ?? "";
}

function assertSafeMentions(body: Record<string, unknown>): void {
  assert.deepEqual(body.allowed_mentions, {
    parse: [],
    users: [],
    roles: [],
    replied_user: false,
  });
}

test("validates only Discord Incoming Webhook URL shapes", () => {
  assert.equal(validateDiscordWebhookUrl(webhookUrl).valid, true);
  assert.equal(
    validateDiscordWebhookUrl(
      "https://canary.discord.com/api/v10/webhooks/123456/token",
    ).valid,
    true,
  );
  for (const unsafe of [
    undefined,
    "http://discord.com/api/webhooks/123/token",
    "https://discord.example/api/webhooks/123/token",
    "https://user:password@discord.com/api/webhooks/123/token",
    "https://discord.com:444/api/webhooks/123/token",
    "https://discord.com/api/webhooks/123",
    "https://discord.com/api/webhooks/123/token/messages/456",
  ]) {
    assert.equal(validateDiscordWebhookUrl(unsafe).valid, false);
  }
});

test("validation CLI requires explicit and unambiguous send-test intent", () => {
  assert.deepEqual(parseDiscordValidationArgs([]), { sendTest: false });
  assert.deepEqual(parseDiscordValidationArgs(["--send-test"]), {
    sendTest: true,
  });
  assert.throws(() => parseDiscordValidationArgs(["--unknown"]), /Usage:/);
  assert.throws(
    () => parseDiscordValidationArgs(["--send-test", "--send-test"]),
    /Usage:/,
  );
});

test("demo CLI accepts only one simulated outcome", () => {
  assert.deepEqual(parseDiscordDemoArgs([]), { outcome: "complete" });
  assert.deepEqual(parseDiscordDemoArgs(["--fail"]), { outcome: "fail" });
  assert.deepEqual(parseDiscordDemoArgs(["--interrupt"]), {
    outcome: "interrupt",
  });
  for (const invalid of [
    ["--unknown"],
    ["--fail", "--fail"],
    ["--interrupt", "--interrupt"],
    ["--fail", "--interrupt"],
  ]) {
    assert.throws(() => parseDiscordDemoArgs(invalid), /Usage:/);
  }
});

test("demo stops safely when Discord is disabled or the webhook is unavailable", async () => {
  for (const fixture of [
    {
      config: settings({ discordNotificationsEnabled: false }),
      status: "disabled",
      message: "Discord notifications are disabled.",
    },
    {
      config: settings({ discordWebhookUrl: undefined }),
      status: "missing-webhook",
      message: "Discord webhook: not configured",
    },
    {
      config: settings({ discordWebhookUrl: "https://example.invalid/webhook" }),
      status: "invalid-webhook",
      message: "Discord webhook: invalid configuration",
    },
    {
      config: settings({
        discordConfigurationIssues: [
          "DISCORD_PROGRESS_EVERY must be a whole number of 1 or greater",
        ],
      }),
      status: "invalid-settings",
      message: "Discord notification configuration is invalid.",
    },
  ] as const) {
    const logs: string[] = [];
    let requests = 0;
    const result = await runDiscordDemo(
      fixture.config,
      { outcome: "complete" },
      {
        log: (message) => logs.push(message),
        delay: async () => {
          throw new Error("A blocked demo must not wait");
        },
        notifier: {
          fetch: (async () => {
            requests += 1;
            throw new Error("A blocked demo must not contact Discord");
          }) as typeof fetch,
        },
      },
    );
    assert.equal(result.status, fixture.status);
    assert.equal(requests, 0);
    assert(logs.includes(fixture.message));
    assert(!logs.join("\n").includes(webhookUrl));
  }
});

test("normal demo creates and edits one live message before completion", async () => {
  const transport = recordingFetch();
  const logs: string[] = [];
  const warnings: string[] = [];
  const delays: number[] = [];
  const result = await runDiscordDemo(
    settings(),
    { outcome: "complete" },
    {
      now: () => new Date("2026-09-04T22:00:00.000Z"),
      delay: async (milliseconds) => {
        delays.push(milliseconds);
      },
      log: (message) => logs.push(message),
      warn: (message) => warnings.push(message),
      notifier: { fetch: transport.fetch },
    },
  );

  assert.equal(result.status, "completed");
  assert.equal(result.runId, "DEMO-20260904T220000Z");
  assert.deepEqual(delays, [1_000, 1_000, 1_000]);
  assert.deepEqual(
    transport.calls.map((call) => call.init?.method),
    ["POST", "PATCH", "PATCH", "PATCH", "POST"],
  );
  for (const call of transport.calls.slice(1, 4)) {
    assert.match(call.url.pathname, /\/messages\/987654321$/);
  }
  assert.equal(
    embedTitle(payload(transport.calls[0]!)),
    "PGN Discord Demo Started",
  );
  assert.equal(embedField(payload(transport.calls[0]!), "Mode"), "Discord Demo");
  assert.equal(embedField(payload(transport.calls[1]!), "Progress"), "1 / 10");
  assert.equal(embedField(payload(transport.calls[2]!), "Progress"), "5 / 10");
  assert.equal(
    embedTitle(payload(transport.calls[3]!)),
    "PGN Discord Demo Completed",
  );
  assert.deepEqual(payload(transport.calls[3]!), payload(transport.calls[4]!));
  for (const call of transport.calls) assertSafeMentions(payload(call));
  assert.equal(warnings.length, 0);
  assert.match(logs.join("\n"), /Demo event: STARTED/);
  assert.match(logs.join("\n"), /Demo event: RUNNING 1 \/ 10/);
  assert.match(logs.join("\n"), /Demo event: RUNNING 5 \/ 10/);
  assert.match(logs.join("\n"), /Demo event: COMPLETED 10 \/ 10/);
  assert(!logs.join("\n").includes(webhookUrl));
});

test("failure demo edits the live message and posts a fake aborted event", async () => {
  const transport = recordingFetch();
  const logs: string[] = [];
  const result = await runDiscordDemo(
    settings(),
    { outcome: "fail" },
    {
      now: () => new Date("2026-09-04T22:00:00.000Z"),
      delay: async () => undefined,
      log: (message) => logs.push(message),
      notifier: { fetch: transport.fetch },
    },
  );

  assert.equal(result.status, "simulated-failure");
  assert.deepEqual(
    transport.calls.map((call) => call.init?.method),
    ["POST", "PATCH", "PATCH", "POST"],
  );
  const failureEdit = payload(transport.calls[2]!);
  assert.equal(embedTitle(failureEdit), "PGN Discord Demo Aborted");
  assert.equal(
    embedField(failureEdit, "Reason"),
    "Simulated failure for Discord notification testing",
  );
  assert.equal(embedField(failureEdit, "Progress"), "5 / 10");
  assert.deepEqual(failureEdit, payload(transport.calls[3]!));
  assert.match(
    logs.at(-1) ?? "",
    /Demo result: simulated failure notification sent successfully/,
  );
});

test("interruption demo emits a simulated SIGINT without killing the process", async () => {
  const transport = recordingFetch();
  const logs: string[] = [];
  const result = await runDiscordDemo(
    settings(),
    { outcome: "interrupt" },
    {
      now: () => new Date("2026-09-04T22:00:00.000Z"),
      delay: async () => undefined,
      log: (message) => logs.push(message),
      notifier: { fetch: transport.fetch },
    },
  );

  assert.equal(result.status, "simulated-interruption");
  assert.deepEqual(
    transport.calls.map((call) => call.init?.method),
    ["POST", "PATCH", "PATCH", "POST"],
  );
  const interruption = payload(transport.calls[2]!);
  assert.equal(embedTitle(interruption), "PGN Discord Demo Interrupted");
  assert.equal(embedField(interruption, "Reason"), "Simulated SIGINT");
  assert.equal(
    embedField(interruption, "Workbook progress"),
    "Demo only; no workbook modified",
  );
  assert.match(
    logs.at(-1) ?? "",
    /Demo result: simulated interruption notification sent successfully/,
  );
});

test("demo respects a mocked 429 retry delay within bounded attempts", async () => {
  const retryDelays: number[] = [];
  const transport = recordingFetch(async (_call, index) =>
    index === 0
      ? response(429, { retry_after: 0.25 }, { "retry-after": "0.25" })
      : response(),
  );
  const result = await runDiscordDemo(
    settings(),
    { outcome: "complete" },
    {
      now: () => new Date("2026-09-04T22:00:00.000Z"),
      delay: async () => undefined,
      log: () => undefined,
      warn: () => undefined,
      notifier: {
        fetch: transport.fetch,
        sleep: async (milliseconds) => {
          retryDelays.push(milliseconds);
        },
      },
    },
  );

  assert.equal(result.status, "completed");
  assert.deepEqual(retryDelays, [250]);
  assert.equal(transport.calls.length, 6);
});

test("demo transport failures stay fail-open and redact the webhook", async () => {
  for (const failingTransport of [
    recordingFetch(async () => response(500, {})),
    recordingFetch(
      async (call) =>
        new Promise<Response>((_resolve, reject) => {
          call.init?.signal?.addEventListener("abort", () => {
            reject(new DOMException(`timed out ${webhookUrl}`, "AbortError"));
          });
        }),
    ),
  ]) {
    const logs: string[] = [];
    const warnings: string[] = [];
    const result = await runDiscordDemo(
      settings(),
      { outcome: "complete" },
      {
        now: () => new Date("2026-09-04T22:00:00.000Z"),
        delay: async () => undefined,
        log: (message) => logs.push(message),
        warn: (message) => warnings.push(message),
        notifier: {
          fetch: failingTransport.fetch,
          timeoutMs: 1,
          maxAttempts: 1,
        },
      },
    );

    assert.equal(result.status, "delivery-warning");
    assert(warnings.length >= 1);
    assert(![...logs, ...warnings].join("\n").includes(webhookUrl));
  }
});

test("safe validation inspects with GET and explicit tests send one message", async () => {
  const inspection = recordingFetch();
  const inspected = await validateDiscordWebhook(
    { discordNotificationsEnabled: false, discordWebhookUrl: webhookUrl },
    { fetch: inspection.fetch },
  );
  assert.equal(inspected.connectivity, "ok");
  assert.equal(inspected.testNotificationSent, false);
  assert.equal(inspection.calls.length, 1);
  assert.equal(inspection.calls[0]?.init?.method, "GET");
  assert.equal(inspection.calls[0]?.init?.body, undefined);

  const testSend = recordingFetch();
  const sent = await validateDiscordWebhook(
    { discordNotificationsEnabled: true, discordWebhookUrl: webhookUrl },
    { fetch: testSend.fetch, sendTest: true },
  );
  assert.equal(sent.testNotificationSent, true);
  assert.equal(testSend.calls.length, 1);
  assert.equal(testSend.calls[0]?.init?.method, "POST");
  assert.equal(testSend.calls[0]?.url.searchParams.get("wait"), "true");
  const body = payload(testSend.calls[0]!);
  assert.equal(
    body.content,
    "PGN Sawala Discord notification test successful.",
  );
  assertSafeMentions(body);
});

test("explicit test sends report ambiguous POST delivery as uncertain", async () => {
  for (const transport of [
    recordingFetch(async () => {
      throw new Error("response was lost");
    }),
    recordingFetch(async () => response(500, {})),
    recordingFetch(async () => response(408, {})),
    recordingFetch(async () => response(200, { unexpected: true })),
  ]) {
    const result = await validateDiscordWebhook(settings(), {
      fetch: transport.fetch,
      sendTest: true,
      maxAttempts: 1,
    });
    assert.equal(result.connectivity, "failed");
    assert.equal(result.testNotificationSent, false);
    assert.equal(result.testNotificationDeliveryUncertain, true);
    assert.equal(transport.calls.length, 1);
  }
});

test("creates one live message, edits it for progress, then finalizes and posts completion", async () => {
  const transport = recordingFetch();
  const notifier = createDiscordNotifier(settings(), {
    fetch: transport.fetch,
    environment: { CODESPACES: "true" },
    platform: "linux",
    minimumProgressIntervalMs: 0,
  });

  await notifier.runStarted(startEvent());
  await notifier.runProgress(progressEvent());
  await notifier.runCompleted({
    ...progressEvent({ completedScenarios: 10, currentScenarioId: "TC-010" }),
    completedAt: new Date("2026-09-03T12:02:00.000Z"),
  });

  assert.deepEqual(
    transport.calls.map((call) => call.init?.method),
    ["POST", "PATCH", "PATCH", "POST"],
  );
  assert.match(transport.calls[1]!.url.pathname, /\/messages\/987654321$/);
  assert.equal(transport.calls[1]!.url.href, transport.calls[2]!.url.href);
  const started = payload(transport.calls[0]!);
  const running = payload(transport.calls[1]!);
  const completedEdit = payload(transport.calls[2]!);
  const completedPost = payload(transport.calls[3]!);
  assert.equal(embedTitle(started), "PGN Full Test Started");
  assert.equal(embedField(started, "Environment"), "GitHub Codespaces");
  assert.equal(embedTitle(running), "PGN Test Running");
  assert.equal(embedTitle(completedEdit), "PGN Test Completed");
  assert.deepEqual(completedEdit, completedPost);
  assert.equal(embedField(completedPost, "Workbook"), "PGN Executed.xlsx");
  for (const call of transport.calls) {
    const body = payload(call);
    assertSafeMentions(body);
    assert(!JSON.stringify(body).includes(webhookUrl));
    assert(!JSON.stringify(body).includes("/private/reports"));
  }
});

test("labels Ready-for-Retest runs throughout their lifecycle", async () => {
  const transport = recordingFetch();
  const notifier = createDiscordNotifier(settings(), { fetch: transport.fetch });
  await notifier.runStarted(startEvent("retest"));
  await notifier.runCompleted({
    ...progressEvent({ completedScenarios: 10 }),
    completedAt: new Date("2026-09-03T12:02:00.000Z"),
  });

  assert.equal(embedTitle(payload(transport.calls[0]!)), "PGN Retest Started");
  assert.equal(
    embedTitle(payload(transport.calls[1]!)),
    "PGN Retest Completed",
  );
  assert.equal(
    embedTitle(payload(transport.calls[2]!)),
    "PGN Retest Completed",
  );
});

test("throttles progress by count, elapsed time, and a minimum edit interval", async () => {
  const transport = recordingFetch();
  const notifier = createDiscordNotifier(settings(), { fetch: transport.fetch });
  await notifier.runStarted(startEvent());
  await notifier.runProgress(
    progressEvent({ updatedAt: new Date("2026-09-03T12:00:01.000Z") }),
  );
  assert.equal(transport.calls.length, 1);

  await notifier.runProgress(
    progressEvent({ updatedAt: new Date("2026-09-03T12:00:11.000Z") }),
  );
  assert.equal(transport.calls.length, 2);

  await notifier.runProgress(
    progressEvent({
      completedScenarios: 6,
      updatedAt: new Date("2026-09-03T12:02:11.000Z"),
    }),
  );
  assert.equal(transport.calls.length, 3);

  await notifier.runProgress(
    progressEvent({
      completedScenarios: 10,
      updatedAt: new Date("2026-09-03T12:05:00.000Z"),
    }),
  );
  assert.equal(transport.calls.length, 3);
});

test("failure messages are fresh, operational, redacted, and mode-specific", async () => {
  const transport = recordingFetch();
  const notifier = createDiscordNotifier(settings(), { fetch: transport.fetch });
  await notifier.runStarted(startEvent("retest"));
  await notifier.runFailed({
    ...progressEvent(),
    failedAt: new Date("2026-09-03T12:01:00.000Z"),
    reason: `Request ${webhookUrl} for +62 812-3456-7890 failed token=secret-value Authorization: Bearer bearer-secret\nstack details`,
    workbookProgress: "Saved through TC-004",
    evidenceProgress: "4 uploaded, 0 errors",
  });

  assert.deepEqual(
    transport.calls.map((call) => call.init?.method),
    ["POST", "PATCH", "POST"],
  );
  const failure = payload(transport.calls[2]!);
  const serialized = JSON.stringify(failure);
  assert.equal(embedTitle(failure), "PGN Retest Aborted");
  assert.match(embedField(failure, "Reason"), /\[REDACTED\]/);
  assert(!serialized.includes(webhookUrl));
  assert(!serialized.includes("812-3456-7890"));
  assert(!serialized.includes("secret-value"));
  assert(!serialized.includes("bearer-secret"));
  assert(!serialized.includes("stack details"));
  assertSafeMentions(failure);
});

test("respects disabled and per-event notification switches", async () => {
  const disabledTransport = recordingFetch();
  const disabled = createDiscordNotifier(
    settings({ discordNotificationsEnabled: false }),
    { fetch: disabledTransport.fetch },
  );
  await disabled.runStarted(startEvent());
  await disabled.runProgress(progressEvent());
  assert.equal(disabledTransport.calls.length, 0);

  const finalOnlyTransport = recordingFetch();
  const finalOnly = createDiscordNotifier(
    settings({
      discordNotifyStart: false,
      discordNotifyProgress: false,
      discordNotifyComplete: true,
    }),
    { fetch: finalOnlyTransport.fetch },
  );
  await finalOnly.runStarted(startEvent());
  await finalOnly.runProgress(progressEvent());
  await finalOnly.runCompleted({
    ...progressEvent({ completedScenarios: 10 }),
    completedAt: new Date("2026-09-03T12:02:00.000Z"),
  });
  assert.deepEqual(
    finalOnlyTransport.calls.map((call) => call.init?.method),
    ["POST"],
  );

  const progressOnlyTransport = recordingFetch();
  const progressOnly = createDiscordNotifier(
    settings({ discordNotifyStart: false, discordNotifyProgress: true }),
    { fetch: progressOnlyTransport.fetch, minimumProgressIntervalMs: 0 },
  );
  await progressOnly.runStarted(startEvent());
  await progressOnly.runProgress(progressEvent());
  await progressOnly.runProgress(
    progressEvent({
      completedScenarios: 6,
      updatedAt: new Date("2026-09-03T12:03:00.000Z"),
    }),
  );
  assert.deepEqual(
    progressOnlyTransport.calls.map((call) => call.init?.method),
    ["POST", "PATCH"],
  );
  assert.equal(
    embedTitle(payload(progressOnlyTransport.calls[0]!)),
    "PGN Test Running",
  );
});

test("terminal flags suppress fresh events while still finalizing a live card", async () => {
  const completionTransport = recordingFetch();
  const completionNotifier = createDiscordNotifier(
    settings({ discordNotifyComplete: false }),
    { fetch: completionTransport.fetch },
  );
  await completionNotifier.runStarted(startEvent());
  await completionNotifier.runCompleted({
    ...progressEvent({ completedScenarios: 10 }),
    completedAt: new Date("2026-09-03T12:02:00.000Z"),
  });
  assert.deepEqual(
    completionTransport.calls.map((call) => call.init?.method),
    ["POST", "PATCH"],
  );

  const failureTransport = recordingFetch();
  const failureNotifier = createDiscordNotifier(
    settings({ discordNotifyFailure: false }),
    { fetch: failureTransport.fetch },
  );
  await failureNotifier.runStarted(startEvent());
  await failureNotifier.runFailed({
    ...progressEvent(),
    failedAt: new Date("2026-09-03T12:01:00.000Z"),
    reason: "Technical error",
    workbookProgress: "Saved progressively",
    evidenceProgress: "4 uploaded",
  });
  assert.deepEqual(
    failureTransport.calls.map((call) => call.init?.method),
    ["POST", "PATCH"],
  );

  const interruptionTransport = recordingFetch();
  const interruptionNotifier = createDiscordNotifier(
    settings({ discordNotifyFailure: false }),
    { fetch: interruptionTransport.fetch },
  );
  await interruptionNotifier.runStarted(startEvent());
  await interruptionNotifier.runInterrupted("SIGINT", progressEvent());
  assert.deepEqual(
    interruptionTransport.calls.map((call) => call.init?.method),
    ["POST", "PATCH"],
  );
  assert.equal(
    embedTitle(payload(interruptionTransport.calls[1]!)),
    "PGN Test Interrupted",
  );
});

test("interruption posts one fresh mode-specific event", async () => {
  const transport = recordingFetch();
  const notifier = createDiscordNotifier(settings(), {
    fetch: transport.fetch,
    now: () => new Date("2026-09-03T12:01:30.000Z"),
  });
  await notifier.runStarted(startEvent("retest"));
  await notifier.runInterrupted("SIGTERM", progressEvent());

  assert.deepEqual(
    transport.calls.map((call) => call.init?.method),
    ["POST", "PATCH", "POST"],
  );
  const interrupted = payload(transport.calls[2]!);
  assert.equal(embedTitle(interrupted), "PGN Retest Interrupted");
  assert.equal(embedField(interrupted, "Reason"), "Process received SIGTERM");
  assertSafeMentions(interrupted);
});

test("interruption corrects in-flight start and progress status writes", async () => {
  let resolveStart: ((value: Response) => void) | undefined;
  const pendingStart = new Promise<Response>((resolve) => {
    resolveStart = resolve;
  });
  const startRaceTransport = recordingFetch(async (_call, index) =>
    index === 0 ? pendingStart : response(),
  );
  const startRaceNotifier = createDiscordNotifier(settings(), {
    fetch: startRaceTransport.fetch,
  });
  const starting = startRaceNotifier.runStarted(startEvent());
  await startRaceNotifier.runInterrupted("SIGTERM", progressEvent());
  resolveStart?.(response());
  await starting;
  assert.deepEqual(
    startRaceTransport.calls.map((call) => call.init?.method),
    ["POST", "POST", "PATCH"],
  );
  assert.equal(
    embedTitle(payload(startRaceTransport.calls.at(-1)!)),
    "PGN Test Interrupted",
  );

  let resolveProgress: ((value: Response) => void) | undefined;
  const pendingProgress = new Promise<Response>((resolve) => {
    resolveProgress = resolve;
  });
  const progressRaceTransport = recordingFetch(async (_call, index) =>
    index === 1 ? pendingProgress : response(),
  );
  const progressRaceNotifier = createDiscordNotifier(settings(), {
    fetch: progressRaceTransport.fetch,
    minimumProgressIntervalMs: 0,
  });
  await progressRaceNotifier.runStarted(startEvent());
  const progressing = progressRaceNotifier.runProgress(progressEvent());
  await progressRaceNotifier.runInterrupted("SIGINT", progressEvent());
  resolveProgress?.(response());
  await progressing;
  assert.deepEqual(
    progressRaceTransport.calls.map((call) => call.init?.method),
    ["POST", "PATCH", "PATCH", "POST", "PATCH"],
  );
  assert.equal(
    embedTitle(payload(progressRaceTransport.calls.at(-1)!)),
    "PGN Test Interrupted",
  );
});

test("retries safe transient requests with bounded delays", async () => {
  const waits: number[] = [];
  const transport = recordingFetch(async (_call, index) => {
    if (index === 0) {
      return response(429, { retry_after: 0.5 }, { "retry-after": "0.5" });
    }
    if (index === 1) return response(500, {});
    return response(200, { id: "123456789", type: 1 });
  });
  const result = await validateDiscordWebhook(settings(), {
    fetch: transport.fetch,
    sleep: async (milliseconds) => {
      waits.push(milliseconds);
    },
  });

  assert.equal(result.connectivity, "ok");
  assert.equal(transport.calls.length, 3);
  assert.deepEqual(waits, [500, 500]);
});

test("fails open instead of retrying long rate limits or ambiguous POST failures", async () => {
  const longRateLimit = recordingFetch(async () =>
    response(429, { retry_after: 30 }, { "retry-after": "30" }),
  );
  const rateLimitWarnings: string[] = [];
  const rateLimited = createDiscordNotifier(settings(), {
    fetch: longRateLimit.fetch,
    sleep: async () => {
      throw new Error("A long retry window must not be slept");
    },
    warn: (message) => rateLimitWarnings.push(message),
  });
  await rateLimited.runStarted(startEvent());
  assert.equal(longRateLimit.calls.length, 1);
  assert.match(rateLimitWarnings[0]!, /retry window exceeds/);

  for (const failingFetch of [
    recordingFetch(async () => response(500, {})),
    recordingFetch(async () => {
      throw new Error("response was lost");
    }),
  ]) {
    const notifier = createDiscordNotifier(settings(), {
      fetch: failingFetch.fetch,
      sleep: async () => undefined,
      warn: () => undefined,
      minimumProgressIntervalMs: 0,
    });
    await notifier.runStarted(startEvent());
    await notifier.runProgress(progressEvent());
    assert.equal(failingFetch.calls.length, 1);
  }
});

test("an ambiguous progress POST is not repeated on later intervals", async () => {
  const transport = recordingFetch(async () => {
    throw new Error("response was lost");
  });
  const notifier = createDiscordNotifier(
    settings({ discordNotifyStart: false }),
    {
      fetch: transport.fetch,
      minimumProgressIntervalMs: 0,
      warn: () => undefined,
    },
  );
  await notifier.runStarted(startEvent());
  await notifier.runProgress(progressEvent({ totalScenarios: 20 }));
  await notifier.runProgress(
    progressEvent({
      completedScenarios: 10,
      totalScenarios: 20,
      updatedAt: new Date("2026-09-03T12:03:00.000Z"),
    }),
  );
  assert.equal(transport.calls.length, 1);
});

test("permanent, malformed, network, and timeout failures remain fail-open", async () => {
  for (const fixture of [
    {
      fetch: recordingFetch(async () => response(404, {})),
      expectedCalls: 1,
      warning: /HTTP 404/,
    },
    {
      fetch: recordingFetch(async () => response(200, { unexpected: true })),
      expectedCalls: 1,
      warning: /invalid message response/,
    },
    {
      fetch: recordingFetch(async () => {
        throw new Error(`network rejected ${webhookUrl}`);
      }),
      expectedCalls: 1,
      warning: /network request failed/,
    },
  ]) {
    const warnings: string[] = [];
    const notifier = createDiscordNotifier(settings(), {
      fetch: fixture.fetch.fetch,
      sleep: async () => undefined,
      warn: (message) => warnings.push(message),
    });
    await assert.doesNotReject(() => notifier.runStarted(startEvent()));
    assert.equal(fixture.fetch.calls.length, fixture.expectedCalls);
    assert.equal(warnings.length, 1);
    assert.match(warnings[0]!, fixture.warning);
    assert(!warnings[0]!.includes(webhookUrl));
  }

  const timeoutFetch = recordingFetch(
    async (call) =>
      new Promise<Response>((_resolve, reject) => {
        call.init?.signal?.addEventListener("abort", () => {
          reject(new DOMException("aborted", "AbortError"));
        });
      }),
  );
  const warnings: string[] = [];
  const timedOut = createDiscordNotifier(settings(), {
    fetch: timeoutFetch.fetch,
    sleep: async () => undefined,
    timeoutMs: 1,
    warn: (message) => warnings.push(message),
  });
  await assert.doesNotReject(() => timedOut.runStarted(startEvent()));
  assert.equal(timeoutFetch.calls.length, 1);
  assert.match(warnings[0]!, /timed out/);
});

test("the request deadline includes successful response-body parsing", async () => {
  const stalledBody = recordingFetch(async () =>
    ({
      ok: true,
      status: 200,
      headers: new Headers(),
      json: async () => new Promise<unknown>(() => undefined),
    }) as Response,
  );
  const warnings: string[] = [];
  const notifier = createDiscordNotifier(settings(), {
    fetch: stalledBody.fetch,
    timeoutMs: 1,
    warn: (message) => warnings.push(message),
  });

  await notifier.runStarted(startEvent());
  assert.equal(stalledBody.calls.length, 1);
  assert.match(warnings[0]!, /timed out/);
});

test("GET inspection rejects malformed and non-Incoming webhook responses", async () => {
  for (const body of [{ id: "123456789" }, { id: "123456789", type: 2 }]) {
    const transport = recordingFetch(async () => response(200, body));
    const result = await validateDiscordWebhook(settings(), {
      fetch: transport.fetch,
    });
    assert.equal(result.connectivity, "failed");
    assert.match(result.reason ?? "", /Incoming Webhook/);
  }
});

test("retest checkpoints use a nonterminal checkpoint title", async () => {
  const transport = recordingFetch();
  const notifier = createDiscordNotifier(settings(), { fetch: transport.fetch });
  await notifier.runStarted(startEvent("retest"));
  await notifier.runCompleted({
    ...progressEvent({ completedScenarios: 8 }),
    completedAt: new Date("2026-09-03T12:02:00.000Z"),
    checkpoint: true,
  });
  assert.equal(
    embedTitle(payload(transport.calls[1]!)),
    "PGN Retest Checkpoint Saved",
  );
  assert.equal(
    embedTitle(payload(transport.calls[2]!)),
    "PGN Retest Checkpoint Saved",
  );
});

test("guard remains fail-open when event formatting and the warning logger fail", async () => {
  const transport = recordingFetch();
  const notifier = createDiscordNotifier(settings(), {
    fetch: transport.fetch,
    warn: () => {
      throw new Error("logger unavailable");
    },
  });
  await assert.doesNotReject(() =>
    notifier.runStarted({
      ...startEvent(),
      startedAt: new Date(Number.NaN),
    }),
  );
  assert.equal(transport.calls.length, 0);
});

test("status output never includes the webhook secret", () => {
  const status = discordStatusLines(settings());
  assert.match(status, /Enabled .* YES/);
  assert.match(status, /Webhook .* configured/);
  assert.match(status, /Events .* start YES, progress YES/);
  assert.match(status, /Connectivity .* not tested/);
  assert(!status.includes(webhookUrl));
  assert(!safeDiscordError(new Error(webhookUrl), webhookUrl).includes(webhookUrl));
});

class FakeSignalSource implements InterruptionSignalSource {
  readonly listeners = new Map<"SIGINT" | "SIGTERM", () => void>();

  once(signal: "SIGINT" | "SIGTERM", listener: () => void): void {
    this.listeners.set(signal, listener);
  }

  removeListener(signal: "SIGINT" | "SIGTERM", listener: () => void): void {
    if (this.listeners.get(signal) === listener) this.listeners.delete(signal);
  }

  emit(signal: "SIGINT" | "SIGTERM"): void {
    const listener = this.listeners.get(signal);
    this.listeners.delete(signal);
    listener?.();
  }
}

test("interruption handlers notify once, remove listeners, and terminate", async () => {
  const source = new FakeSignalSource();
  const notifications: string[] = [];
  const notifier: DiscordNotifier = {
    runStarted: async () => undefined,
    runProgress: async () => undefined,
    runCompleted: async () => undefined,
    runFailed: async () => undefined,
    runInterrupted: async (signal) => {
      notifications.push(signal);
    },
  };
  let resolveTermination: (() => void) | undefined;
  const terminated = new Promise<void>((resolve) => {
    resolveTermination = resolve;
  });
  const terminations: string[] = [];
  registerDiscordInterruptionHandlers({
    notifier,
    progress: () => progressEvent(),
    signalSource: source,
    terminate: (signal) => {
      terminations.push(signal);
      resolveTermination?.();
    },
  });

  source.emit("SIGINT");
  source.emit("SIGTERM");
  await terminated;
  assert.deepEqual(notifications, ["SIGINT"]);
  assert.deepEqual(terminations, ["SIGINT"]);
  assert.equal(source.listeners.size, 0);
});

test("interruption waits for status settlement and browser cleanup before terminating", async () => {
  const source = new FakeSignalSource();
  let resolveSettlement: (() => void) | undefined;
  let resolveCleanup: (() => void) | undefined;
  const settlement = new Promise<void>((resolve) => {
    resolveSettlement = resolve;
  });
  const cleanup = new Promise<void>((resolve) => {
    resolveCleanup = resolve;
  });
  const calls: string[] = [];
  const notifier: DiscordNotifier = {
    runStarted: async () => undefined,
    runProgress: async () => undefined,
    runCompleted: async () => undefined,
    runFailed: async () => undefined,
    runInterrupted: async () => {
      calls.push("notification");
    },
  };
  let resolveTermination: (() => void) | undefined;
  const terminated = new Promise<void>((resolve) => {
    resolveTermination = resolve;
  });
  registerDiscordInterruptionHandlers({
    notifier,
    progress: () => progressEvent(),
    settle: async () => {
      calls.push("settlement");
      await settlement;
    },
    cleanup: async () => {
      calls.push("cleanup");
      await cleanup;
    },
    signalSource: source,
    notificationTimeoutMs: 1_000,
    terminate: () => {
      calls.push("terminate");
      resolveTermination?.();
    },
  });

  source.emit("SIGTERM");
  await new Promise<void>((resolve) => setImmediate(resolve));
  assert.deepEqual(calls.sort(), ["cleanup", "notification", "settlement"]);
  resolveSettlement?.();
  await new Promise<void>((resolve) => setImmediate(resolve));
  assert.equal(calls.includes("terminate"), false);
  resolveCleanup?.();
  await terminated;
  assert.equal(calls.at(-1), "terminate");
});

test("interruption still terminates when progress collection fails", async () => {
  const source = new FakeSignalSource();
  let terminated = false;
  const notifier: DiscordNotifier = {
    runStarted: async () => undefined,
    runProgress: async () => undefined,
    runCompleted: async () => undefined,
    runFailed: async () => undefined,
    runInterrupted: async () => undefined,
  };
  const completion = new Promise<void>((resolve) => {
    registerDiscordInterruptionHandlers({
      notifier,
      progress: () => {
        throw new Error("progress unavailable");
      },
      signalSource: source,
      terminate: () => {
        terminated = true;
        resolve();
      },
    });
  });
  source.emit("SIGTERM");
  await completion;
  assert.equal(terminated, true);
});

test("interruption termination has a hard deadline when notification hangs", async () => {
  const source = new FakeSignalSource();
  let resolveTermination: (() => void) | undefined;
  const terminated = new Promise<void>((resolve) => {
    resolveTermination = resolve;
  });
  const notifier: DiscordNotifier = {
    runStarted: async () => undefined,
    runProgress: async () => undefined,
    runCompleted: async () => undefined,
    runFailed: async () => undefined,
    runInterrupted: async () => new Promise<void>(() => undefined),
  };
  registerDiscordInterruptionHandlers({
    notifier,
    progress: () => progressEvent(),
    signalSource: source,
    notificationTimeoutMs: 1,
    terminate: () => resolveTermination?.(),
  });

  source.emit("SIGINT");
  assert.equal(source.listeners.size, 0);
  await terminated;
});
