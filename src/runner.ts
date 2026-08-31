import path from "node:path";
import { pathToFileURL } from "node:url";
import { loadConfig, requireTarget, type AppConfig } from "./config";
import { loadTestCases } from "./excel/loader";
import { writeResultsWorkbook } from "./excel/writer";
import type {
  TestCase,
  TestResult,
  TranscriptEntry,
  SentMessage,
} from "./types";
import { WhatsAppClient } from "./whatsapp/client";

function createRunId(): string {
  return new Date().toISOString().replace(/[-:.]/g, "").replace("Z", "Z");
}

function safeFileName(value: string): string {
  return value.replace(/[^a-zA-Z0-9_-]+/g, "-").slice(0, 80) || "test";
}

function relativeToProject(config: AppConfig, absolutePath: string): string {
  return path.relative(config.projectRoot, absolutePath).replaceAll(path.sep, "/");
}

async function saveDebugSafely(
  client: WhatsAppClient,
  name: string,
): Promise<string | undefined> {
  try {
    return (await client.saveDebugArtifacts(name)).screenshotPath;
  } catch (error) {
    const detail = error instanceof Error ? error.message : String(error);
    console.error(`[Debug] Could not save ${name}: ${detail}`);
    return undefined;
  }
}

export async function runTestCases(
  testCases: TestCase[],
  config = loadConfig(),
): Promise<TestResult[]> {
  if (testCases.length === 0) {
    throw new Error("At least one test case is required");
  }
  const testIds = new Set<string>();
  for (const testCase of testCases) {
    if (testIds.has(testCase.testId)) {
      throw new Error(`Duplicate Test ID: ${testCase.testId}`);
    }
    testIds.add(testCase.testId);
  }

  const target = requireTarget(config);
  const runId = createRunId();
  const client = new WhatsAppClient(config);
  const results: TestResult[] = [];
  const transcript: TranscriptEntry[] = [];

  try {
    await client.open();
    await client.ensureAuthenticated({ allowQrLogin: false });
    await client.openChat(target);

    for (let index = 0; index < testCases.length; index += 1) {
      const testCase = testCases[index];
      const startedAt = new Date();
      let sentMessage: SentMessage | undefined;
      const artifactKey = `${runId}-${String(index + 1).padStart(4, "0")}-${safeFileName(testCase.testId)}`;
      console.log(`[Test] ${testCase.testId}: Sending: ${testCase.userInput}`);

      try {
        const baseline = await client.captureMessageState();
        sentMessage = await client.sendMessage(testCase.userInput, baseline);
        const response = await client.waitForBotResponse(baseline, sentMessage);
        const evidenceAbsolutePath = path.join(
          config.evidenceDir,
          `${artifactKey}.png`,
        );
        await client.captureScreenshot(evidenceAbsolutePath);
        const evidencePath = relativeToProject(config, evidenceAbsolutePath);
        const status = response.timedOut ? "TIMEOUT" : "CAPTURED";
        const error = response.timedOut
          ? `Bot response did not complete within ${config.responseTimeoutMs} ms`
          : undefined;

        if (response.combinedResponse) {
          console.log(`[Bot] ${response.combinedResponse}`);
        }
        if (response.firstResponseMs !== undefined) {
          console.log(`[Test] First response: ${response.firstResponseMs} ms`);
        }
        console.log(`[Test] Total response: ${response.totalResponseMs} ms`);
        console.log(`[Test] Bot messages: ${response.messages.length}`);

        results.push({
          runId,
          testCase,
          botResponse: response.combinedResponse,
          firstResponseMs: response.firstResponseMs,
          totalResponseMs: response.totalResponseMs,
          status,
          startedAt,
          completedAt: response.completedAt,
          error,
          evidencePath,
        });
        transcript.push({
          testId: testCase.testId,
          sequence: 1,
          role: "USER",
          message: testCase.userInput,
          timestamp: sentMessage.sentAt,
        });
        response.messages.forEach((message, messageIndex) => {
          transcript.push({
            testId: testCase.testId,
            sequence: messageIndex + 2,
            role: "BOT",
            message: message.text,
            timestamp: message.observedAt,
          });
        });

        if (response.timedOut) {
          await saveDebugSafely(client, `response-timeout-${artifactKey}`);
          console.log(`[Test] ${testCase.testId}: TIMEOUT`);
        } else {
          console.log(`[Test] ${testCase.testId}: COMPLETE`);
        }
      } catch (error) {
        const completedAt = new Date();
        const detail = error instanceof Error ? error.message : String(error);
        const debugScreenshotPath = await saveDebugSafely(
          client,
          `test-failure-${artifactKey}`,
        );
        const evidencePath = debugScreenshotPath
          ? relativeToProject(config, debugScreenshotPath)
          : undefined;
        results.push({
          runId,
          testCase,
          botResponse: "",
          status: "ERROR",
          startedAt,
          completedAt,
          error: detail,
          evidencePath,
        });
        if (sentMessage) {
          transcript.push({
            testId: testCase.testId,
            sequence: 1,
            role: "USER",
            message: testCase.userInput,
            timestamp: sentMessage.sentAt,
          });
        }
        console.error(`[Test] ${testCase.testId}: ERROR: ${detail}`);
      }

      await writeResultsWorkbook(config.reportFilePath, results, transcript);
      if (index < testCases.length - 1 && config.betweenTestsMs > 0) {
        console.log(
          `[Test] Waiting ${config.betweenTestsMs} ms before the next case`,
        );
        await new Promise((resolve) =>
          setTimeout(resolve, config.betweenTestsMs),
        );
      }
    }
  } finally {
    await client.close();
  }

  console.log(
    `[Report] ${relativeToProject(config, config.reportFilePath)} (${results.length} result(s))`,
  );
  return results;
}

async function main(): Promise<void> {
  const config = loadConfig();
  const testCases = await loadTestCases(config.dataFilePath);
  console.log(`[Test] Loaded ${testCases.length} test case(s)`);
  await runTestCases(testCases, config);
}

if (import.meta.url === pathToFileURL(process.argv[1]).href) {
  main().catch((error) => {
    console.error(error instanceof Error ? error.message : error);
    process.exitCode = 1;
  });
}
