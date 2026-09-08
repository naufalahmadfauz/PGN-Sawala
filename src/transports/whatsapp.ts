import path from "node:path";
import type { Workbook } from "exceljs";
import { requireTarget, type AppConfig } from "../config";
import type { EvidenceDrivePublisher } from "../evidence/google-drive";
import { safeGoogleCredentialError } from "../evidence/google-service-account";
import { evidenceFileName } from "../evidence/evidence-filename";
import { appendPostResetDrainTranscript, appendSessionResetTranscript, saveExecutedPgnWorkbook } from "../excel/pgn-workbook-writer";
import type { PgnTestScenario, PgnTestTurn } from "../excel/pgn-types";
import { shouldResetBeforeScenario, type SessionMode } from "../session-mode";
import type { BotSessionResetAttempt, SentMessage } from "../types";
import { WhatsAppClient } from "../whatsapp/client";
import { BotSessionResetError, resetBotSession, waitForPostResetQuiet } from "../whatsapp/session-reset";
import type { TestTransport, TransportResponse } from "./test-transport";

export interface RunEvidenceContext { publisher: EvidenceDrivePublisher; folderId: string }
function safeFileName(value: string): string { return value.replace(/[^a-zA-Z0-9_-]+/g, "-").slice(0, 80) || "test"; }

export class WhatsAppTransport implements TestTransport {
  readonly type = "whatsapp";
  readonly #client: WhatsAppClient;
  constructor(
    private readonly config: AppConfig,
    private readonly runId: string,
    private readonly sessionMode: SessionMode,
    private readonly workbook: Workbook,
    private readonly finalScenario: PgnTestScenario,
    private readonly onReset: () => Promise<void>,
    private readonly checkInterrupted: () => void,
    private readonly driveEvidence?: RunEvidenceContext,
  ) { this.#client = new WhatsAppClient(config, { handleProcessSignals: false }); }

  #relative(file: string): string { return path.relative(this.config.projectRoot, file).replaceAll(path.sep, "/"); }

  async initializeRun(): Promise<void> {
    await this.#client.open();
    this.checkInterrupted();
    await this.#client.ensureAuthenticated({ allowQrLogin: false });
    this.checkInterrupted();
    await this.#client.openChat(requireTarget(this.config));
    this.checkInterrupted();
  }

  #relativeReset(attempt: BotSessionResetAttempt): void {
    if (attempt.evidencePath) attempt.evidencePath = this.#relative(attempt.evidencePath);
    if (attempt.diagnosticsPath) attempt.diagnosticsPath = this.#relative(attempt.diagnosticsPath);
  }

  async #reset(scenario: PgnTestScenario, finalCleanup = false): Promise<void> {
    await this.onReset();
    this.checkInterrupted();
    try {
      const attempt = await resetBotSession(this.#client, {
        command: this.config.resetCommand, confirmation: this.config.resetConfirmation,
        timeoutMs: this.config.resetTimeoutMs,
        failureArtifactName: `reset-failure-${safeFileName(scenario.testCaseId)}-${this.runId}`,
      });
      const drain = await waitForPostResetQuiet(this.#client, {
        baselineMessages: attempt.messageStateAtCompletion ?? attempt.responseMessages,
        quietMs: this.config.postResetQuietMs,
      });
      this.#relativeReset(attempt);
      appendSessionResetTranscript(this.workbook, this.runId, scenario, attempt);
      appendPostResetDrainTranscript(this.workbook, this.runId, scenario, drain);
      await saveExecutedPgnWorkbook(this.workbook, this.config.pgnExecutedWorkbookPath);
    } catch (error) {
      if (!(error instanceof BotSessionResetError)) throw error;
      this.#relativeReset(error.attempt);
      appendSessionResetTranscript(this.workbook, this.runId, scenario, error.attempt);
      await saveExecutedPgnWorkbook(this.workbook, this.config.pgnExecutedWorkbookPath);
      if (error.attempt.evidencePath) console.error(`[Session] Debug screenshot: ${error.attempt.evidencePath}`);
      if (error.attempt.diagnosticsPath) console.error(`[Session] Diagnostics: ${error.attempt.diagnosticsPath}`);
      console.error("[Session] ABORTING TEST RUN");
      console.error(finalCleanup
        ? "[Session] Reason: Unable to confirm final PGN bot session cleanup."
        : "[Session] Reason: Unable to confirm clean PGN bot session before next scenario.");
      console.error("[Session] Completed test results have been saved.");
      if (!finalCleanup) console.error("[Session] Remaining scenarios were NOT executed.");
      throw new Error(finalCleanup ? "Unable to confirm final PGN bot session cleanup." : "Unable to confirm clean PGN bot session before next scenario.", { cause: error });
    }
  }

  async beginScenario(scenario: PgnTestScenario, index: number): Promise<void> {
    if (shouldResetBeforeScenario(this.sessionMode, index)) {
      if (this.sessionMode === "continuous") console.log("[Session] Preparing one clean initial session for the continuous run");
      await this.#reset(scenario);
    }
  }

  async #failureEvidence(name: string): Promise<string | undefined> {
    try { return this.#relative((await this.#client.saveDebugArtifacts(name)).screenshotPath); }
    catch (error) {
      console.error(`[Debug] Could not save failure evidence: ${error instanceof Error ? error.message : String(error)}`);
      return undefined;
    }
  }

  async sendMessage(scenario: PgnTestScenario, turn: PgnTestTurn): Promise<TransportResponse> {
    const artifactKey = `${this.runId}-${safeFileName(scenario.testCaseId)}-turn-${turn.turnNumber}`;
    let sentMessage: SentMessage | undefined;
    try {
      const baseline = await this.#client.captureMessageState();
      sentMessage = await this.#client.sendMessage(turn.userInput, baseline);
      let response;
      try { response = await this.#client.waitForBotResponse(baseline, sentMessage, `${scenario.testCaseId} Turn ${turn.turnNumber}`); }
      catch (error) {
        return {
          transport: "whatsapp", technicalStatus: "CHAT_ERROR", sentAt: sentMessage.sentAt,
          completedAt: new Date(), botMessages: [], combinedResponse: "",
          error: error instanceof Error ? error.message : String(error),
          evidencePath: await this.#failureEvidence(`chat-error-${artifactKey}`), evidenceStatus: "EVIDENCE_CAPTURE_ERROR",
        };
      }
      const absolutePath = path.join(this.config.evidenceDir, `${artifactKey}.png`);
      let evidencePath: string | undefined;
      let evidenceUrl: string | undefined;
      let evidenceStatus: TransportResponse["evidenceStatus"];
      let evidenceDriveFileId: string | undefined;
      let evidenceDriveFileName: string | undefined;
      try {
        await this.#client.captureScreenshot(absolutePath);
        evidencePath = this.#relative(absolutePath);
        evidenceStatus = this.driveEvidence ? "EVIDENCE_PENDING" : "EVIDENCE_LOCAL_ONLY";
      } catch (error) {
        evidenceStatus = "EVIDENCE_CAPTURE_ERROR";
        console.error(`[Evidence] Screenshot failed: ${error instanceof Error ? error.message : String(error)}`);
      }
      if (this.driveEvidence && evidencePath) {
        try {
          const uploaded = await this.driveEvidence.publisher.uploadPng({ folderId: this.driveEvidence.folderId, localPath: absolutePath, fileName: evidenceFileName(scenario.testCaseId, turn.turnNumber) });
          evidenceUrl = uploaded.webViewLink;
          evidenceDriveFileId = uploaded.id;
          evidenceDriveFileName = uploaded.name;
          evidenceStatus = "EVIDENCE_SYNCED";
          console.log(`[Evidence] Uploaded ${uploaded.name}`);
        } catch (error) {
          evidenceStatus = "EVIDENCE_UPLOAD_ERROR";
          console.error(`[Evidence] EVIDENCE_UPLOAD_ERROR: ${safeGoogleCredentialError(error, this.config.googleServiceAccount?.value)}`);
        }
      }
      if (response.timedOut) evidencePath ??= await this.#failureEvidence(`response-timeout-${artifactKey}`);
      return {
        transport: "whatsapp", technicalStatus: response.timedOut ? "TIMEOUT" : "CAPTURED",
        sentAt: sentMessage.sentAt, completedAt: response.completedAt,
        botMessages: response.messages.map((message, index) => ({ sequence: index + 1, message: message.text, timestamp: message.observedAt })),
        combinedResponse: response.combinedResponse, firstResponseMs: response.firstResponseMs, totalResponseMs: response.totalResponseMs,
        error: response.timedOut ? `TIMEOUT after ${this.config.responseTimeoutMs} ms` : undefined,
        evidencePath, evidenceUrl, evidenceStatus, evidenceDriveFileId, evidenceDriveFileName,
      };
    } catch (error) {
      return {
        transport: "whatsapp", technicalStatus: "SEND_ERROR", sentAt: sentMessage?.sentAt,
        completedAt: new Date(), botMessages: [], combinedResponse: "",
        error: error instanceof Error ? error.message : String(error),
        evidencePath: await this.#failureEvidence(`send-error-${artifactKey}`), evidenceStatus: "EVIDENCE_CAPTURE_ERROR",
      };
    }
  }

  async endScenario(_scenario: PgnTestScenario): Promise<void> {}
  async finalizeRun(): Promise<void> {
    if (this.sessionMode === "isolated") await this.#reset(this.finalScenario, true);
  }
  async close(): Promise<void> { await this.#client.close(); }
}
