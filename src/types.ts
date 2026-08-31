export type MessageDirection = "incoming" | "outgoing";
export type MessageDeliveryStatus =
  | "pending"
  | "sent"
  | "delivered"
  | "failed";

export interface WhatsAppMessage {
  id: string;
  direction: MessageDirection;
  text: string;
  domIndex: number;
  observedAt: Date;
  deliveryStatus?: MessageDeliveryStatus;
}

export interface SentMessage {
  sentAt: Date;
  messageId: string;
  renderedText: string;
}

export interface MessageSnapshot {
  ids: Set<string>;
  messageCount: number;
  messages: WhatsAppMessage[];
}

export interface ResponseCapture {
  messages: WhatsAppMessage[];
  combinedResponse: string;
  sentAt: Date;
  firstResponseAt?: Date;
  completedAt: Date;
  firstResponseMs?: number;
  totalResponseMs: number;
  timedOut: boolean;
}

export interface TestCase {
  testId: string;
  category: string;
  userInput: string;
  expectedBehaviour: string;
  scenarioId?: string;
  sourceRow?: number;
}

export type TestStatus = "CAPTURED" | "TIMEOUT" | "ERROR";

export interface TestResult {
  runId: string;
  testCase: TestCase;
  botResponse: string;
  firstResponseMs?: number;
  totalResponseMs?: number;
  status: TestStatus;
  startedAt: Date;
  completedAt: Date;
  error?: string;
  evidencePath?: string;
}

export interface TranscriptEntry {
  testId: string;
  sequence: number;
  role: "USER" | "BOT";
  message: string;
  timestamp: Date;
}
