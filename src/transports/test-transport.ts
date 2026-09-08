import type { ExecutedTurn, PgnTestScenario, PgnTestTurn } from "../excel/pgn-types";
import type { ExecutionTransport } from "../session-mode";

export interface TransportResponse extends Omit<ExecutedTurn, "turn"> {
  transport: ExecutionTransport;
}

export interface TestTransport {
  readonly type: ExecutionTransport;
  initializeRun(): Promise<void>;
  beginScenario(scenario: PgnTestScenario, index: number): Promise<void>;
  sendMessage(scenario: PgnTestScenario, turn: PgnTestTurn): Promise<TransportResponse>;
  endScenario(scenario: PgnTestScenario): Promise<void>;
  finalizeRun(): Promise<void>;
  close(): Promise<void>;
}
