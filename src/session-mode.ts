export type SessionMode = "isolated" | "continuous";
export type ExecutionTransport = "whatsapp";

export interface RunExecutionContext {
  transport: ExecutionTransport;
  sessionMode: SessionMode;
}

export const CONTINUOUS_SESSION_WARNING = [
  "Continuous Session Mode",
  "One clean initial reset is required. Session resets between scenarios and final cleanup are disabled.",
  "All scenarios share the same bot conversation; previous testcase context may influence later responses.",
  "Recommended for rapid development checks, exploratory testing, and context-stress testing.",
  "Not recommended as the only final acceptance run. Results are not independent testcase outcomes.",
].join("\n");

export const CONTINUOUS_RECOVERY_WARNING =
  "Continuous Session run interrupted. The original shared conversation context may no longer be reliable. Mid-stream resume is disabled; restart the entire continuous run from the beginning or abandon it.";

export function readSessionMode(value: unknown): SessionMode {
  if (value === undefined) return "isolated"; // V1 persisted runs were isolated.
  if (value === "isolated" || value === "continuous") return value;
  throw new Error("Session mode must be isolated or continuous");
}

export function readExecutionTransport(value: unknown): ExecutionTransport {
  if (value === undefined || value === "whatsapp") return "whatsapp";
  throw new Error("Only the WhatsApp execution transport is supported");
}

export function sessionModeLabel(mode: SessionMode): string {
  return mode === "continuous" ? "Continuous" : "Isolated";
}

export function shouldResetBeforeScenario(mode: SessionMode, scenarioIndex: number): boolean {
  return mode === "isolated" || scenarioIndex === 0;
}
