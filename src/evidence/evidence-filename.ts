import { createHash } from "node:crypto";

export function safePathSegment(value: string): string {
  const sanitized = value.replace(/[^A-Za-z0-9_-]+/g, "-").replace(/^-+|-+$/g, "").slice(0, 100);
  if (!sanitized) throw new Error("Evidence identifier cannot be converted to a safe filename");
  if (sanitized === value) return sanitized;
  const suffix = createHash("sha256").update(value).digest("hex").slice(0, 8);
  return `${sanitized.slice(0, 90)}-${suffix}`;
}
export function evidenceFileName(testCaseId: string, turnNumber: number): string {
  return `${safePathSegment(testCaseId)}-turn-${turnNumber}.png`;
}
