export class RestApiError extends Error {
  constructor(message: string, readonly global = true, readonly status?: number) {
    super(message);
    this.name = "RestApiError";
  }
}

export function redactRestText(text: string, ...secrets: Array<string | undefined>): string {
  for (const secret of secrets.filter((value): value is string => Boolean(value))) {
    text = text.split(secret).join("[REDACTED]");
    text = text.split(encodeURIComponent(secret)).join("[REDACTED]");
  }
  return text
    .replace(/\b(?:Bearer\s+)?eyJ[A-Za-z0-9_-]+\.[A-Za-z0-9_-]+\.[A-Za-z0-9_-]+/g, "[REDACTED]")
    .replace(/((?:authorization|x-lp-on-behalf|lp-on-behalf|client_secret|clientSecret|access_token|refresh_token|AppJWT|ConsumerJWS)["']?\s*[:=]\s*)[^\r\n,}]+/gi, "$1[REDACTED]")
    .replace(/-----BEGIN [^-]*PRIVATE KEY-----[\s\S]*?-----END [^-]*PRIVATE KEY-----/g, "[REDACTED]");
}

export function safeRestError(error: unknown, ...secrets: Array<string | undefined>): string {
  return redactRestText(error instanceof Error ? error.message : "LivePerson REST operation failed", ...secrets).slice(0, 1000);
}
