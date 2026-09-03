import { readFileSync } from "node:fs";
import type { EnvironmentValueSource } from "../environment";

export const GOOGLE_SERVICE_ACCOUNT_SOURCES = [
  "GOOGLE_SERVICE_ACCOUNT_JSON",
  "GOOGLE_SERVICE_ACCOUNT_JSON_BASE64",
  "GOOGLE_SERVICE_ACCOUNT_FILE",
] as const;

export type GoogleServiceAccountSource =
  (typeof GOOGLE_SERVICE_ACCOUNT_SOURCES)[number];

export interface GoogleServiceAccountConfiguration {
  source: GoogleServiceAccountSource;
  configuredSources: GoogleServiceAccountSource[];
  environmentSource: EnvironmentValueSource;
  value?: string;
  filePath?: string;
  fileDisplayPath?: string;
}

export interface ServiceAccountCredentials {
  type: "service_account";
  client_email: string;
  private_key: string;
  project_id?: string;
}

export interface ResolvedGoogleServiceAccount {
  source: GoogleServiceAccountSource;
  credentials: ServiceAccountCredentials;
  filePath?: string;
}

export function safeGoogleCredentialError(
  error: unknown,
  configuredSecret?: string,
): string {
  let message =
    error instanceof Error ? error.message : "Unknown Google Drive error";
  if (configuredSecret && message.includes(configuredSecret)) {
    message = message.split(configuredSecret).join("[REDACTED]");
  }
  return message
    .replace(/-----BEGIN[^]*?-----END PRIVATE KEY-----/g, "[REDACTED]")
    .replace(/ya29\.[A-Za-z0-9._-]+/g, "[REDACTED]")
    .replace(/1\/\/[A-Za-z0-9._-]+/g, "[REDACTED]")
    .replace(
      /("(?:private_key|access_token|refresh_token)"\s*:\s*")[^"]*/gi,
      "$1[REDACTED]",
    );
}

function parseServiceAccountJson(
  raw: string,
  source: GoogleServiceAccountSource,
): ServiceAccountCredentials {
  let parsed: unknown;
  try {
    parsed = JSON.parse(raw);
  } catch {
    throw new Error(`${source} is not valid JSON`);
  }
  if (!parsed || typeof parsed !== "object" || Array.isArray(parsed)) {
    throw new Error(`${source} must contain a JSON object`);
  }
  const credentials = parsed as Record<string, unknown>;
  if (credentials.type !== "service_account") {
    throw new Error(`${source} must have type "service_account"`);
  }
  if (
    typeof credentials.client_email !== "string" ||
    !credentials.client_email.includes("@")
  ) {
    throw new Error(`${source} must contain client_email`);
  }
  if (
    typeof credentials.private_key !== "string" ||
    !credentials.private_key.trim()
  ) {
    throw new Error(`${source} must contain private_key`);
  }
  return {
    type: "service_account",
    client_email: credentials.client_email.trim(),
    private_key: credentials.private_key,
    project_id:
      typeof credentials.project_id === "string"
        ? credentials.project_id
        : undefined,
  };
}

function decodeBase64(value: string): string {
  const compact = value.replace(/\s+/g, "");
  if (
    !compact ||
    compact.length % 4 === 1 ||
    !/^[A-Za-z0-9+/]*={0,2}$/.test(compact)
  ) {
    throw new Error("GOOGLE_SERVICE_ACCOUNT_JSON_BASE64 is not valid base64");
  }
  const decoded = Buffer.from(compact, "base64");
  if (
    decoded.toString("base64").replace(/=+$/, "") !==
    compact.replace(/=+$/, "")
  ) {
    throw new Error("GOOGLE_SERVICE_ACCOUNT_JSON_BASE64 is not valid base64");
  }
  return decoded.toString("utf8");
}

export function resolveGoogleServiceAccount(
  configuration: GoogleServiceAccountConfiguration | undefined,
): ResolvedGoogleServiceAccount {
  if (!configuration) {
    throw new Error(
      "Configure GOOGLE_SERVICE_ACCOUNT_JSON, GOOGLE_SERVICE_ACCOUNT_JSON_BASE64, or GOOGLE_SERVICE_ACCOUNT_FILE",
    );
  }

  let raw: string;
  if (configuration.source === "GOOGLE_SERVICE_ACCOUNT_JSON") {
    raw = configuration.value ?? "";
  } else if (
    configuration.source === "GOOGLE_SERVICE_ACCOUNT_JSON_BASE64"
  ) {
    raw = decodeBase64(configuration.value ?? "");
  } else {
    const displayPath = JSON.stringify(
      configuration.fileDisplayPath ?? "configured credential file",
    );
    if (!configuration.filePath) {
      throw new Error(`GOOGLE_SERVICE_ACCOUNT_FILE was not found: ${displayPath}`);
    }
    try {
      raw = readFileSync(configuration.filePath, "utf8");
    } catch (error) {
      const code = (error as NodeJS.ErrnoException).code;
      if (code === "ENOENT") {
        throw new Error(
          `GOOGLE_SERVICE_ACCOUNT_FILE was not found: ${displayPath}`,
        );
      }
      throw new Error(
        `GOOGLE_SERVICE_ACCOUNT_FILE is not readable: ${displayPath}`,
      );
    }
  }

  return {
    source: configuration.source,
    credentials: parseServiceAccountJson(raw, configuration.source),
    filePath: configuration.filePath,
  };
}
