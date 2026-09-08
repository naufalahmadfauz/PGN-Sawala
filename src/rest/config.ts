export type LivePersonDomain = "sentinel" | "idp" | "asyncMessagingEnt" | "messagingRestApiDomain";
export const LIVEPERSON_DOMAINS: LivePersonDomain[] = ["sentinel", "idp", "asyncMessagingEnt", "messagingRestApiDomain"];

export interface LivePersonRestConfig {
  enabled: boolean;
  accountId?: string;
  clientId?: string;
  clientSecret?: string;
  skillId?: string;
  domains: Partial<Record<LivePersonDomain, string>>;
  responseIdleMs: number;
  responseTimeoutMs: number;
  pollIntervalMs: number;
  requestTimeoutMs: number;
}

export function domainOrigin(value: string): string {
  let url: URL;
  try { url = new URL(value.includes("://") ? value : `https://${value}`); }
  catch { throw new Error("LivePerson domain override must be a valid HTTPS hostname"); }
  if (url.protocol !== "https:" || url.username || url.password || url.search || url.hash || url.pathname !== "/" || !url.hostname) {
    throw new Error("LivePerson domains must be HTTPS origins without credentials, paths, or query parameters");
  }
  return url.origin;
}

export function loadRestConfig(environment: NodeJS.ProcessEnv): LivePersonRestConfig {
  const rawEnabled = environment.LIVEPERSON_REST_ENABLED?.trim().toLowerCase();
  if (rawEnabled && !["true", "false", "1", "0"].includes(rawEnabled)) throw new Error("LIVEPERSON_REST_ENABLED must be true or false");
  const enabled = rawEnabled === "true" || rawEnabled === "1";
  const integer = (name: string, fallback: number): number => {
    const value = environment[name]?.trim();
    if (!value) return fallback;
    const parsed = Number(value);
    if (!Number.isSafeInteger(parsed) || parsed < 1) throw new Error(`${name} must be a positive integer`);
    return parsed;
  };
  const domains: LivePersonRestConfig["domains"] = {};
  for (const [service, name] of [
    ["sentinel", "LIVEPERSON_SENTINEL_DOMAIN"], ["idp", "LIVEPERSON_IDP_DOMAIN"],
    ["asyncMessagingEnt", "LIVEPERSON_ASYNC_MESSAGING_DOMAIN"], ["messagingRestApiDomain", "LIVEPERSON_MESSAGING_REST_DOMAIN"],
  ] as const) {
    const value = environment[name]?.trim();
    if (value) domains[service] = domainOrigin(value);
  }
  const config: LivePersonRestConfig = {
    enabled, accountId: environment.LIVEPERSON_ACCOUNT_ID?.trim() || undefined,
    clientId: environment.LIVEPERSON_CLIENT_ID?.trim() || undefined,
    clientSecret: environment.LIVEPERSON_CLIENT_SECRET?.trim() || undefined,
    skillId: environment.LIVEPERSON_SKILL_ID?.trim() || undefined, domains,
    responseIdleMs: integer("REST_RESPONSE_IDLE_MS", 3000),
    responseTimeoutMs: integer("REST_RESPONSE_TIMEOUT_MS", 60000),
    pollIntervalMs: integer("REST_POLL_INTERVAL_MS", 750),
    requestTimeoutMs: integer("REST_REQUEST_TIMEOUT_MS", 15000),
  };
  if (config.responseIdleMs >= config.responseTimeoutMs) throw new Error("REST_RESPONSE_IDLE_MS must be less than REST_RESPONSE_TIMEOUT_MS");
  return config;
}

export function assertRestConfig(config: LivePersonRestConfig | undefined): asserts config is LivePersonRestConfig & { accountId: string; clientId: string; clientSecret: string; skillId: string } {
  if (!config?.enabled) throw new Error("LivePerson REST testing is disabled; configure LIVEPERSON_REST_ENABLED=true explicitly");
  const missing = [
    ["LIVEPERSON_ACCOUNT_ID", config.accountId], ["LIVEPERSON_CLIENT_ID", config.clientId],
    ["LIVEPERSON_CLIENT_SECRET", config.clientSecret], ["LIVEPERSON_SKILL_ID", config.skillId],
  ].filter(([, value]) => !value?.trim()).map(([name]) => name);
  if (missing.length) throw new Error(`LivePerson REST configuration missing: ${missing.join(", ")}`);
  if (!/^[A-Za-z0-9_-]+$/.test(config.accountId!) || !/^\d+$/.test(config.skillId!)) {
    throw new Error("LivePerson account ID must be alphanumeric and skill ID must be numeric");
  }
}
