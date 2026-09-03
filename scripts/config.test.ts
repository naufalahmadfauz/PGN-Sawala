import assert from "node:assert/strict";
import {
  copyFile,
  mkdir,
  mkdtemp,
  readFile,
  rm,
  writeFile,
} from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import test, { type TestContext } from "node:test";
import { loadConfig } from "../src/config";
import { createGoogleDriveEvidencePublisher } from "../src/evidence/google-drive";
import {
  resolveGoogleServiceAccount,
  safeGoogleCredentialError,
} from "../src/evidence/google-service-account";
import { REPOSITORY_ROOT } from "../src/environment";
import { createFreshPgnWorkbook } from "../src/excel/fresh-workbook";

const sourceWorkbookPath = path.resolve(
  "data/PGN AI Assistant - Knowledge Base Testing Report - User Inputs.xlsx",
);
const privateKey = "not-a-real-private-key";

function serviceAccount(email: string): string {
  return JSON.stringify({
    type: "service_account",
    project_id: "configuration-test",
    client_email: email,
    private_key: privateKey,
  });
}

async function temporaryRepository(
  context: TestContext,
): Promise<string> {
  const repositoryRoot = await mkdtemp(path.join(tmpdir(), "pgn-config-"));
  context.after(() => rm(repositoryRoot, { recursive: true, force: true }));
  await Promise.all([
    mkdir(path.join(repositoryRoot, "data"), { recursive: true }),
    mkdir(path.join(repositoryRoot, "reports"), { recursive: true }),
    mkdir(path.join(repositoryRoot, ".secrets"), { recursive: true }),
  ]);
  return repositoryRoot;
}

test("repository .env loads a relative service-account file", async (context) => {
  const repositoryRoot = await temporaryRepository(context);
  await writeFile(
    path.join(repositoryRoot, ".secrets", "google-service-account.json"),
    serviceAccount("local-file@example.iam.gserviceaccount.com"),
  );
  await writeFile(
    path.join(repositoryRoot, ".env"),
    [
      "GOOGLE_DRIVE_EVIDENCE_ENABLED=true",
      "GOOGLE_DRIVE_EVIDENCE_PARENT_FOLDER=https://drive.google.com/drive/folders/1LocalFolder12345",
      "GOOGLE_SERVICE_ACCOUNT_FILE=.secrets/google-service-account.json",
    ].join("\n"),
  );

  const config = loadConfig({ repositoryRoot, environment: {} });
  const resolved = resolveGoogleServiceAccount(config.googleServiceAccount);
  assert.equal(config.environmentFileLoaded, true);
  assert.equal(config.googleDriveConfigurationSource, ".env");
  assert.equal(
    config.googleServiceAccount?.source,
    "GOOGLE_SERVICE_ACCOUNT_FILE",
  );
  assert.equal(config.googleServiceAccount?.environmentSource, ".env");
  assert.equal(
    config.googleServiceAccount?.filePath,
    path.join(repositoryRoot, ".secrets", "google-service-account.json"),
  );
  assert.equal(
    resolved.credentials.client_email,
    "local-file@example.iam.gserviceaccount.com",
  );
});

test("missing .env continues with process environment credentials", async (context) => {
  const repositoryRoot = await temporaryRepository(context);
  const config = loadConfig({
    repositoryRoot,
    environment: {
      GOOGLE_DRIVE_EVIDENCE_ENABLED: "true",
      GOOGLE_DRIVE_EVIDENCE_PARENT_FOLDER: "1ProcessFolder123",
      GOOGLE_SERVICE_ACCOUNT_JSON: serviceAccount(
        "process-env@example.iam.gserviceaccount.com",
      ),
    },
  });

  assert.equal(config.environmentFileLoaded, false);
  assert.equal(config.googleDriveConfigurationSource, "process environment");
  assert.equal(
    resolveGoogleServiceAccount(config.googleServiceAccount).credentials
      .client_email,
    "process-env@example.iam.gserviceaccount.com",
  );
});

test("process environment wins over .env for the same key", async (context) => {
  const repositoryRoot = await temporaryRepository(context);
  await Promise.all([
    writeFile(
      path.join(repositoryRoot, ".secrets", "dotenv-account.json"),
      serviceAccount("dotenv@example.iam.gserviceaccount.com"),
    ),
    writeFile(
      path.join(repositoryRoot, ".secrets", "process-account.json"),
      serviceAccount("process@example.iam.gserviceaccount.com"),
    ),
    writeFile(
      path.join(repositoryRoot, ".env"),
      [
        "GOOGLE_DRIVE_EVIDENCE_ENABLED=true",
        "GOOGLE_DRIVE_EVIDENCE_PARENT_FOLDER=1DotenvFolder123",
        "GOOGLE_SERVICE_ACCOUNT_FILE=.secrets/dotenv-account.json",
      ].join("\n"),
    ),
  ]);

  const config = loadConfig({
    repositoryRoot,
    environment: {
      GOOGLE_SERVICE_ACCOUNT_FILE: ".secrets/process-account.json",
    },
  });
  assert.equal(config.googleDriveConfigurationSource, "process environment + .env");
  assert.equal(
    config.googleServiceAccount?.environmentSource,
    "process environment",
  );
  assert.equal(
    resolveGoogleServiceAccount(config.googleServiceAccount).credentials
      .client_email,
    "process@example.iam.gserviceaccount.com",
  );
});

test("default configuration resolves the repository root independently of cwd", async (context) => {
  const arbitraryWorkingDirectory = await temporaryRepository(context);
  const originalWorkingDirectory = process.cwd();
  const environment: NodeJS.ProcessEnv = {
    PGN_WHATSAPP_PHONE: "",
    PGN_WHATSAPP_CHAT: "",
    WHATSAPP_LOGIN_TIMEOUT_MS: "",
    WHATSAPP_AUTH_TIMEOUT_MS: "",
    WHATSAPP_RESPONSE_TIMEOUT_MS: "",
    WHATSAPP_RESPONSE_IDLE_MS: "",
    WHATSAPP_BETWEEN_TESTS_MS: "",
    WHATSAPP_HEADLESS: "",
    WHATSAPP_BROWSER_CHANNEL: "",
    WHATSAPP_PROFILE_DIR: "",
    PGN_RESET_COMMAND: "",
    PGN_RESET_CONFIRMATION: "",
    PGN_RESET_TIMEOUT_MS: "",
    POST_RESET_QUIET_MS: "",
    GOOGLE_DRIVE_EVIDENCE_ENABLED: "false",
    GOOGLE_DRIVE_EVIDENCE_PARENT_FOLDER: "",
    GOOGLE_DRIVE_EVIDENCE_FOLDER_PREFIX: "",
    GOOGLE_DRIVE_RETEST_FOLDER_PREFIX: "",
    GOOGLE_SERVICE_ACCOUNT_JSON: "",
    GOOGLE_SERVICE_ACCOUNT_JSON_BASE64: "",
    GOOGLE_SERVICE_ACCOUNT_FILE:
      ".secrets/default-root-service-account.json",
    LEGACY_EVIDENCE_CROP_LEFT: "",
    PGN_TEST_DATA_FILE: "",
    PGN_TEST_REPORT_FILE: "",
    PGN_SOURCE_WORKBOOK: "",
    PGN_EXECUTED_WORKBOOK: "",
  };
  try {
    process.chdir(arbitraryWorkingDirectory);
    const config = loadConfig({ environment });
    assert.equal(config.projectRoot, REPOSITORY_ROOT);
    assert.equal(
      config.googleServiceAccount?.filePath,
      path.join(
        REPOSITORY_ROOT,
        ".secrets",
        "default-root-service-account.json",
      ),
    );
  } finally {
    process.chdir(originalWorkingDirectory);
  }
});

test("credential file errors are clear and do not expose contents", async (context) => {
  const repositoryRoot = await temporaryRepository(context);
  const missing = loadConfig({
    repositoryRoot,
    environment: {
      GOOGLE_SERVICE_ACCOUNT_FILE: ".secrets/missing-service-account.json",
    },
  });
  assert.throws(
    () => resolveGoogleServiceAccount(missing.googleServiceAccount),
    /GOOGLE_SERVICE_ACCOUNT_FILE was not found:.*missing-service-account\.json/,
  );

  const malformedPath = path.join(
    repositoryRoot,
    ".secrets",
    "malformed-service-account.json",
  );
  await writeFile(malformedPath, `{ "private_key": "${privateKey}"`);
  const malformed = loadConfig({
    repositoryRoot,
    environment: {
      GOOGLE_SERVICE_ACCOUNT_FILE:
        ".secrets/malformed-service-account.json",
    },
  });
  let error: unknown;
  try {
    resolveGoogleServiceAccount(malformed.googleServiceAccount);
  } catch (caught) {
    error = caught;
  }
  assert(error instanceof Error);
  const safeMessage = safeGoogleCredentialError(error);
  assert.match(safeMessage, /GOOGLE_SERVICE_ACCOUNT_FILE is not valid JSON/);
  assert(!safeMessage.includes("not-a-real-private-key"));

  await writeFile(
    malformedPath,
    JSON.stringify({
      type: "authorized_user",
      client_email: "wrong-type@example.com",
      private_key: privateKey,
    }),
  );
  assert.throws(
    () => resolveGoogleServiceAccount(malformed.googleServiceAccount),
    /must have type "service_account"/,
  );
  await writeFile(
    malformedPath,
    JSON.stringify({
      type: "service_account",
      client_email: "missing-key@example.com",
    }),
  );
  assert.throws(
    () => resolveGoogleServiceAccount(malformed.googleServiceAccount),
    /must contain private_key/,
  );
});

test("raw JSON has priority over base64 and file credentials", async (context) => {
  const repositoryRoot = await temporaryRepository(context);
  await writeFile(
    path.join(repositoryRoot, ".secrets", "fallback-service-account.json"),
    serviceAccount("file@example.iam.gserviceaccount.com"),
  );
  const raw = serviceAccount("raw@example.iam.gserviceaccount.com");
  const base64 = Buffer.from(
    serviceAccount("base64@example.iam.gserviceaccount.com"),
  ).toString("base64");
  const config = loadConfig({
    repositoryRoot,
    environment: {
      GOOGLE_DRIVE_EVIDENCE_ENABLED: "true",
      GOOGLE_DRIVE_EVIDENCE_PARENT_FOLDER: "1CredentialPriority",
      GOOGLE_SERVICE_ACCOUNT_JSON: raw,
      GOOGLE_SERVICE_ACCOUNT_JSON_BASE64: base64,
      GOOGLE_SERVICE_ACCOUNT_FILE:
        ".secrets/fallback-service-account.json",
    },
  });

  assert.equal(
    config.googleServiceAccount?.source,
    "GOOGLE_SERVICE_ACCOUNT_JSON",
  );
  assert.deepEqual(config.googleServiceAccount?.configuredSources, [
    "GOOGLE_SERVICE_ACCOUNT_JSON",
    "GOOGLE_SERVICE_ACCOUNT_JSON_BASE64",
    "GOOGLE_SERVICE_ACCOUNT_FILE",
  ]);
  const publisher = createGoogleDriveEvidencePublisher(config);
  assert.equal(
    publisher.serviceAccountEmail,
    "raw@example.iam.gserviceaccount.com",
  );
  const unsafeError = new Error(`Authentication failed: ${raw} ya29.fake 1//fake`);
  const safeMessage = safeGoogleCredentialError(unsafeError, raw);
  assert(!safeMessage.includes(raw));
  assert(!safeMessage.includes("ya29.fake"));
  assert(!safeMessage.includes("1//fake"));
});

test("base64 JSON credentials decode when higher-priority JSON is absent", async (context) => {
  const repositoryRoot = await temporaryRepository(context);
  const config = loadConfig({
    repositoryRoot,
    environment: {
      GOOGLE_SERVICE_ACCOUNT_JSON_BASE64: Buffer.from(
        serviceAccount("base64-only@example.iam.gserviceaccount.com"),
      ).toString("base64"),
      GOOGLE_SERVICE_ACCOUNT_FILE: ".secrets/lower-priority.json",
    },
  });
  const resolved = resolveGoogleServiceAccount(config.googleServiceAccount);
  assert.equal(resolved.source, "GOOGLE_SERVICE_ACCOUNT_JSON_BASE64");
  assert.equal(
    resolved.credentials.client_email,
    "base64-only@example.iam.gserviceaccount.com",
  );
  assert.deepEqual(config.googleServiceAccount?.configuredSources, [
    "GOOGLE_SERVICE_ACCOUNT_JSON_BASE64",
    "GOOGLE_SERVICE_ACCOUNT_FILE",
  ]);
  const invalid = loadConfig({
    repositoryRoot,
    environment: { GOOGLE_SERVICE_ACCOUNT_JSON_BASE64: "not base64!" },
  });
  assert.throws(
    () => resolveGoogleServiceAccount(invalid.googleServiceAccount),
    /is not valid base64/,
  );
});

test("fresh-run preparation preserves local environment and credentials", async (context) => {
  const repositoryRoot = await temporaryRepository(context);
  const envPath = path.join(repositoryRoot, ".env");
  const credentialPath = path.join(
    repositoryRoot,
    ".secrets",
    "google-service-account.json",
  );
  await writeFile(
    envPath,
    [
      "GOOGLE_DRIVE_EVIDENCE_ENABLED=true",
      "GOOGLE_DRIVE_EVIDENCE_PARENT_FOLDER=1FreshFolder12345",
      "GOOGLE_SERVICE_ACCOUNT_FILE=.secrets/google-service-account.json",
      "PGN_SOURCE_WORKBOOK=data/source.xlsx",
      "PGN_EXECUTED_WORKBOOK=reports/executed.xlsx",
    ].join("\n"),
  );
  await writeFile(
    credentialPath,
    serviceAccount("fresh@example.iam.gserviceaccount.com"),
  );
  const config = loadConfig({ repositoryRoot, environment: {} });
  await copyFile(sourceWorkbookPath, config.pgnSourceWorkbookPath);
  await copyFile(sourceWorkbookPath, config.pgnExecutedWorkbookPath);
  const envBefore = await readFile(envPath);
  const credentialsBefore = await readFile(credentialPath);

  await createFreshPgnWorkbook(config, new Date("2026-09-03T00:00:00Z"));

  assert((await readFile(envPath)).equals(envBefore));
  assert((await readFile(credentialPath)).equals(credentialsBefore));
});
