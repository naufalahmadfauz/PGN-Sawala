import { createReadStream } from "node:fs";
import { google, type drive_v3 } from "googleapis";
import type { AppConfig } from "../config";

const DRIVE_SCOPE = "https://www.googleapis.com/auth/drive";
const FOLDER_MIME_TYPE = "application/vnd.google-apps.folder";

export interface DriveEvidenceItem {
  id: string;
  name: string;
  webViewLink: string;
  reused: boolean;
}

export interface EvidenceDrivePublisher {
  readonly parentFolderId: string;
  readonly folderPrefix: string;
  readonly serviceAccountEmail?: string;
  validateParentFolder(): Promise<DriveEvidenceItem>;
  ensureRunFolder(
    runId: string,
    existingFolderId?: string,
  ): Promise<DriveEvidenceItem>;
  uploadPng(options: {
    folderId: string;
    localPath: string;
    fileName: string;
    existingFileId?: string;
  }): Promise<DriveEvidenceItem>;
}

interface ServiceAccountCredentials {
  client_email: string;
  private_key: string;
  project_id?: string;
}

function parseServiceAccountJson(raw: string): ServiceAccountCredentials {
  let parsed: unknown;
  try {
    parsed = JSON.parse(raw);
  } catch {
    throw new Error("GOOGLE_SERVICE_ACCOUNT_JSON is not valid JSON");
  }
  if (!parsed || typeof parsed !== "object") {
    throw new Error("GOOGLE_SERVICE_ACCOUNT_JSON must contain a JSON object");
  }
  const credentials = parsed as Record<string, unknown>;
  if (
    typeof credentials.client_email !== "string" ||
    !credentials.client_email.includes("@") ||
    typeof credentials.private_key !== "string" ||
    !credentials.private_key.includes("PRIVATE KEY")
  ) {
    throw new Error(
      "GOOGLE_SERVICE_ACCOUNT_JSON must contain client_email and private_key",
    );
  }
  return {
    client_email: credentials.client_email,
    private_key: credentials.private_key,
    project_id:
      typeof credentials.project_id === "string"
        ? credentials.project_id
        : undefined,
  };
}

function driveQueryValue(value: string): string {
  return value.replace(/\\/g, "\\\\").replace(/'/g, "\\'");
}

function errorStatus(error: unknown): number | undefined {
  if (!error || typeof error !== "object") {
    return undefined;
  }
  const candidate = error as {
    code?: number;
    response?: { status?: number };
  };
  return candidate.response?.status ?? candidate.code;
}

function requireDriveItem(
  file: drive_v3.Schema$File,
  expectedName?: string,
  reused = false,
): DriveEvidenceItem {
  if (!file.id || !file.name || !file.webViewLink) {
    throw new Error(
      `Google Drive did not return id, name, and webViewLink${expectedName ? ` for ${expectedName}` : ""}`,
    );
  }
  return {
    id: file.id,
    name: file.name,
    webViewLink: file.webViewLink,
    reused,
  };
}

function wait(milliseconds: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, milliseconds));
}

export class GoogleDriveEvidencePublisher implements EvidenceDrivePublisher {
  readonly parentFolderId: string;
  readonly folderPrefix: string;
  readonly serviceAccountEmail: string;
  private readonly drive: drive_v3.Drive;

  constructor(options: {
    parentFolderId: string;
    folderPrefix: string;
    serviceAccountJson: string;
  }) {
    const credentials = parseServiceAccountJson(options.serviceAccountJson);
    const auth = new google.auth.GoogleAuth({
      credentials,
      scopes: [DRIVE_SCOPE],
    });
    this.drive = google.drive({ version: "v3", auth });
    this.parentFolderId = options.parentFolderId;
    this.folderPrefix = options.folderPrefix;
    this.serviceAccountEmail = credentials.client_email;
  }

  async validateParentFolder(): Promise<DriveEvidenceItem> {
    const response = await this.drive.files.get({
      fileId: this.parentFolderId,
      fields: "id,name,mimeType,webViewLink",
      supportsAllDrives: true,
    });
    if (response.data.mimeType !== FOLDER_MIME_TYPE) {
      throw new Error("Configured Google Drive parent is not a folder");
    }
    return requireDriveItem(response.data, "configured parent folder");
  }

  async ensureRunFolder(
    runId: string,
    existingFolderId?: string,
  ): Promise<DriveEvidenceItem> {
    if (existingFolderId) {
      try {
        const existing = await this.drive.files.get({
          fileId: existingFolderId,
          fields: "id,name,mimeType,parents,webViewLink",
          supportsAllDrives: true,
        });
        if (
          existing.data.mimeType !== FOLDER_MIME_TYPE ||
          !existing.data.parents?.includes(this.parentFolderId)
        ) {
          throw new Error(
            "Stored evidence folder is not under the configured Drive parent",
          );
        }
        return requireDriveItem(existing.data, "stored evidence folder", true);
      } catch (error) {
        if (errorStatus(error) !== 404) {
          throw error;
        }
      }
    }

    const folderName = `${this.folderPrefix}-${runId}`;
    const found = await this.findChild(folderName, FOLDER_MIME_TYPE);
    if (found) {
      return found;
    }
    const response = await this.drive.files.create({
      requestBody: {
        name: folderName,
        mimeType: FOLDER_MIME_TYPE,
        parents: [this.parentFolderId],
      },
      fields: "id,name,webViewLink",
      supportsAllDrives: true,
    });
    return requireDriveItem(response.data, folderName);
  }

  async uploadPng(options: {
    folderId: string;
    localPath: string;
    fileName: string;
    existingFileId?: string;
  }): Promise<DriveEvidenceItem> {
    if (options.existingFileId) {
      try {
        const existing = await this.drive.files.get({
          fileId: options.existingFileId,
          fields: "id,name,webViewLink",
          supportsAllDrives: true,
        });
        return requireDriveItem(existing.data, options.fileName, true);
      } catch (error) {
        if (errorStatus(error) !== 404) {
          throw error;
        }
      }
    }

    let lastError: unknown;
    for (let attempt = 1; attempt <= 3; attempt += 1) {
      try {
        const found = await this.findChild(options.fileName, "image/png", options.folderId);
        if (found) {
          return found;
        }
        const response = await this.drive.files.create({
          requestBody: {
            name: options.fileName,
            mimeType: "image/png",
            parents: [options.folderId],
          },
          media: {
            mimeType: "image/png",
            body: createReadStream(options.localPath),
          },
          fields: "id,name,webViewLink",
          supportsAllDrives: true,
        });
        return requireDriveItem(response.data, options.fileName);
      } catch (error) {
        lastError = error;
        if (attempt < 3) {
          await wait(500 * 2 ** (attempt - 1));
        }
      }
    }
    throw lastError;
  }

  private async findChild(
    name: string,
    mimeType: string,
    parentFolderId = this.parentFolderId,
  ): Promise<DriveEvidenceItem | undefined> {
    const response = await this.drive.files.list({
      q: `'${driveQueryValue(parentFolderId)}' in parents and name = '${driveQueryValue(name)}' and mimeType = '${mimeType}' and trashed = false`,
      spaces: "drive",
      fields: "files(id,name,webViewLink)",
      pageSize: 10,
      supportsAllDrives: true,
      includeItemsFromAllDrives: true,
    });
    const files = [...(response.data.files ?? [])].sort((left, right) =>
      (left.id ?? "").localeCompare(right.id ?? ""),
    );
    return files[0] ? requireDriveItem(files[0], name, true) : undefined;
  }
}

export function createGoogleDriveEvidencePublisher(
  config: AppConfig,
): GoogleDriveEvidencePublisher {
  if (!config.googleDriveEvidenceEnabled) {
    throw new Error("Google Drive evidence upload is disabled");
  }
  if (!config.googleDriveEvidenceParentFolderId) {
    throw new Error("GOOGLE_DRIVE_EVIDENCE_PARENT_FOLDER is required");
  }
  if (!config.googleServiceAccountJson) {
    throw new Error("GOOGLE_SERVICE_ACCOUNT_JSON is required");
  }
  return new GoogleDriveEvidencePublisher({
    parentFolderId: config.googleDriveEvidenceParentFolderId,
    folderPrefix: config.googleDriveEvidenceFolderPrefix,
    serviceAccountJson: config.googleServiceAccountJson,
  });
}
