import { createReadStream } from "node:fs";
import { createHash } from "node:crypto";
import { readFile } from "node:fs/promises";
import { google, type drive_v3 } from "googleapis";
import type { AppConfig } from "../config";
import {
  resolveGoogleServiceAccount,
  type ServiceAccountCredentials,
} from "./google-service-account";

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
  readonly retestFolderPrefix: string;
  readonly serviceAccountEmail?: string;
  validateParentFolder(): Promise<DriveEvidenceItem>;
  ensureRunFolder(
    runId: string,
    existingFolderId?: string,
    expectedFolderName?: string,
  ): Promise<DriveEvidenceItem>;
  uploadPng(options: {
    folderId: string;
    localPath: string;
    fileName: string;
    existingFileId?: string;
  }): Promise<DriveEvidenceItem>;
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
  readonly retestFolderPrefix: string;
  readonly serviceAccountEmail: string;
  private readonly drive: drive_v3.Drive;

  constructor(options: {
    parentFolderId: string;
    folderPrefix: string;
    retestFolderPrefix: string;
    credentials: ServiceAccountCredentials;
  }) {
    const auth = new google.auth.GoogleAuth({
      credentials: options.credentials,
      scopes: [DRIVE_SCOPE],
    });
    this.drive = google.drive({ version: "v3", auth });
    this.parentFolderId = options.parentFolderId;
    this.folderPrefix = options.folderPrefix;
    this.retestFolderPrefix = options.retestFolderPrefix;
    this.serviceAccountEmail = options.credentials.client_email;
  }

  async validateParentFolder(): Promise<DriveEvidenceItem> {
    const response = await this.drive.files.get({
      fileId: this.parentFolderId,
      fields:
        "id,name,mimeType,trashed,webViewLink,capabilities(canAddChildren)",
      supportsAllDrives: true,
    });
    if (response.data.mimeType !== FOLDER_MIME_TYPE || response.data.trashed) {
      throw new Error("Configured Google Drive parent is not a folder");
    }
    if (response.data.capabilities?.canAddChildren !== true) {
      throw new Error(
        "Configured Google Drive parent does not grant write access to the service account",
      );
    }
    return requireDriveItem(response.data, "configured parent folder");
  }

  async ensureRunFolder(
    runId: string,
    existingFolderId?: string,
    expectedFolderName?: string,
  ): Promise<DriveEvidenceItem> {
    const folderName = expectedFolderName ?? `${this.folderPrefix}-${runId}`;
    if (existingFolderId) {
      try {
        const existing = await this.drive.files.get({
          fileId: existingFolderId,
          fields: "id,name,mimeType,parents,trashed,webViewLink",
          supportsAllDrives: true,
        });
        if (
          existing.data.mimeType === FOLDER_MIME_TYPE &&
          existing.data.parents?.includes(this.parentFolderId) &&
          !existing.data.trashed &&
          existing.data.name === folderName
        ) {
          return requireDriveItem(existing.data, "stored evidence folder", true);
        }
      } catch (error) {
        if (errorStatus(error) !== 404) {
          throw error;
        }
      }
    }

    const found = await this.findChild(folderName, FOLDER_MIME_TYPE);
    if (found) {
      return requireDriveItem(found, folderName, true);
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
    const localMd5 = createHash("md5")
      .update(await readFile(options.localPath))
      .digest("hex");
    let lastError: unknown;
    for (let attempt = 1; attempt <= 3; attempt += 1) {
      try {
        let found: drive_v3.Schema$File | undefined;
        if (options.existingFileId) {
          try {
            const existing = await this.drive.files.get({
              fileId: options.existingFileId,
              fields: "id,name,mimeType,parents,md5Checksum,webViewLink",
              supportsAllDrives: true,
            });
            if (
              existing.data.name === options.fileName &&
              existing.data.mimeType === "image/png" &&
              existing.data.parents?.includes(options.folderId)
            ) {
              found = existing.data;
            }
          } catch (error) {
            if (errorStatus(error) !== 404) {
              throw error;
            }
          }
        }
        found ??= await this.findChild(
          options.fileName,
          "image/png",
          options.folderId,
        );
        if (found?.md5Checksum === localMd5) {
          return requireDriveItem(found, options.fileName, true);
        }
        if (found?.id) {
          const response = await this.drive.files.update({
            fileId: found.id,
            requestBody: {
              name: options.fileName,
              mimeType: "image/png",
            },
            media: {
              mimeType: "image/png",
              body: createReadStream(options.localPath),
            },
            fields: "id,name,webViewLink",
            supportsAllDrives: true,
          });
          return requireDriveItem(response.data, options.fileName);
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
  ): Promise<drive_v3.Schema$File | undefined> {
    const response = await this.drive.files.list({
      q: `'${driveQueryValue(parentFolderId)}' in parents and name = '${driveQueryValue(name)}' and mimeType = '${mimeType}' and trashed = false`,
      spaces: "drive",
      fields: "files(id,name,mimeType,parents,md5Checksum,webViewLink)",
      pageSize: 10,
      supportsAllDrives: true,
      includeItemsFromAllDrives: true,
    });
    const files = [...(response.data.files ?? [])].sort((left, right) =>
      (left.id ?? "").localeCompare(right.id ?? ""),
    );
    return files[0];
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
  const serviceAccount = resolveGoogleServiceAccount(
    config.googleServiceAccount,
  );
  return new GoogleDriveEvidencePublisher({
    parentFolderId: config.googleDriveEvidenceParentFolderId,
    folderPrefix: config.googleDriveEvidenceFolderPrefix,
    retestFolderPrefix: config.googleDriveRetestFolderPrefix,
    credentials: serviceAccount.credentials,
  });
}
