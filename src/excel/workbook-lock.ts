import { randomUUID } from "node:crypto";
import {
  mkdir,
  readFile,
  readdir,
  rename,
  rm,
  rmdir,
  stat,
  unlink,
  writeFile,
} from "node:fs/promises";
import path from "node:path";

const ORPHANED_LOCK_AGE_MS = 30_000;

interface WorkbookLockRecord {
  pid: number;
  purpose: string;
  createdAt: string;
  token?: string;
}

interface ExistingWorkbookLock {
  directory: boolean;
  emptyDirectory?: boolean;
  missing?: boolean;
  oldEnoughToRecover?: boolean;
  ownerPath?: string;
  record?: WorkbookLockRecord;
}

function processIsRunning(pid: number): boolean {
  try {
    process.kill(pid, 0);
    return true;
  } catch (error) {
    return (error as NodeJS.ErrnoException).code !== "ESRCH";
  }
}

async function oldEnoughToRecover(targetPath: string): Promise<boolean> {
  try {
    const details = await stat(targetPath);
    return Date.now() - details.mtimeMs >= ORPHANED_LOCK_AGE_MS;
  } catch {
    return false;
  }
}

async function readExistingLock(
  lockPath: string,
): Promise<ExistingWorkbookLock> {
  try {
    const entries = await readdir(lockPath);
    const ownerEntries = entries.filter(
      (entry) => entry.startsWith("owner-") && entry.endsWith(".json"),
    );
    if (!ownerEntries.length) {
      return {
        directory: true,
        emptyDirectory: entries.length === 0,
        oldEnoughToRecover:
          entries.length === 0 && (await oldEnoughToRecover(lockPath)),
      };
    }
    const ownerPath = path.join(lockPath, ownerEntries[0]);
    try {
      const record = JSON.parse(
        await readFile(ownerPath, "utf8"),
      ) as WorkbookLockRecord;
      const expectedOwnerName = record.token
        ? `owner-${record.token}.json`
        : ownerEntries[0];
      if (ownerEntries[0] !== expectedOwnerName) {
        throw new Error("Lock owner token does not match its file name");
      }
      return { directory: true, ownerPath, record };
    } catch {
      return {
        directory: true,
        ownerPath,
        oldEnoughToRecover: await oldEnoughToRecover(ownerPath),
      };
    }
  } catch (error) {
    const code = (error as NodeJS.ErrnoException).code;
    if (code === "ENOENT") {
      return { directory: false, missing: true };
    }
    if (code !== "ENOTDIR") {
      return { directory: false };
    }
    try {
      return {
        directory: false,
        record: JSON.parse(
          await readFile(lockPath, "utf8"),
        ) as WorkbookLockRecord,
      };
    } catch {
      return { directory: false };
    }
  }
}

async function removeOwnedLock(
  lockPath: string,
  ownerPath: string,
): Promise<void> {
  try {
    await unlink(ownerPath);
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code === "ENOENT") {
      return;
    }
    throw error;
  }
  try {
    await rmdir(lockPath);
  } catch (error) {
    const code = (error as NodeJS.ErrnoException).code;
    if (code !== "ENOENT" && code !== "ENOTEMPTY" && code !== "EEXIST") {
      throw error;
    }
  }
}

async function removeEmptyLockDirectory(lockPath: string): Promise<void> {
  try {
    await rmdir(lockPath);
  } catch (error) {
    const code = (error as NodeJS.ErrnoException).code;
    if (code !== "ENOENT" && code !== "ENOTEMPTY" && code !== "EEXIST") {
      throw error;
    }
  }
}

export async function acquireWorkbookLock(
  workbookPath: string,
  purpose: string,
): Promise<() => Promise<void>> {
  const lockPath = `${workbookPath}.lock`;
  await mkdir(path.dirname(lockPath), { recursive: true });
  const token = randomUUID();
  const ownerName = `owner-${token}.json`;
  const ownerPath = path.join(lockPath, ownerName);
  const pendingOwnerPath = `${lockPath}.pending-${token}.lock`;
  const record: WorkbookLockRecord = {
    pid: process.pid,
    purpose,
    createdAt: new Date().toISOString(),
    token,
  };
  await writeFile(pendingOwnerPath, JSON.stringify(record), { flag: "wx" });
  try {
    for (let attempt = 0; attempt < 4; attempt += 1) {
      try {
        await mkdir(lockPath);
      } catch (error) {
        if ((error as NodeJS.ErrnoException).code !== "EEXIST") {
          throw error;
        }
        const existing = await readExistingLock(lockPath);
        if (existing.missing) {
          continue;
        }
        if (
          attempt < 3 &&
          existing.ownerPath &&
          existing.record &&
          Number.isInteger(existing.record.pid) &&
          !processIsRunning(existing.record.pid)
        ) {
          await removeOwnedLock(lockPath, existing.ownerPath);
          continue;
        }
        if (
          attempt < 3 &&
          existing.directory &&
          existing.ownerPath &&
          !existing.record &&
          existing.oldEnoughToRecover
        ) {
          await removeOwnedLock(lockPath, existing.ownerPath);
          continue;
        }
        if (
          attempt < 3 &&
          existing.emptyDirectory &&
          existing.oldEnoughToRecover
        ) {
          await removeEmptyLockDirectory(lockPath);
          continue;
        }
        const staleLegacyLock =
          !existing.directory &&
          existing.record &&
          Number.isInteger(existing.record.pid) &&
          !processIsRunning(existing.record.pid);
        throw new Error(
          staleLegacyLock
            ? "Executed workbook has a stale legacy lock; remove it after confirming no workbook process is running"
            : `Executed workbook is locked${existing.record?.purpose ? ` by ${existing.record.purpose}` : ""}`,
        );
      }

      let ownerInstalled = false;
      try {
        await rename(pendingOwnerPath, ownerPath);
        ownerInstalled = true;
        const entries = await readdir(lockPath);
        if (entries.length !== 1 || entries[0] !== ownerName) {
          throw new Error("Executed workbook lock ownership changed during acquisition");
        }
      } catch (error) {
        if (ownerInstalled) {
          await removeOwnedLock(lockPath, ownerPath);
        } else {
          await removeEmptyLockDirectory(lockPath);
        }
        throw error;
      }

      let released = false;
      return async () => {
        if (released) {
          return;
        }
        released = true;
        await removeOwnedLock(lockPath, ownerPath);
      };
    }
    throw new Error("Could not acquire executed workbook lock");
  } finally {
    await rm(pendingOwnerPath, { force: true });
  }
}
