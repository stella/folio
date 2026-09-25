/**
 * Commit one transaction's bytes to its destination: under the document's
 * lease, re-check the source version, stage the package beside the
 * destination, validate it, keep a backup of what an in-place write replaces,
 * journal the commit, and rename the stage into place. Every step before the
 * rename leaves the destination untouched.
 */

import { Result } from "better-result";
import { copyFile, mkdir, open, readdir, readFile, rename, rm, stat } from "node:fs/promises";
import path from "node:path";

import { fileVersionOf } from "./document";
import { cliError, FOLIO_CLI_ERROR_CODES, type FolioCliError } from "./errors";
import { appendJournal, stagePathFor, type JournalCommit } from "./journal";
import { checkPackageIntegrity, type ChangedPart } from "./package-parts";

/** In-place writes keep this many backups beside the document. */
export const MAX_BACKUPS = 20;

export type WriteDestination =
  | { type: "inPlace" }
  | { type: "file"; path: string; overwrite: boolean };

const fileSystemError = (message: string, error: unknown): FolioCliError =>
  cliError({
    code: FOLIO_CLI_ERROR_CODES.internal,
    message: `${message}: ${error instanceof Error ? error.message : String(error)}`,
  });

const readIfPresent = async (filePath: string): Promise<Uint8Array | null> => {
  const bytes = await Result.tryPromise(() => readFile(filePath));
  return bytes.isOk() ? new Uint8Array(bytes.value) : null;
};

/** Write a new file and flush it to disk; refuses to replace an existing one. */
const writeSynced = async (
  filePath: string,
  bytes: Uint8Array,
): Promise<Result<void, FolioCliError>> => {
  const written = await Result.tryPromise(async () => {
    const handle = await open(filePath, "wx");
    try {
      await handle.writeFile(bytes);
      await handle.sync();
    } finally {
      await handle.close();
    }
  });
  return written.isOk()
    ? Result.ok()
    : Result.err(fileSystemError(`Cannot write ${filePath}`, written.error.cause));
};

/** Flush a directory entry change (a rename) where the platform supports it. */
const syncDirectory = async (directory: string): Promise<void> => {
  await Result.tryPromise(async () => {
    const handle = await open(directory, "r");
    try {
      await handle.sync();
    } finally {
      await handle.close();
    }
  });
};

export const backupDirectoryFor = (documentPath: string): string =>
  path.join(path.dirname(documentPath), ".folio", "backups");

/** Keep the newest {@link MAX_BACKUPS} backups by modification time. */
const pruneBackups = async (directory: string): Promise<void> => {
  const names = await Result.tryPromise(() => readdir(directory));
  if (names.isErr()) return;
  const dated = await Promise.all(
    names.value
      .filter((name) => name.endsWith(".docx"))
      .map(async (name) => {
        const info = await Result.tryPromise(() => stat(path.join(directory, name)));
        return { name, mtime: info.isOk() ? info.value.mtimeMs : 0 };
      }),
  );
  const stale = dated.toSorted((left, right) => right.mtime - left.mtime).slice(MAX_BACKUPS);
  await Promise.all(stale.map(({ name }) => rm(path.join(directory, name), { force: true })));
};

/** Copy the file an in-place write replaces to `.folio/backups/<version>.docx`. */
const backUp = async (
  documentPath: string,
  version: string,
): Promise<Result<string, FolioCliError>> => {
  const directory = backupDirectoryFor(documentPath);
  const backupPath = path.join(directory, `${version}.docx`);
  const copied = await Result.tryPromise(async () => {
    await mkdir(directory, { recursive: true });
    await copyFile(documentPath, backupPath);
  });
  if (copied.isErr()) {
    return Result.err(fileSystemError(`Cannot back up ${documentPath}`, copied.error.cause));
  }
  await pruneBackups(directory);
  return Result.ok(backupPath);
};

export type CommitRequest = {
  txId: string;
  destinationPath: string;
  destination: WriteDestination;
  /** The file the transaction read; re-hashed under the lease before anything is written. */
  sourcePath: string;
  fromVersion: string;
  bytes: Uint8Array<ArrayBuffer>;
  changedParts: readonly ChangedPart[];
  journalPath: string;
  /** The journal line, completed with the stage, versions, and backup here. */
  entry: Omit<
    JournalCommit,
    "type" | "path" | "toVersion" | "destinationVersionBefore" | "stage" | "receipt"
  >;
  /** The receipt returned to the caller; completed with `toVersion` and `backup`. */
  receipt: Readonly<Record<string, unknown>>;
};

export type CommitResult = {
  toVersion: string;
  backup: string | undefined;
  receipt: Readonly<Record<string, unknown>>;
};

/**
 * Commit under a lease the caller already holds. Refuses when the source
 * changed since it was read, or when a new destination already exists.
 */
export const commitTransaction = async (
  request: CommitRequest,
): Promise<Result<CommitResult, FolioCliError>> => {
  const { txId, destinationPath, destination, sourcePath, fromVersion, bytes } = request;

  const source = await readIfPresent(sourcePath);
  const sourceVersion = source === null ? null : fileVersionOf(source);
  if (sourceVersion !== fromVersion) {
    return Result.err(
      cliError({
        code: FOLIO_CLI_ERROR_CODES.staleVersion,
        message: `${sourcePath} changed while this transaction ran; nothing was written.`,
        hint: "Re-read the document and retry against its current fileVersion.",
        details: { expected: fromVersion, actual: sourceVersion },
      }),
    );
  }

  const existing = await readIfPresent(destinationPath);
  if (existing !== null && destination.type === "file" && !destination.overwrite) {
    return Result.err(
      cliError({
        code: FOLIO_CLI_ERROR_CODES.destinationExists,
        message: `${destinationPath} already exists.`,
        hint: "Pass --overwrite to replace it, or choose another -o path.",
      }),
    );
  }
  const destinationVersionBefore = existing === null ? null : fileVersionOf(existing);

  const stagePath = stagePathFor(destinationPath, txId);
  const staged = await writeSynced(stagePath, bytes);
  if (staged.isErr()) return Result.err(staged.error);
  const discardStage = () => rm(stagePath, { force: true });

  const valid = await checkPackageIntegrity(bytes, request.changedParts);
  if (valid.isErr()) {
    await discardStage();
    return Result.err(valid.error);
  }

  let backup: string | undefined;
  if (destination.type === "inPlace" && destinationVersionBefore !== null) {
    const backedUp = await backUp(destinationPath, destinationVersionBefore);
    if (backedUp.isErr()) {
      await discardStage();
      return Result.err(backedUp.error);
    }
    backup = backedUp.value;
  }

  const toVersion = fileVersionOf(bytes);
  const receipt = {
    ...request.receipt,
    fileVersion: toVersion,
    ...(backup !== undefined && { backup }),
  };
  const journaled = await appendJournal(request.journalPath, {
    type: "commit",
    ...request.entry,
    path: destinationPath,
    toVersion,
    destinationVersionBefore,
    stage: path.basename(stagePath),
    receipt,
  });
  if (journaled.isErr()) {
    await discardStage();
    return Result.err(journaled.error);
  }

  const renamed = await Result.tryPromise(() => rename(stagePath, destinationPath));
  if (renamed.isErr()) {
    // The journal line stands: the next write to this document rolls the
    // stage forward, or discards it if the destination moved on meanwhile.
    return Result.err(
      fileSystemError(
        `Cannot move the staged package into ${destinationPath}`,
        renamed.error.cause,
      ),
    );
  }
  await syncDirectory(path.dirname(destinationPath));
  return Result.ok({ toVersion, backup, receipt });
};
