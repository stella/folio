/**
 * Commit one transaction's bytes to its destination: under the document's
 * lease, re-check the source and destination versions, stage the package
 * beside the destination, validate it, back up whatever the write replaces,
 * confirm the lease is still ours, journal the commit, and rename the stage
 * into place. Every step before the rename leaves the destination untouched.
 */

import { Result } from "better-result";
import { readdir, realpath, rename, rm } from "node:fs/promises";
import path from "node:path";

import { fileVersionOf, readDocumentFile, sameFile, type FileIdentity } from "./document";
import { cliError, FOLIO_CLI_ERROR_CODES, type FolioCliError } from "./errors";
import {
  appendJournal,
  stagePathFor,
  type JournalCommit,
  type RecoveryDiscardReason,
} from "./journal";
import { checkPackageIntegrity, type ChangedPart } from "./package-parts";
import {
  inspectPath,
  readSidecarFile,
  sidecarDirectory,
  syncDirectory,
  unsafePath,
  writeNewSidecarFile,
  type PathEntry,
} from "./sidecar";

/** Backups kept per document. */
export const MAX_BACKUPS = 20;

export type WriteDestination =
  | { type: "inPlace" }
  | {
      type: "file";
      path: string;
      /** Allow replacing an existing file; its version must then be `expectedVersion`. */
      overwrite: boolean;
      expectedVersion: string | undefined;
    };

const fileSystemError = (message: string, error: unknown): FolioCliError =>
  cliError({
    code: FOLIO_CLI_ERROR_CODES.internal,
    message: `${message}: ${error instanceof Error ? error.message : String(error)}`,
  });

const BACKUP_NAME = /^[0-9a-f]{64}\.docx$/u;

/** `.folio/backups/<document name>`, created private and never through a symlink. */
export const backupDirectoryFor = (documentPath: string) =>
  sidecarDirectory(documentPath, "backups", path.basename(documentPath));

/**
 * Keep the newest `keep` backups in one document's backup directory. Only
 * regular files named `<version>.docx` count; nothing else is touched.
 */
export const pruneBackups = async (
  directory: string,
  keep: number,
): Promise<Result<void, FolioCliError>> => {
  const names = await Result.tryPromise(() => readdir(directory));
  if (names.isErr())
    return Result.err(fileSystemError(`Cannot list ${directory}`, names.error.cause));
  const backups: { name: string; modifiedMs: number }[] = [];
  for (const name of names.value.filter((entry) => BACKUP_NAME.test(entry))) {
    const entry = await inspectPath(path.join(directory, name));
    if (entry.isErr()) return Result.err(entry.error);
    if (entry.value.type === "file") backups.push({ name, modifiedMs: entry.value.modifiedMs });
  }
  const stale = backups
    .toSorted(
      (left, right) => right.modifiedMs - left.modifiedMs || right.name.localeCompare(left.name),
    )
    .slice(keep);
  for (const { name } of stale) {
    await rm(path.join(directory, name), { force: true });
  }
  return Result.ok();
};

type BackUpOptions = { documentPath: string; version: string; bytes: Uint8Array };

/**
 * Write the bytes a transaction replaces to
 * `.folio/backups/<document name>/<version>.docx`, flushed with its
 * directory. A backup of the same version already there is kept when it
 * holds those bytes.
 */
const backUp = async ({
  documentPath,
  version,
  bytes,
}: BackUpOptions): Promise<Result<string, FolioCliError>> => {
  const directory = await backupDirectoryFor(documentPath);
  if (directory.isErr()) return Result.err(directory.error);
  const backupPath = path.join(directory.value, `${version}.docx`);
  const existing = await inspectPath(backupPath);
  if (existing.isErr()) return Result.err(existing.error);
  if (existing.value.type !== "missing") {
    const kept = await readSidecarFile(backupPath);
    if (kept.isErr()) return Result.err(kept.error);
    if (fileVersionOf(kept.value) !== version) {
      return Result.err(unsafePath(`${backupPath} does not hold the version its name says.`));
    }
  } else {
    const written = await writeNewSidecarFile(backupPath, bytes);
    if (written.isErr()) return Result.err(written.error);
    const synced = await syncDirectory(directory.value);
    if (synced.isErr()) return Result.err(synced.error);
  }
  const pruned = await pruneBackups(directory.value, MAX_BACKUPS);
  return pruned.isErr() ? Result.err(pruned.error) : Result.ok(backupPath);
};

export type CommitRequest = {
  txId: string;
  destinationPath: string;
  destination: WriteDestination;
  /** The file the transaction read; re-read under the lease before anything is written. */
  sourcePath: string;
  sourceIdentity: FileIdentity;
  fromVersion: string;
  bytes: Uint8Array<ArrayBuffer>;
  changedParts: readonly ChangedPart[];
  journalPath: string;
  /** Refuses unless the transaction still holds the destination's lease. */
  verifyLease: () => Promise<Result<void, FolioCliError>>;
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

type ExistingDestination =
  | { type: "missing" }
  | { type: "file"; identity: FileIdentity; version: string; bytes: Uint8Array<ArrayBuffer> };

/** The destination as it is now: missing, or a regular single-link file read without following a symlink. */
const readDestination = async (
  destinationPath: string,
): Promise<Result<ExistingDestination, FolioCliError>> => {
  const entry = await inspectPath(destinationPath);
  if (entry.isErr()) return Result.err(entry.error);
  if (entry.value.type === "missing") return Result.ok({ type: "missing" });
  if (entry.value.type !== "file" || entry.value.links !== 1) {
    return Result.err(
      unsafePath(
        `${destinationPath} is not a regular file with a single link; folio does not replace it.`,
      ),
    );
  }
  const bytes = await readSidecarFile(destinationPath);
  if (bytes.isErr()) return Result.err(bytes.error);
  return Result.ok({
    type: "file",
    identity: entry.value.identity,
    version: fileVersionOf(bytes.value),
    bytes: bytes.value,
  });
};

const staleError = (message: string, expected: string, actual: string | null): FolioCliError =>
  cliError({
    code: FOLIO_CLI_ERROR_CODES.staleVersion,
    message,
    hint: "Re-read the document and retry against its current fileVersion.",
    details: { expected, actual },
  });

/** The replaced destination must be exactly the version the caller named. */
const checkDestination = (
  request: CommitRequest,
  existing: ExistingDestination,
): Result<void, FolioCliError> => {
  const { destination, destinationPath } = request;
  if (destination.type === "inPlace" || existing.type === "missing") {
    return destination.type === "file" && destination.expectedVersion !== undefined
      ? Result.err(
          staleError(
            `${destinationPath} does not exist, but a destination version was expected.`,
            destination.expectedVersion,
            null,
          ),
        )
      : Result.ok();
  }
  if (!destination.overwrite) {
    return Result.err(
      cliError({
        code: FOLIO_CLI_ERROR_CODES.destinationExists,
        message: `${destinationPath} already exists.`,
        hint: "Pass --overwrite with --expect-destination-version to replace it, or choose another -o path.",
      }),
    );
  }
  if (destination.expectedVersion === undefined) {
    return Result.err(
      cliError({
        code: FOLIO_CLI_ERROR_CODES.invalidInput,
        message: `Replacing ${destinationPath} needs its current fileVersion.`,
        hint: "Read it (folio read) and pass --expect-destination-version (MCP: expectedDestinationVersion).",
      }),
    );
  }
  return destination.expectedVersion === existing.version
    ? Result.ok()
    : Result.err(
        staleError(
          `${destinationPath} changed since it was read; nothing was written.`,
          destination.expectedVersion,
          existing.version,
        ),
      );
};

type RenameGuard = {
  parentRealPath: string;
  stagePath: string;
  stageIdentity: FileIdentity;
  destinationPath: string;
  existing: ExistingDestination;
};

const sameEntry = (entry: PathEntry, identity: FileIdentity): boolean =>
  entry.type === "file" && entry.links === 1 && sameFile(entry.identity, identity);

/**
 * Just before the rename: the directory still resolves where it did, the
 * stage is still the file written, and the destination is still the file
 * checked (or still absent).
 */
const checkBeforeRename = async (guard: RenameGuard): Promise<Result<void, FolioCliError>> => {
  const parent = await Result.tryPromise(() => realpath(path.dirname(guard.destinationPath)));
  if (parent.isErr() || parent.value !== guard.parentRealPath) {
    return Result.err(unsafePath(`${path.dirname(guard.destinationPath)} moved during the write.`));
  }
  const stage = await inspectPath(guard.stagePath);
  if (stage.isErr()) return Result.err(stage.error);
  if (!sameEntry(stage.value, guard.stageIdentity)) {
    return Result.err(unsafePath(`${guard.stagePath} was replaced during the write.`));
  }
  const destination = await inspectPath(guard.destinationPath);
  if (destination.isErr()) return Result.err(destination.error);
  const unchanged =
    guard.existing.type === "missing"
      ? destination.value.type === "missing"
      : sameEntry(destination.value, guard.existing.identity);
  return unchanged
    ? Result.ok()
    : Result.err(unsafePath(`${guard.destinationPath} was replaced during the write.`));
};

/**
 * Commit under a lease the caller already holds. Refuses when the source
 * changed since it was read, when a new destination already exists, or when
 * an existing destination is not the version the caller named.
 */
export const commitTransaction = async (
  request: CommitRequest,
): Promise<Result<CommitResult, FolioCliError>> => {
  const { txId, destinationPath, sourcePath, fromVersion, bytes } = request;

  const source = await readDocumentFile(sourcePath);
  const sourceVersion = source.isOk() ? source.value.fileVersion : null;
  if (
    sourceVersion !== fromVersion ||
    !source.isOk() ||
    !sameFile(source.value.identity, request.sourceIdentity)
  ) {
    return Result.err(
      staleError(
        `${sourcePath} changed while this transaction ran; nothing was written.`,
        fromVersion,
        sourceVersion,
      ),
    );
  }

  const existing = await readDestination(destinationPath);
  if (existing.isErr()) return Result.err(existing.error);
  const destinationChecked = checkDestination(request, existing.value);
  if (destinationChecked.isErr()) return Result.err(destinationChecked.error);
  const destinationVersionBefore = existing.value.type === "file" ? existing.value.version : null;
  const parentRealPath = await Result.tryPromise(() => realpath(path.dirname(destinationPath)));
  if (parentRealPath.isErr()) {
    return Result.err(
      fileSystemError(`Cannot resolve ${destinationPath}`, parentRealPath.error.cause),
    );
  }

  const stagePath = stagePathFor(destinationPath, txId);
  const staged = await writeNewSidecarFile(stagePath, bytes);
  if (staged.isErr()) return Result.err(staged.error);
  const discardStage = async <T>(error: FolioCliError): Promise<Result<T, FolioCliError>> => {
    await rm(stagePath, { force: true });
    return Result.err(error);
  };

  const valid = await checkPackageIntegrity(bytes, request.changedParts);
  if (valid.isErr()) return await discardStage(valid.error);

  let backup: string | undefined;
  if (existing.value.type === "file") {
    const backedUp = await backUp({
      documentPath: destinationPath,
      version: existing.value.version,
      bytes: existing.value.bytes,
    });
    if (backedUp.isErr()) return await discardStage(backedUp.error);
    backup = backedUp.value;
  }

  const stillOurs = await request.verifyLease();
  if (stillOurs.isErr()) return await discardStage(stillOurs.error);

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
  if (journaled.isErr()) return await discardStage(journaled.error);

  // The journal line stands from here. A refusal before the rename records
  // the discard, so recovery never completes a transaction that lost its
  // lease or found its files replaced.
  const abandon = async (
    reason: RecoveryDiscardReason,
    error: FolioCliError,
  ): Promise<Result<CommitResult, FolioCliError>> => {
    await rm(stagePath, { force: true });
    const recorded = await appendJournal(request.journalPath, {
      type: "recovery",
      txId,
      path: destinationPath,
      time: request.entry.time,
      action: "discarded",
      reason,
    });
    return Result.err(recorded.isErr() ? recorded.error : error);
  };
  const fenced = await request.verifyLease();
  if (fenced.isErr()) return await abandon("leaseLost", fenced.error);
  const guarded = await checkBeforeRename({
    parentRealPath: parentRealPath.value,
    stagePath,
    stageIdentity: staged.value,
    destinationPath,
    existing: existing.value,
  });
  if (guarded.isErr()) return await abandon("renameRefused", guarded.error);

  const renamed = await Result.tryPromise(() => rename(stagePath, destinationPath));
  if (renamed.isErr()) {
    return Result.err(
      fileSystemError(
        `Cannot move the staged package into ${destinationPath}`,
        renamed.error.cause,
      ),
    );
  }
  const synced = await syncDirectory(path.dirname(destinationPath));
  if (synced.isErr()) {
    return Result.err(
      cliError({
        code: FOLIO_CLI_ERROR_CODES.internal,
        message: `${destinationPath} was replaced, but its directory could not be flushed to disk: ${synced.error.message}`,
        details: { receipt },
      }),
    );
  }
  return Result.ok({ toVersion, backup, receipt });
};
