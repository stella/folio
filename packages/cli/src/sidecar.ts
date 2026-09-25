/**
 * The one boundary every file folio keeps beside a document goes through:
 * the lock, the stage, the journal, and backups. None of them follows a
 * symlink or writes through a hard link. The `.folio` directory and its
 * subdirectories are created private (0700), and an existing one that is a
 * symlink or not a directory is refused rather than followed.
 */

import { Result } from "better-result";
import { constants } from "node:fs";
import { lstat, mkdir, open, type FileHandle } from "node:fs/promises";
import path from "node:path";

import { errnoCode, NO_FOLLOW, type FileIdentity } from "./document";
import { cliError, FOLIO_CLI_ERROR_CODES, type FolioCliError } from "./errors";

export const SIDECAR_DIRECTORY = ".folio";

const PRIVATE_DIRECTORY_MODE = 0o700;
const PRIVATE_FILE_MODE = 0o600;

const describe = (error: unknown): string =>
  error instanceof Error ? error.message : String(error);

export const unsafePath = (message: string, hint?: string): FolioCliError =>
  cliError({ code: FOLIO_CLI_ERROR_CODES.unsafePath, message, hint });

const fileSystemError = (message: string, error: unknown): FolioCliError =>
  cliError({ code: FOLIO_CLI_ERROR_CODES.internal, message: `${message}: ${describe(error)}` });

/** What a path names, without following a final symlink. */
export type PathEntry =
  | { type: "missing" }
  | { type: "file"; identity: FileIdentity; links: number; modifiedMs: number }
  | { type: "directory"; identity: FileIdentity }
  | { type: "other" };

export const inspectPath = async (filePath: string): Promise<Result<PathEntry, FolioCliError>> => {
  const info = await Result.tryPromise(() => lstat(filePath));
  if (info.isErr()) {
    return errnoCode(info.error.cause) === "ENOENT"
      ? Result.ok({ type: "missing" })
      : Result.err(fileSystemError(`Cannot inspect ${filePath}`, info.error.cause));
  }
  const stats = info.value;
  const identity = { dev: stats.dev, ino: stats.ino };
  if (stats.isFile()) {
    return Result.ok({ type: "file", identity, links: stats.nlink, modifiedMs: stats.mtimeMs });
  }
  if (stats.isDirectory()) return Result.ok({ type: "directory", identity });
  return Result.ok({ type: "other" });
};

/**
 * Create `directory` private when missing; refuse it when it exists as a
 * symlink or anything but a directory. Its parent must already exist.
 */
const ensurePrivateDirectory = async (directory: string): Promise<Result<void, FolioCliError>> => {
  for (let attempt = 0; attempt < 2; attempt++) {
    const entry = await inspectPath(directory);
    if (entry.isErr()) return Result.err(entry.error);
    if (entry.value.type === "directory") return Result.ok();
    if (entry.value.type !== "missing") {
      return Result.err(
        unsafePath(
          `${directory} is not a plain directory; folio does not write through it.`,
          "Remove or rename it (it may be a symlink), then retry.",
        ),
      );
    }
    const made = await Result.tryPromise(() => mkdir(directory, { mode: PRIVATE_DIRECTORY_MODE }));
    if (made.isOk()) return Result.ok();
    if (errnoCode(made.error.cause) !== "EEXIST") {
      return Result.err(fileSystemError(`Cannot create ${directory}`, made.error.cause));
    }
  }
  return Result.err(unsafePath(`${directory} changed while it was being created.`));
};

/**
 * `<document directory>/.folio/<segments...>`, each level checked or created
 * private. The document's own directory is the caller's choice and is not
 * inspected.
 */
export const sidecarDirectory = async (
  documentPath: string,
  ...segments: readonly string[]
): Promise<Result<string, FolioCliError>> => {
  let directory = path.join(path.dirname(documentPath), SIDECAR_DIRECTORY);
  const top = await ensurePrivateDirectory(directory);
  if (top.isErr()) return Result.err(top.error);
  for (const segment of segments) {
    if (segment === "" || segment === "." || segment === ".." || segment.includes(path.sep)) {
      return Result.err(unsafePath(`${JSON.stringify(segment)} is not a plain name.`));
    }
    directory = path.join(directory, segment);
    const level = await ensurePrivateDirectory(directory);
    if (level.isErr()) return Result.err(level.error);
  }
  return Result.ok(directory);
};

/** Check the directory a sidecar file lives in; create it only when it is `.folio` itself. */
export const ensureSidecarParent = async (
  filePath: string,
): Promise<Result<void, FolioCliError>> => {
  const parent = path.dirname(filePath);
  if (path.basename(parent) === SIDECAR_DIRECTORY) return await ensurePrivateDirectory(parent);
  const entry = await inspectPath(parent);
  if (entry.isErr()) return Result.err(entry.error);
  return entry.value.type === "directory"
    ? Result.ok()
    : Result.err(unsafePath(`${parent} is not a plain directory.`));
};

type OpenedSidecar = { handle: FileHandle; identity: FileIdentity; size: number };

/**
 * Open a sidecar file without following a final symlink, and refuse it
 * unless it is a regular file with exactly one link.
 */
export const openSidecarFile = async (
  filePath: string,
  flags: number,
): Promise<Result<OpenedSidecar, FolioCliError>> => {
  const opened = await Result.tryPromise(() =>
    open(filePath, flags | NO_FOLLOW, PRIVATE_FILE_MODE),
  );
  if (opened.isErr()) {
    const code = errnoCode(opened.error.cause);
    return code === "ELOOP" || code === "EMLINK"
      ? Result.err(unsafePath(`${filePath} is a symlink; folio does not follow it.`))
      : Result.err(fileSystemError(`Cannot open ${filePath}`, opened.error.cause));
  }
  const handle = opened.value;
  const info = await Result.tryPromise(() => handle.stat());
  if (info.isErr() || !info.value.isFile() || info.value.nlink !== 1) {
    await handle.close();
    return Result.err(
      unsafePath(`${filePath} is not a regular file with a single link; folio does not use it.`),
    );
  }
  return Result.ok({
    handle,
    identity: { dev: info.value.dev, ino: info.value.ino },
    size: info.value.size,
  });
};

/** Read a whole sidecar file through {@link openSidecarFile}. */
export const readSidecarFile = async (
  filePath: string,
): Promise<Result<Uint8Array<ArrayBuffer>, FolioCliError>> => {
  const opened = await openSidecarFile(filePath, constants.O_RDONLY);
  if (opened.isErr()) return Result.err(opened.error);
  const { handle } = opened.value;
  const read = await Result.tryPromise(async () => new Uint8Array(await handle.readFile()));
  await handle.close();
  return read.isOk()
    ? Result.ok(read.value)
    : Result.err(fileSystemError(`Cannot read ${filePath}`, read.error.cause));
};

/** Create a new sidecar file (never an existing one), write it, and flush it to disk. */
export const writeNewSidecarFile = async (
  filePath: string,
  bytes: Uint8Array,
): Promise<Result<FileIdentity, FolioCliError>> => {
  const opened = await openSidecarFile(
    filePath,
    constants.O_WRONLY | constants.O_CREAT | constants.O_EXCL,
  );
  if (opened.isErr()) return Result.err(opened.error);
  const { handle, identity } = opened.value;
  const written = await Result.tryPromise(async () => {
    try {
      await handle.writeFile(bytes);
      await handle.sync();
    } finally {
      await handle.close();
    }
  });
  return written.isOk()
    ? Result.ok(identity)
    : Result.err(fileSystemError(`Cannot write ${filePath}`, written.error.cause));
};

/**
 * Flush a directory's entries (a create or rename) to disk. Platforms that
 * cannot open a directory for syncing report nothing; every other failure
 * is returned.
 */
export const syncDirectory = async (directory: string): Promise<Result<void, FolioCliError>> => {
  if (process.platform === "win32") return Result.ok();
  const synced = await Result.tryPromise(async () => {
    const handle = await open(directory, constants.O_RDONLY);
    try {
      await handle.sync();
    } finally {
      await handle.close();
    }
  });
  return synced.isOk()
    ? Result.ok()
    : Result.err(fileSystemError(`Cannot flush ${directory}`, synced.error.cause));
};
