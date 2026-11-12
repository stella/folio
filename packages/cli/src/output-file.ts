/**
 * Write a rendered file (PDF, PNG, HTML) next to where the caller asked: to a
 * temporary name in the same directory, flushed, then renamed into place, so
 * a reader never sees half a file. An existing file is replaced only with
 * `overwrite`, and never when it is the input document.
 */

import { Result } from "better-result";
import { randomUUID } from "node:crypto";
import { constants } from "node:fs";
import { open, realpath, rename, rm, stat } from "node:fs/promises";
import path from "node:path";

import { errnoCode, type FileIdentity, NO_FOLLOW, sameFile } from "./document";
import { cliError, FOLIO_CLI_ERROR_CODES, type FolioCliError } from "./errors";
import { inspectPath, unsafePath } from "./sidecar";

type WriteOutputOptions = {
  target: string;
  bytes: Uint8Array;
  overwrite: boolean;
  /** The document rendered, which the output must not replace. */
  input: FileIdentity;
};

const fileSystemError = (message: string, error: unknown): FolioCliError =>
  cliError({
    code: FOLIO_CLI_ERROR_CODES.internal,
    message: `${message}: ${error instanceof Error ? error.message : String(error)}`,
  });

type RenameGuard = {
  directory: string;
  temporary: string;
  temporaryIdentity: FileIdentity;
  destination: string;
  /** The destination as checked: absent, or the file the caller may replace. */
  checked: FileIdentity | null;
};

/**
 * Just before the rename: the directory still resolves where it did, the
 * temporary file is still the one written, and the destination is still the
 * file checked (or still absent).
 */
const checkBeforeRename = async (
  guard: RenameGuard,
): Promise<Result<void, FolioCliError>> => {
  const directory = await Result.tryPromise(() => realpath(guard.directory));
  if (directory.isErr() || directory.value !== guard.directory) {
    return Result.err(unsafePath(`${guard.directory} moved during the write.`));
  }
  const temporary = await inspectPath(guard.temporary);
  if (temporary.isErr()) return Result.err(temporary.error);
  if (
    temporary.value.type !== "file" ||
    !sameFile(temporary.value.identity, guard.temporaryIdentity)
  ) {
    return Result.err(
      unsafePath(`${guard.temporary} was replaced during the write.`),
    );
  }
  // Checked as before: through a final symlink, which the rename replaces.
  const destination = await Result.tryPromise(() => stat(guard.destination));
  const unchanged =
    guard.checked === null
      ? destination.isErr() && errnoCode(destination.error.cause) === "ENOENT"
      : destination.isOk() &&
        sameFile(
          { dev: destination.value.dev, ino: destination.value.ino },
          guard.checked,
        );
  return unchanged
    ? Result.ok()
    : Result.err(
        unsafePath(`${guard.destination} was replaced during the write.`),
      );
};

export const writeOutputFile = async ({
  target,
  bytes,
  overwrite,
  input,
}: WriteOutputOptions): Promise<Result<string, FolioCliError>> => {
  const absolute = path.resolve(target);
  const directory = await Result.tryPromise(() =>
    realpath(path.dirname(absolute)),
  );
  if (directory.isErr()) {
    return Result.err(
      cliError({
        code: FOLIO_CLI_ERROR_CODES.notFound,
        message: `No directory ${path.dirname(absolute)}.`,
      }),
    );
  }
  const destination = path.join(directory.value, path.basename(absolute));
  const existing = await Result.tryPromise(() => stat(destination));
  let checked: FileIdentity | null = null;
  if (existing.isOk()) {
    checked = { dev: existing.value.dev, ino: existing.value.ino };
    if (sameFile({ dev: existing.value.dev, ino: existing.value.ino }, input)) {
      return Result.err(
        cliError({
          code: FOLIO_CLI_ERROR_CODES.invalidDestination,
          message: `${destination} is the document being rendered.`,
        }),
      );
    }
    if (!overwrite) {
      return Result.err(
        cliError({
          code: FOLIO_CLI_ERROR_CODES.destinationExists,
          message: `${destination} already exists.`,
          hint: "Pass --overwrite to replace it.",
        }),
      );
    }
  } else if (errnoCode(existing.error.cause) !== "ENOENT") {
    return Result.err(
      fileSystemError(`Cannot inspect ${destination}`, existing.error.cause),
    );
  }

  const temporary = path.join(
    directory.value,
    `.${path.basename(absolute)}.${randomUUID()}.tmp`,
  );
  const written = await Result.tryPromise(async () => {
    const handle = await open(
      temporary,
      constants.O_WRONLY | constants.O_CREAT | constants.O_EXCL | NO_FOLLOW,
      0o644,
    );
    try {
      await handle.writeFile(bytes);
      await handle.sync();
      const info = await handle.stat();
      return { dev: info.dev, ino: info.ino };
    } finally {
      await handle.close();
    }
  });
  if (written.isErr()) {
    await rm(temporary, { force: true });
    return Result.err(
      fileSystemError(`Cannot write ${destination}`, written.error.cause),
    );
  }
  const guarded = await checkBeforeRename({
    directory: directory.value,
    temporary,
    temporaryIdentity: written.value,
    destination,
    checked,
  });
  if (guarded.isErr()) {
    await rm(temporary, { force: true });
    return Result.err(guarded.error);
  }
  const renamed = await Result.tryPromise(() => rename(temporary, destination));
  if (renamed.isErr()) {
    await rm(temporary, { force: true });
    return Result.err(
      fileSystemError(`Cannot write ${destination}`, renamed.error.cause),
    );
  }
  return Result.ok(destination);
};
