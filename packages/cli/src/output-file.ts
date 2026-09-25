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

import { errnoCode, NO_FOLLOW, sameFile, type FileIdentity } from "./document";
import { cliError, FOLIO_CLI_ERROR_CODES, type FolioCliError } from "./errors";

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

export const writeOutputFile = async ({
  target,
  bytes,
  overwrite,
  input,
}: WriteOutputOptions): Promise<Result<string, FolioCliError>> => {
  const absolute = path.resolve(target);
  const directory = await Result.tryPromise(() => realpath(path.dirname(absolute)));
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
  if (existing.isOk()) {
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
    return Result.err(fileSystemError(`Cannot inspect ${destination}`, existing.error.cause));
  }

  const temporary = path.join(directory.value, `.${path.basename(absolute)}.${randomUUID()}.tmp`);
  const written = await Result.tryPromise(async () => {
    const handle = await open(
      temporary,
      constants.O_WRONLY | constants.O_CREAT | constants.O_EXCL | NO_FOLLOW,
      0o644,
    );
    try {
      await handle.writeFile(bytes);
      await handle.sync();
    } finally {
      await handle.close();
    }
    await rename(temporary, destination);
  });
  if (written.isErr()) {
    await rm(temporary, { force: true });
    return Result.err(fileSystemError(`Cannot write ${destination}`, written.error.cause));
  }
  return Result.ok(destination);
};
