import { Result } from "better-result";
import { createHash } from "node:crypto";
import { constants } from "node:fs";
import { open, realpath, stat } from "node:fs/promises";
import path from "node:path";

import { FolioDocxReviewer } from "@stll/folio-core/server";

import { cliError, FOLIO_CLI_ERROR_CODES, type FolioCliError } from "./errors";

/** Largest `.docx` the CLI opens; larger inputs are refused before reading. */
export const MAX_DOCUMENT_BYTES = 64 * 1024 * 1024;

const FILE_VERSION_PATTERN = /^[0-9a-f]{64}$/u;

/** A file's version: the lowercase hex SHA-256 of its bytes. */
export const fileVersionOf = (bytes: Uint8Array): string =>
  createHash("sha256").update(bytes).digest("hex");

export const isFileVersion = (value: unknown): value is string =>
  typeof value === "string" && FILE_VERSION_PATTERN.test(value);

/** Which file a path named when it was read: device and inode. */
export type FileIdentity = { dev: number; ino: number };

export const sameFile = (left: FileIdentity, right: FileIdentity): boolean =>
  left.dev === right.dev && left.ino === right.ino;

/** The bytes of one input file and the version they hash to. */
export type LoadedFile = {
  /** Real path (symlinks resolved) the bytes were read from. */
  path: string;
  bytes: Uint8Array<ArrayBuffer>;
  fileVersion: string;
  identity: FileIdentity;
  /** Hard links to the file; a write refuses a file with more than one. */
  links: number;
};

/**
 * `O_NOFOLLOW` where the platform has it: opening a path whose last
 * component is a symlink fails instead of following it.
 */
export const NO_FOLLOW: number = constants.O_NOFOLLOW ?? 0;

/** The `code` of a Node.js file-system error, when there is one. */
export const errnoCode = (error: unknown): string | undefined => {
  if (typeof error !== "object" || error === null || !("code" in error)) {
    return undefined;
  }
  return typeof error.code === "string" ? error.code : undefined;
};

const describeError = (error: unknown): string =>
  error instanceof Error ? error.message : String(error);

const fileSystemError = (filePath: string, error: unknown): FolioCliError =>
  errnoCode(error) === "ENOENT"
    ? cliError({
        code: FOLIO_CLI_ERROR_CODES.notFound,
        message: `No file at ${filePath}.`,
      })
    : cliError({
        code: FOLIO_CLI_ERROR_CODES.invalidInput,
        message: `Cannot read ${filePath}: ${describeError(error)}`,
      });

type OpenedFile = { bytes: Uint8Array<ArrayBuffer>; identity: FileIdentity; links: number };

/**
 * Read through one handle: the path is resolved, stat'ed, opened without
 * following a final symlink, and the handle must be the same file the stat
 * saw, so a swap between the check and the read is refused.
 */
const readThroughHandle = async (real: string): Promise<Result<OpenedFile, FolioCliError>> => {
  const checked = await Result.tryPromise({
    try: () => stat(real),
    catch: (error) => fileSystemError(real, error),
  });
  if (checked.isErr()) return Result.err(checked.error);
  if (!checked.value.isFile()) {
    return Result.err(
      cliError({ code: FOLIO_CLI_ERROR_CODES.invalidInput, message: `${real} is not a file.` }),
    );
  }
  const read = await Result.tryPromise({
    try: async (): Promise<OpenedFile | "swapped" | "tooLarge"> => {
      const handle = await open(real, constants.O_RDONLY | NO_FOLLOW);
      try {
        const opened = await handle.stat();
        const identity = { dev: opened.dev, ino: opened.ino };
        if (!opened.isFile() || !sameFile(identity, checked.value)) return "swapped";
        if (opened.size > MAX_DOCUMENT_BYTES) return "tooLarge";
        const bytes = new Uint8Array(await handle.readFile());
        return { bytes, identity, links: opened.nlink };
      } finally {
        await handle.close();
      }
    },
    catch: (error) => fileSystemError(real, error),
  });
  if (read.isErr()) return Result.err(read.error);
  if (read.value === "swapped") {
    return Result.err(
      cliError({
        code: FOLIO_CLI_ERROR_CODES.invalidInput,
        message: `${real} changed while it was being opened.`,
        hint: "Retry once nothing else is replacing the file.",
      }),
    );
  }
  if (read.value === "tooLarge") {
    return Result.err(
      cliError({
        code: FOLIO_CLI_ERROR_CODES.tooLarge,
        message: `${real} is over the ${MAX_DOCUMENT_BYTES}-byte limit.`,
      }),
    );
  }
  return Result.ok(read.value);
};

/** Read one input file, refusing directories and files over the size bound. */
export const readDocumentFile = async (
  inputPath: string,
): Promise<Result<LoadedFile, FolioCliError>> => {
  const absolute = path.resolve(inputPath);
  const real = await Result.tryPromise({
    try: () => realpath(absolute),
    catch: (error) => fileSystemError(absolute, error),
  });
  if (real.isErr()) return Result.err(real.error);
  const opened = await readThroughHandle(real.value);
  if (opened.isErr()) return Result.err(opened.error);
  const { bytes, identity, links } = opened.value;
  return Result.ok({
    path: real.value,
    bytes,
    fileVersion: fileVersionOf(bytes),
    identity,
    links,
  });
};

/**
 * Parse a loaded file into a headless reviewer. `author` is required for a
 * reviewer that will author changes; reads pass none and never write.
 */
export const openReviewer = async (
  file: LoadedFile,
  author?: string,
): Promise<Result<FolioDocxReviewer, FolioCliError>> =>
  await Result.tryPromise({
    try: () =>
      FolioDocxReviewer.fromBuffer(file.bytes.slice().buffer, {
        ...(author !== undefined && { author }),
      }),
    catch: (error) =>
      cliError({
        code: FOLIO_CLI_ERROR_CODES.invalidDocument,
        message: `${file.path} is not a readable .docx package: ${describeError(error)}`,
      }),
  });

/** Refuse when the caller's expected version is not the file's current one. */
export const checkExpectedVersion = (
  file: LoadedFile,
  expected: string | undefined,
): Result<void, FolioCliError> => {
  if (expected === undefined || expected === file.fileVersion) {
    return Result.ok();
  }
  return Result.err(
    cliError({
      code: FOLIO_CLI_ERROR_CODES.staleVersion,
      message: `${file.path} changed since it was read: expected version ${expected}, found ${file.fileVersion}.`,
      hint: "Re-read the document and retry against its current fileVersion.",
      details: { expected, actual: file.fileVersion },
    }),
  );
};
