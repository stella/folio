import { Result } from "better-result";
import { createHash } from "node:crypto";
import { readFile, stat } from "node:fs/promises";
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

/** The bytes of one input file and the version they hash to. */
export type LoadedFile = {
  /** Absolute path the bytes were read from. */
  path: string;
  bytes: Uint8Array<ArrayBuffer>;
  fileVersion: string;
};

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

/** Read one input file, refusing directories and files over the size bound. */
export const readDocumentFile = async (
  inputPath: string,
): Promise<Result<LoadedFile, FolioCliError>> => {
  const absolute = path.resolve(inputPath);
  const info = await Result.tryPromise({
    try: () => stat(absolute),
    catch: (error) => fileSystemError(absolute, error),
  });
  if (info.isErr()) {
    return Result.err(info.error);
  }
  if (!info.value.isFile()) {
    return Result.err(
      cliError({ code: FOLIO_CLI_ERROR_CODES.invalidInput, message: `${absolute} is not a file.` }),
    );
  }
  if (info.value.size > MAX_DOCUMENT_BYTES) {
    return Result.err(
      cliError({
        code: FOLIO_CLI_ERROR_CODES.tooLarge,
        message: `${absolute} is ${info.value.size} bytes, over the ${MAX_DOCUMENT_BYTES}-byte limit.`,
      }),
    );
  }
  const bytes = await Result.tryPromise({
    try: async () => new Uint8Array(await readFile(absolute)),
    catch: (error) => fileSystemError(absolute, error),
  });
  if (bytes.isErr()) {
    return Result.err(bytes.error);
  }
  return Result.ok({ path: absolute, bytes: bytes.value, fileVersion: fileVersionOf(bytes.value) });
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
