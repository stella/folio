import { Result } from "better-result";
import { constants } from "node:fs";
import { open, realpath, stat } from "node:fs/promises";
import path from "node:path";

import { ensureParaIds, FolioDocxReviewer } from "@stll/folio-core/server";

import { cliError, FOLIO_CLI_ERROR_CODES, type FolioCliError } from "./errors";
import { errnoCode, fileVersionOf, NO_FOLLOW, sameFile, type FileIdentity } from "./file-system";
import { checkWordprocessingPackage } from "./main-document-part";

export { errnoCode, fileVersionOf, NO_FOLLOW, sameFile, type FileIdentity } from "./file-system";

/** Largest `.docx` the CLI opens; larger inputs are refused before reading. */
export const MAX_DOCUMENT_BYTES = 64 * 1024 * 1024;

const FILE_VERSION_PATTERN = /^[0-9a-f]{64}$/u;

export const isFileVersion = (value: unknown): value is string =>
  typeof value === "string" && FILE_VERSION_PATTERN.test(value);

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

/**
 * Read one input file, refusing directories, files over the size bound, and
 * anything that is not a WordprocessingML package. Every command and tool
 * that takes a document reads it here, so none can treat an archive without
 * a main document part as an empty document.
 */
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
  const checked = await checkWordprocessingPackage(real.value, bytes);
  if (checked.isErr()) return Result.err(checked.error);
  return Result.ok({
    path: real.value,
    bytes,
    fileVersion: fileVersionOf(bytes),
    identity,
    links,
  });
};

/**
 * Where a block id comes from. `package` ids are the paragraph's own
 * `w14:paraId` in the file; `synthetic` ids were minted for a paragraph the
 * file gives none, and are valid for the fileVersion they were read at. The
 * first change written to such a file stores them, so from then on they are
 * the package's own.
 */
export type FolioBlockIdSource = "package" | "synthetic";

/** A reviewer over one file and which of its block ids the file itself carries. */
export type OpenedDocument = {
  reviewer: FolioDocxReviewer;
  /** Ids minted for paragraphs without one; `null` when none could be minted up front. */
  mintedIds: ReadonlySet<string> | null;
};

/**
 * Give every paragraph a `w14:paraId`, deterministically from the file's
 * bytes, so the ids a read reports are the ids a later change writes into the
 * file: a paragraph edited by one call is found by the same id in the next.
 * A package the pass cannot or may not rewrite (a signed one) is opened as it
 * is, with position-derived ids.
 */
const withParagraphIds = async (
  bytes: Uint8Array<ArrayBuffer>,
): Promise<{ bytes: Uint8Array; mintedIds: ReadonlySet<string> | null }> => {
  const normalized = await Result.tryPromise(() => ensureParaIds(bytes));
  return normalized.isOk()
    ? { bytes: normalized.value.docx, mintedIds: new Set(normalized.value.mintedParaIds) }
    : { bytes, mintedIds: null };
};

/**
 * Parse a loaded file into a headless reviewer. `author` is required for a
 * reviewer that will author changes; reads pass none and never write.
 */
export const openReviewer = async (
  file: LoadedFile,
  author?: string,
): Promise<Result<OpenedDocument, FolioCliError>> => {
  const { bytes, mintedIds } = await withParagraphIds(file.bytes);
  const reviewer = await Result.tryPromise({
    try: () =>
      FolioDocxReviewer.fromBuffer(bytes.slice().buffer, {
        ...(author !== undefined && { author }),
      }),
    catch: (error) =>
      cliError({
        code: FOLIO_CLI_ERROR_CODES.invalidDocument,
        message: `${file.path} is not a readable .docx package: ${describeError(error)}`,
      }),
  });
  return reviewer.isErr()
    ? Result.err(reviewer.error)
    : Result.ok({ reviewer: reviewer.value, mintedIds });
};

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
