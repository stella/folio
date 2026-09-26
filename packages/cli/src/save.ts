/**
 * `folio save`: commit a whole package an editor serialized (a selective
 * patch where the editor could, a full repack otherwise) with the same
 * transaction as a tool call: under the lease, after crash recovery, against
 * the version the editor loaded, staged, checked, backed up, journaled
 * (`tool: "editor_save"`), and renamed into place.
 *
 * Unlike a tool call, a save may replace every part: the editor has to be
 * able to save any edit, and the backup and journal line keep the previous
 * version recoverable; the host warns the user. It is not an MCP tool, so
 * an agent never replaces a whole package.
 */

import { Result } from "better-result";
import { createHash, randomUUID } from "node:crypto";
import { parseArgs } from "node:util";

import { FolioDocxReviewer } from "@stll/folio-core/server";

import type { FolioCliIo } from "./cli";
import {
  checkExpectedVersion,
  fileVersionOf,
  isFileVersion,
  MAX_DOCUMENT_BYTES,
  readDocumentFile,
  sameFile,
} from "./document";
import { acquireLeaseForWrite, type FlushOutcome } from "./editor-lease";
import {
  cliError,
  EXIT_CODES,
  exitCodeForError,
  FOLIO_CLI_ERROR_CODES,
  type FolioCliError,
} from "./errors";
import {
  canonicalJson,
  EDITOR_SAVE_TOOL,
  resolveTarget,
  TX_ID_PATTERN,
  type WriteDestination,
} from "./execute-write";
import { findCommit, journalPathFor, recoverStages } from "./journal";
import { adoptLease, DEFAULT_LEASE_OWNER, isLeaseOwner, type AcquiredLease } from "./lock";
import {
  failureEnvelope,
  isOutputFormat,
  OUTPUT_FORMATS,
  renderEnvelope,
  successEnvelope,
  type OutputFormat,
} from "./output";
import { diffPackages } from "./package-parts";
import { resolveAuthor, resolveTransactionDate } from "./provenance";
import { unsafePath } from "./sidecar";
import { commitTransaction } from "./transaction";

/** How the editor produced the bytes, as it reports it. */
export const EDITOR_SAVE_STRATEGIES = ["selective", "full-repack", "unspecified"] as const;

export type EditorSaveStrategy = (typeof EDITOR_SAVE_STRATEGIES)[number];

const isEditorSaveStrategy = (value: unknown): value is EditorSaveStrategy =>
  EDITOR_SAVE_STRATEGIES.some((strategy) => strategy === value);

export type SaveDocumentOptions = {
  /** The document the editor opened. */
  path: string;
  /** The whole package the editor serialized. */
  bytes: Uint8Array;
  /** The fileVersion the editor loaded (or last saved). */
  expectedVersion: string;
  author: string;
  /** The lease owner a transient lease records (default `folio-cli`). */
  owner?: string;
  /** Which editor saved, for the journal (`vscode`, `herdr`; default `cli`). */
  surface?: string;
  saveStrategy?: EditorSaveStrategy;
  /** Idempotency key; generated when absent. */
  txId?: string;
  /** After waiting for another editor to flush, take the lease over. */
  force?: boolean;
  /** Default in place. */
  destination?: WriteDestination;
  /** Save under a lease the caller holds (an editor's long-lived lease) instead of taking one. */
  leaseToken?: string;
  /** How long to wait for another editor holding the lease to save and release. */
  flushWaitMs?: number;
  /** The transaction's UTC timestamp (default now). */
  date?: string;
};

const invalidInput = (message: string, hint?: string): FolioCliError =>
  cliError({ code: FOLIO_CLI_ERROR_CODES.invalidInput, message, hint });

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

const NAME_PATTERN = /^[\w.-]{1,64}$/u;

const checkOptions = (options: SaveDocumentOptions): Result<void, FolioCliError> => {
  if (options.txId !== undefined && !TX_ID_PATTERN.test(options.txId)) {
    return Result.err(invalidInput("txId must be 1 to 128 letters, digits, '.', '_' or '-'."));
  }
  if (options.owner !== undefined && !isLeaseOwner(options.owner)) {
    return Result.err(invalidInput("owner must be 1 to 64 letters, digits, '.', '_' or '-'."));
  }
  if (options.surface !== undefined && !NAME_PATTERN.test(options.surface)) {
    return Result.err(invalidInput("surface must be 1 to 64 letters, digits, '.', '_' or '-'."));
  }
  if (!isFileVersion(options.expectedVersion)) {
    return Result.err(invalidInput("expectedVersion must be a 64-character hex fileVersion."));
  }
  if (options.bytes.byteLength > MAX_DOCUMENT_BYTES) {
    return Result.err(
      cliError({
        code: FOLIO_CLI_ERROR_CODES.tooLarge,
        message: `The saved package is over the ${MAX_DOCUMENT_BYTES}-byte limit.`,
      }),
    );
  }
  return Result.ok();
};

/** The bytes must open as a document before anything is staged. */
const checkParses = async (bytes: Uint8Array): Promise<Result<void, FolioCliError>> => {
  const parsed = await Result.tryPromise({
    try: () => FolioDocxReviewer.fromBuffer(bytes.slice().buffer),
    catch: (error) =>
      cliError({
        code: FOLIO_CLI_ERROR_CODES.invalidDocument,
        message: `The saved bytes are not a readable .docx package: ${error instanceof Error ? error.message : String(error)}`,
        hint: "Nothing was written; the editor's serializer produced a broken package.",
      }),
  });
  return parsed.isErr() ? Result.err(parsed.error) : Result.ok();
};

type SaveLease = { lease: AcquiredLease; flush: FlushOutcome };

const leaseForSave = async (
  documentPath: string,
  txId: string,
  options: SaveDocumentOptions,
): Promise<Result<SaveLease, FolioCliError>> => {
  if (options.leaseToken !== undefined) {
    const adopted = await adoptLease(documentPath, options.leaseToken);
    return adopted.isOk()
      ? Result.ok({ lease: adopted.value, flush: { type: "notAsked" } })
      : Result.err(adopted.error);
  }
  return await acquireLeaseForWrite({
    documentPath,
    txId,
    force: options.force === true,
    owner: options.owner ?? DEFAULT_LEASE_OWNER,
    ...(options.flushWaitMs !== undefined && { flushWaitMs: options.flushWaitMs }),
  });
};

/**
 * Commit an editor's serialized package as one transaction and return its
 * receipt (`fileVersion`, `backup`, `changedParts`, ...). Refuses a stale
 * `expectedVersion`, a lease another live holder keeps, bytes that do not
 * open as a document, and anything over 64 MiB, writing nothing.
 */
export const saveDocumentBytes = async (
  options: SaveDocumentOptions,
): Promise<Result<Readonly<Record<string, unknown>>, FolioCliError>> => {
  const checked = checkOptions(options);
  if (checked.isErr()) return Result.err(checked.error);
  const date = resolveTransactionDate(options.date);
  if (date.isErr()) return Result.err(date.error);
  const destination: WriteDestination = options.destination ?? { type: "inPlace" };
  const txId = options.txId ?? randomUUID();
  const surface = options.surface ?? "cli";
  const saveStrategy = options.saveStrategy ?? "unspecified";
  const owner = options.owner ?? DEFAULT_LEASE_OWNER;
  const bytes = new Uint8Array(options.bytes);

  const parses = await checkParses(bytes);
  if (parses.isErr()) return Result.err(parses.error);
  const target = await resolveTarget({ sourcePath: options.path, destination });
  if (target.isErr()) return Result.err(target.error);
  const { sourcePath, sourceIdentity, destinationPath } = target.value;
  const inPlace = destinationPath === sourcePath;
  const journalPath = journalPathFor(destinationPath, undefined);

  const acquired = await leaseForSave(destinationPath, txId, options);
  if (acquired.isErr()) return Result.err(acquired.error);
  const { lease, flush } = acquired.value;
  try {
    const recovered = await recoverStages({
      documentPath: destinationPath,
      journalPath,
      now: date.value,
    });
    if (recovered.isErr()) return Result.err(recovered.error);

    const toVersion = fileVersionOf(bytes);
    const requestHash = createHash("sha256")
      .update(
        canonicalJson({
          tool: EDITOR_SAVE_TOOL,
          toVersion,
          source: sourcePath,
          destination: destinationPath,
        }),
      )
      .digest("hex");
    if (options.txId !== undefined) {
      const found = await findCommit(journalPath, txId);
      if (found.isErr()) return Result.err(found.error);
      const prior = found.value;
      if (prior !== undefined) {
        return prior.requestHash === requestHash
          ? Result.ok({ ...prior.receipt, status: "replayed" })
          : Result.err(
              cliError({
                code: FOLIO_CLI_ERROR_CODES.transactionConflict,
                message: `Transaction ${txId} already committed a different save.`,
                hint: "Use a new txId for a new save.",
              }),
            );
      }
    }

    const source = await readDocumentFile(sourcePath);
    if (source.isErr()) return Result.err(source.error);
    const replacedInPlace =
      inPlace &&
      (flush.type !== "notAsked" ||
        recovered.value.some(({ action }) => action === "rolledForward"));
    const expectedIdentity = replacedInPlace ? source.value.identity : sourceIdentity;
    if (!sameFile(source.value.identity, expectedIdentity)) {
      return Result.err(unsafePath(`${sourcePath} was replaced while the save started.`));
    }
    if (inPlace && source.value.links > 1) {
      return Result.err(
        unsafePath(
          `${sourcePath} has other hard links; an in-place save would detach them.`,
          "Save to a new file with -o instead.",
        ),
      );
    }
    const version = checkExpectedVersion(source.value, options.expectedVersion);
    if (version.isErr()) {
      return Result.err(
        cliError({
          code: version.error.code,
          message: version.error.message,
          hint: "Reload the document (it changed on disk) and save again.",
          details: {
            ...(isRecord(version.error.details) ? version.error.details : {}),
            ...(recovered.value.length > 0 && { recovered: recovered.value }),
          },
        }),
      );
    }

    const base = {
      txId,
      tool: EDITOR_SAVE_TOOL,
      path: destinationPath,
      fromVersion: source.value.fileVersion,
      ...(!inPlace && { source: { path: sourcePath, fileVersion: source.value.fileVersion } }),
      author: options.author,
      time: date.value,
      owner,
      surface,
      saveStrategy,
    };
    if (inPlace && toVersion === source.value.fileVersion) {
      return Result.ok({ ...base, status: "unchanged", fileVersion: toVersion, changedParts: [] });
    }
    const changedParts = await diffPackages(source.value.bytes, bytes);
    if (changedParts.isErr()) return Result.err(changedParts.error);

    const receipt = {
      ...base,
      status: "committed",
      changedParts: changedParts.value,
      ...(recovered.value.length > 0 && { recovered: recovered.value }),
    };
    const committed = await commitTransaction({
      txId,
      destinationPath,
      destination,
      sourcePath,
      sourceIdentity: expectedIdentity,
      fromVersion: source.value.fileVersion,
      bytes,
      verifyLease: lease.verify,
      changedParts: changedParts.value,
      journalPath,
      entry: {
        txId,
        requestHash,
        tool: EDITOR_SAVE_TOOL,
        fromVersion: source.value.fileVersion,
        author: options.author,
        time: date.value,
        ops: { surface, saveStrategy, owner },
        receipts: [],
      },
      receipt,
    });
    return committed.isErr() ? Result.err(committed.error) : Result.ok(committed.value.receipt);
  } finally {
    await lease.release();
  }
};

const usageError = (message: string, hint?: string): FolioCliError =>
  cliError({ code: FOLIO_CLI_ERROR_CODES.usage, message, hint });

export const SAVE_HELP = [
  "Usage: folio save <file> --from <saved.docx> --expect-version <sha> [flags]",
  "",
  "Commit a whole package an editor serialized as a folio transaction: under the write",
  "lease, against the version the editor loaded, checked, backed up to .folio/backups,",
  'journaled (tool "editor_save"), and renamed into place atomically. Unlike the tool',
  "commands it may replace every part. It is not offered over MCP.",
  "",
  "An editor holding the lease long-lived saves under it with --lease-token. Any other",
  "save that finds such an editor asks it to save and release first, waiting up to",
  "--flush-wait ms; its own --expect-version then no longer matches and it is refused.",
  "",
  "Flags:",
  "  --from <path>                       The package to commit (required)",
  "  --expect-version <sha>              The fileVersion the editor loaded (required)",
  "  -o, --out <path>                    Save to this path instead of in place",
  "  --overwrite                         Let -o replace an existing file",
  "  --expect-destination-version <sha>  The fileVersion of the file -o --overwrite replaces",
  "  --author <name>                     Author (default: FOLIO_AUTHOR, then git user.name)",
  "  --owner <name>                      Lease owner to record (default: folio-cli)",
  "  --surface <name>                    Which editor saved, for the journal (default: cli)",
  `  --save-strategy <kind>              ${EDITOR_SAVE_STRATEGIES.join(", ")}`,
  "  --lease-token <token>               Save under this held lease instead of taking one",
  "  --flush-wait <ms>                   Wait for another editor to flush (default 5000)",
  "  --tx-id <key>                       Idempotency key",
  "  --force                             After the wait, take the lease over",
  "  --date <iso>                        Transaction timestamp (default: now)",
  "  --output <json|text>                Envelope format",
  "",
].join("\n");

const saveOptions = {
  from: { type: "string" },
  "expect-version": { type: "string" },
  out: { type: "string", short: "o" },
  overwrite: { type: "boolean" },
  "expect-destination-version": { type: "string" },
  author: { type: "string" },
  owner: { type: "string" },
  surface: { type: "string" },
  "save-strategy": { type: "string" },
  "lease-token": { type: "string" },
  "flush-wait": { type: "string" },
  "tx-id": { type: "string" },
  force: { type: "boolean" },
  date: { type: "string" },
  output: { type: "string" },
  help: { type: "boolean", short: "h" },
} as const;

/** `--flush-wait <ms>`: a whole number of milliseconds from 0. */
export const parseFlushWait = (
  value: string | undefined,
): Result<number | undefined, FolioCliError> => {
  if (value === undefined) return Result.ok(undefined);
  const parsed = Number(value);
  return Number.isInteger(parsed) && parsed >= 0
    ? Result.ok(parsed)
    : Result.err(usageError("--flush-wait expects a whole number of milliseconds."));
};

const parseSaveArgs = (rest: readonly string[]) =>
  parseArgs({ args: [...rest], options: saveOptions, allowPositionals: true, strict: true });

type SaveValues = ReturnType<typeof parseSaveArgs>["values"];

const runSaveValues = async (
  values: SaveValues,
  positionals: readonly string[],
  io: FolioCliIo,
): Promise<Result<unknown, FolioCliError>> => {
  const [file, ...extra] = positionals;
  if (file === undefined || extra.length > 0) {
    return Result.err(
      usageError("Usage: folio save <file> --from <saved.docx> --expect-version <sha>."),
    );
  }
  if (values.from === undefined) {
    return Result.err(usageError("folio save needs --from <saved.docx>."));
  }
  if (values["expect-version"] === undefined) {
    return Result.err(
      usageError("folio save needs --expect-version <fileVersion> of the file the editor loaded."),
    );
  }
  const saveStrategy = values["save-strategy"];
  if (saveStrategy !== undefined && !isEditorSaveStrategy(saveStrategy)) {
    return Result.err(usageError(`--save-strategy must be ${EDITOR_SAVE_STRATEGIES.join(", ")}.`));
  }
  const expectedDestination = values["expect-destination-version"];
  if (
    (values.overwrite === true || expectedDestination !== undefined) &&
    values.out === undefined
  ) {
    return Result.err(
      usageError("--overwrite and --expect-destination-version only apply to -o <path>."),
    );
  }
  const flushWaitMs = parseFlushWait(values["flush-wait"]);
  if (flushWaitMs.isErr()) return Result.err(flushWaitMs.error);
  const author = resolveAuthor({ explicit: values.author, env: io.env, cwd: io.cwd });
  if (author.isErr()) return Result.err(author.error);
  const saved = await readDocumentFile(values.from);
  if (saved.isErr()) return Result.err(saved.error);
  return await saveDocumentBytes({
    path: file,
    bytes: saved.value.bytes,
    expectedVersion: values["expect-version"],
    author: author.value,
    ...(values.owner !== undefined && { owner: values.owner }),
    ...(values.surface !== undefined && { surface: values.surface }),
    ...(saveStrategy !== undefined && { saveStrategy }),
    ...(values["tx-id"] !== undefined && { txId: values["tx-id"] }),
    ...(values["lease-token"] !== undefined && { leaseToken: values["lease-token"] }),
    ...(flushWaitMs.value !== undefined && { flushWaitMs: flushWaitMs.value }),
    ...(values.date !== undefined && { date: values.date }),
    force: values.force === true,
    destination:
      values.out === undefined
        ? { type: "inPlace" }
        : {
            type: "file",
            path: values.out,
            overwrite: values.overwrite === true,
            expectedVersion: expectedDestination,
          },
  });
};

const emit = (
  io: FolioCliIo,
  format: OutputFormat,
  result: Result<unknown, FolioCliError>,
): number => {
  const envelope = result.isOk() ? successEnvelope(result.value) : failureEnvelope(result.error);
  const rendered = renderEnvelope({ envelope, format, tool: EDITOR_SAVE_TOOL });
  if (rendered.stdout !== "") io.stdout(rendered.stdout);
  if (rendered.stderr !== "") io.stderr(rendered.stderr);
  return result.isOk() ? EXIT_CODES.ok : exitCodeForError(result.error.code);
};

export const runSave = async (rest: readonly string[], io: FolioCliIo): Promise<number> => {
  const fallback: OutputFormat = io.isTTY ? "text" : "json";
  const parsed = Result.try({
    try: () => parseSaveArgs(rest),
    catch: (error) =>
      usageError(error instanceof Error ? error.message : String(error), "Run folio save --help."),
  });
  if (parsed.isErr()) return emit(io, fallback, parsed);
  const { values, positionals } = parsed.value;
  if (values.help === true) {
    io.stdout(SAVE_HELP);
    return EXIT_CODES.ok;
  }
  const output = values.output;
  if (output !== undefined && !isOutputFormat(output)) {
    return emit(
      io,
      fallback,
      Result.err(usageError(`--output must be ${OUTPUT_FORMATS.join(" or ")}.`)),
    );
  }
  return emit(io, output ?? fallback, await runSaveValues(values, positionals, io));
};
