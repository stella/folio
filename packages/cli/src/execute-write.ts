/**
 * Run one mutating tool call as a transaction: take the destination's lease
 * (asking an editor that holds it to save and release first), settle any
 * crashed stage, replay a committed `txId`, apply the change to a fresh
 * reviewer, save it (selective unless a full repack is allowed), and commit
 * the bytes atomically with a journal line and, in place, a backup.
 */

import { panic, Result } from "better-result";
import { createHash, randomUUID } from "node:crypto";
import { realpath, stat } from "node:fs/promises";
import path from "node:path";

import { createReviewerBridge } from "@stll/folio-agents/bridges/reviewer";
import { generateRedlineDocx } from "@stll/folio-agents/compare";
import { executeFolioToolCallUntyped } from "@stll/folio-agents/execute";
import { FOLIO_AGENT_TOOL_NAMES } from "@stll/folio-agents/types";
import type { FolioAIEditApplyMode, FolioDocxReviewer } from "@stll/folio-core/server";

import {
  checkExpectedVersion,
  errnoCode,
  openReviewer,
  readDocumentFile,
  sameFile,
  type FileIdentity,
  type LoadedFile,
} from "./document";
import { cliError, FOLIO_CLI_ERROR_CODES, type FolioCliError } from "./errors";
import type { FileToolCall } from "./execute-read";
import { acquireLeaseForWrite, type FlushOutcome } from "./editor-lease";
import { findCommit, journalPathFor, latestCommitFor, recoverStages } from "./journal";
import { diffPackages, nextRevisionIdSeed } from "./package-parts";
import type { FolioFileToolSpec, ResolveChangeAction } from "./registry";
import { SIDECAR_DIRECTORY, unsafePath } from "./sidecar";
import { commitTransaction, type WriteDestination } from "./transaction";

export type { WriteDestination } from "./transaction";

/** Everything about a write that is not the tool's own arguments. */
export type WriteOptions = {
  destination: WriteDestination;
  /** Already resolved: a write never proceeds without one. */
  author: string;
  /** The transaction's one UTC timestamp. */
  date: string;
  /** Whether a save the selective patch cannot express may rewrite every part. */
  repack: "allow" | "refuse";
  /** Take the lease over from another live holder. */
  force: boolean;
  /**
   * `required`: the call must name the source's fileVersion (every MCP
   * write, and the CLI unless `--no-expect-version`). `waived`: explicitly
   * skipped.
   */
  sourcePrecondition: "required" | "waived";
  /** Idempotency key; generated when absent. */
  txId: string | undefined;
  journalPath: string | undefined;
  /** For tools whose edit mode is selectable. */
  mode: FolioAIEditApplyMode;
  /** How long to wait for an editor holding the lease to save and release (`editor-lease.ts`). */
  flushWaitMs?: number;
};

/** How a transaction's bytes were produced. */
export type SaveStrategy = "selective" | "full-repack" | "redline";

export const TX_ID_PATTERN = /^[\w.-]{1,128}$/u;

const invalidInput = (message: string, hint?: string): FolioCliError =>
  cliError({ code: FOLIO_CLI_ERROR_CODES.invalidInput, message, hint });

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

/** JSON with object keys sorted at every depth, for a stable request hash. */
export const canonicalJson = (value: unknown): string => {
  if (Array.isArray(value)) return `[${value.map(canonicalJson).join(",")}]`;
  if (isRecord(value)) {
    const keys = Object.keys(value).toSorted();
    return `{${keys.map((key) => `${JSON.stringify(key)}:${canonicalJson(value[key])}`).join(",")}}`;
  }
  return JSON.stringify(value) ?? "null";
};

const STALE_ISSUE_CODES: ReadonlySet<string> = new Set([
  "missingBlock",
  "changedBlock",
  "staleRange",
  "preconditionFailed",
  "documentVersionMismatch",
]);

/**
 * A batch lands whole or not at all: any skipped operation refuses the
 * transaction. Stale targets are a conflict with current state; an ambiguous
 * `find` or an operation the document cannot take is an input error.
 */
const refusalFor = (summary: unknown): FolioCliError | null => {
  if (!isRecord(summary) || !Array.isArray(summary["skipped"])) return null;
  if (summary["skipped"].length === 0) return null;
  const issues: unknown[] = Array.isArray(summary["issues"]) ? summary["issues"] : [];
  const codes = issues
    .map((issue) => (isRecord(issue) ? issue["code"] : undefined))
    .filter((code) => typeof code === "string" && code !== "atomicBatchRejected");
  const details = { skipped: summary["skipped"], issues };
  if (codes.some((code) => typeof code === "string" && STALE_ISSUE_CODES.has(code))) {
    return cliError({
      code: FOLIO_CLI_ERROR_CODES.staleTarget,
      message: "An operation's target changed or no longer exists; nothing was written.",
      hint: "Re-read the document (read, find) and retry with fresh ids, ranges, and fileVersion.",
      details,
    });
  }
  if (codes.includes("ambiguousFind")) {
    return cliError({
      code: FOLIO_CLI_ERROR_CODES.ambiguousTarget,
      message: "An operation's find text matches more than once in its block; nothing was written.",
      hint: "Make the find text unique within the block, or target a range returned by find.",
      details,
    });
  }
  return cliError({
    code: FOLIO_CLI_ERROR_CODES.operationRejected,
    message: "The document refused an operation; nothing was written.",
    hint: "See error.details.skipped for each operation's reason.",
    details,
  });
};

/** What a mutation produced, before it is committed. */
type Mutation = {
  bytes: Uint8Array<ArrayBuffer>;
  saveStrategy: SaveStrategy;
  repackReason?: string;
  result: unknown;
  receipts: readonly unknown[];
  /** Other inputs the transaction read, reported in the receipt. */
  inputs?: Readonly<Record<string, { path: string; fileVersion: string }>>;
};

type MutateOptions = {
  tool: FolioFileToolSpec;
  source: LoadedFile;
  args: Readonly<Record<string, unknown>>;
  options: WriteOptions;
};

const saveReviewer = async (
  reviewer: FolioDocxReviewer,
  repack: WriteOptions["repack"],
): Promise<Result<Pick<Mutation, "bytes" | "saveStrategy" | "repackReason">, FolioCliError>> => {
  const saved = await Result.tryPromise({
    try: () => reviewer.save({ repack }),
    catch: (error) =>
      cliError({
        code: FOLIO_CLI_ERROR_CODES.internal,
        message: `The document could not be saved: ${error instanceof Error ? error.message : String(error)}`,
      }),
  });
  if (saved.isErr()) return Result.err(saved.error);
  switch (saved.value.type) {
    case "selective":
      return Result.ok({ bytes: new Uint8Array(saved.value.buffer), saveStrategy: "selective" });
    case "full-repack":
      return Result.ok({
        bytes: new Uint8Array(saved.value.buffer),
        saveStrategy: "full-repack",
        repackReason: saved.value.reason,
      });
    case "repackRefused":
      return Result.err(
        cliError({
          code: FOLIO_CLI_ERROR_CODES.repackRequired,
          message: `This change needs a full repack of the package (${saved.value.reason}); nothing was written.`,
          hint: "Pass --allow-repack (MCP: allowRepack: true) to rewrite every part.",
          details: { reason: saved.value.reason },
        }),
      );
    default: {
      const unreachable: never = saved.value;
      return panic("Unhandled save result", { unreachable });
    }
  }
};

const commentExists = (reviewer: FolioDocxReviewer, commentId: unknown): boolean =>
  reviewer
    .getComments()
    .some(
      (thread) =>
        String(thread.id) === commentId ||
        thread.replies.some((reply) => String(reply.id) === commentId),
    );

const mutateWithAgentTool = async ({
  tool,
  source,
  args,
  options,
}: MutateOptions): Promise<Result<Mutation, FolioCliError>> => {
  if (tool.type !== "agentWrite") return panic("mutateWithAgentTool needs an agentWrite tool");
  const reviewer = await openReviewer(source, options.author);
  if (reviewer.isErr()) return Result.err(reviewer.error);
  const idSeed = await nextRevisionIdSeed(source.bytes);
  if (idSeed.isErr()) return Result.err(idSeed.error);

  const commentTool =
    tool.agentTool === FOLIO_AGENT_TOOL_NAMES.replyComment ||
    tool.agentTool === FOLIO_AGENT_TOOL_NAMES.resolveComment;
  if (commentTool && !commentExists(reviewer.value, args["commentId"])) {
    return Result.err(
      cliError({
        code: FOLIO_CLI_ERROR_CODES.notFound,
        message: `No comment with id ${JSON.stringify(args["commentId"])}.`,
        hint: "Take the id from folio comments.",
      }),
    );
  }

  const bridge = createReviewerBridge(reviewer.value, {
    mode: tool.editMode === "tracked-or-direct" ? options.mode : "tracked-changes",
    revisionStamp: { date: options.date, idSeed: idSeed.value },
  });
  const executed = executeFolioToolCallUntyped(tool.agentTool, args, bridge);
  if (!executed.ok) return Result.err(invalidInput(executed.error));
  const refusal = refusalFor(executed.result);
  if (refusal !== null) return Result.err(refusal);

  const saved = await saveReviewer(reviewer.value, options.repack);
  if (saved.isErr()) return Result.err(saved.error);
  const receipts =
    isRecord(executed.result) && Array.isArray(executed.result["receipts"])
      ? executed.result["receipts"]
      : [];
  return Result.ok({ ...saved.value, result: executed.result, receipts });
};

const isResolveAction = (value: unknown): value is ResolveChangeAction =>
  value === "accept" || value === "reject";

const resolveChanges = async ({
  source,
  args,
  options,
}: MutateOptions): Promise<Result<Mutation, FolioCliError>> => {
  const { action, ids, all } = args;
  if (!isResolveAction(action)) {
    return Result.err(invalidInput('resolve_changes needs action "accept" or "reject".'));
  }
  const idList = Array.isArray(ids) && ids.every((id) => typeof id === "string") ? ids : undefined;
  if ((idList !== undefined && idList.length > 0) === (all === true)) {
    return Result.err(
      invalidInput(
        "Pass change ids or all, exactly one of them.",
        "folio accept --id 12 or --all.",
      ),
    );
  }
  const reviewer = await openReviewer(source, options.author);
  if (reviewer.isErr()) return Result.err(reviewer.error);

  let resolved: number | readonly string[];
  if (all === true) {
    const count = action === "accept" ? reviewer.value.acceptAll() : reviewer.value.rejectAll();
    if (count === 0) {
      return Result.err(
        cliError({
          code: FOLIO_CLI_ERROR_CODES.operationRejected,
          message: "The document has no pending tracked changes; nothing was written.",
        }),
      );
    }
    resolved = count;
  } else {
    const selected = idList ?? [];
    const known = new Set(reviewer.value.getChanges().map(({ id }) => String(id)));
    const missing = selected.filter((id) => !known.has(id));
    if (missing.length > 0) {
      return Result.err(
        cliError({
          code: FOLIO_CLI_ERROR_CODES.notFound,
          message: `No pending change with id ${missing.join(", ")}.`,
          hint: "Take ids from folio changes.",
          details: { missing },
        }),
      );
    }
    for (const id of selected) {
      const changed =
        action === "accept"
          ? reviewer.value.acceptChange(Number(id))
          : reviewer.value.rejectChange(Number(id));
      if (!changed) {
        return Result.err(
          cliError({
            code: FOLIO_CLI_ERROR_CODES.operationRejected,
            message: `Change ${id} could not be ${action}ed; nothing was written.`,
          }),
        );
      }
    }
    resolved = selected;
  }
  const saved = await saveReviewer(reviewer.value, options.repack);
  if (saved.isErr()) return Result.err(saved.error);
  return Result.ok({
    ...saved.value,
    result: { action, resolved, remaining: reviewer.value.getChanges().length },
    receipts: [],
  });
};

const writeRedline = async ({
  source,
  args,
  options,
}: MutateOptions): Promise<Result<Mutation, FolioCliError>> => {
  const revisedPath = args["revisedPath"];
  if (typeof revisedPath !== "string" || revisedPath === "") {
    return Result.err(invalidInput("compare_documents needs revisedPath."));
  }
  const revised = await readDocumentFile(revisedPath);
  if (revised.isErr()) return Result.err(revised.error);
  const revisedVersion = args["revisedFileVersion"];
  const version = checkExpectedVersion(
    revised.value,
    typeof revisedVersion === "string" ? revisedVersion : undefined,
  );
  if (version.isErr()) return Result.err(version.error);

  const redline = await Result.tryPromise({
    try: () =>
      generateRedlineDocx(source.bytes.slice().buffer, revised.value.bytes.slice().buffer, {
        author: options.author,
      }),
    catch: (error) =>
      cliError({
        code: FOLIO_CLI_ERROR_CODES.invalidDocument,
        message: `The documents could not be compared: ${error instanceof Error ? error.message : String(error)}`,
      }),
  });
  if (redline.isErr()) return Result.err(redline.error);
  const { buffer, applied, skipped, unprocessedStories } = redline.value;
  return Result.ok({
    bytes: new Uint8Array(buffer),
    saveStrategy: "redline",
    result: { applied: applied.length, skipped, unprocessedStories },
    receipts: [],
    inputs: { revised: { path: revised.value.path, fileVersion: revised.value.fileVersion } },
  });
};

const mutate = (options: MutateOptions): Promise<Result<Mutation, FolioCliError>> => {
  switch (options.tool.type) {
    case "agentWrite":
      return mutateWithAgentTool(options);
    case "resolveChanges":
      return resolveChanges(options);
    case "compare":
      return writeRedline(options);
    case "agentRead":
      return panic("A read tool reached the write executor", { tool: options.tool.name });
    default: {
      const unreachable: never = options.tool;
      return panic("Unhandled tool type", { unreachable });
    }
  }
};

type ResolvedTarget = {
  sourcePath: string;
  sourceIdentity: FileIdentity;
  destinationPath: string;
};

/** The `tool` an editor's whole-package save journals (`save.ts`). */
export const EDITOR_SAVE_TOOL = "editor_save";

type RebaseCheck = {
  tool: FolioFileToolSpec;
  flush: FlushOutcome;
  inPlace: boolean;
  documentPath: string;
  expected: string | undefined;
  current: string;
};

/** How a write that waited for an editor's save came to run on the saved version. */
type Rebased = { fromVersion: string; toVersion: string; flushedBy: string };

/**
 * Whether a write whose `fileVersion` went stale only because the editor
 * holding the lease saved when asked may run on the saved version. It may
 * when it is an in-place agent tool (its targets are block ids, with
 * optional text-hash preconditions re-checked against the new version), the
 * version it read was current when the flush was requested, and the
 * journal's newest commit is the editor's save from exactly that version to
 * the file as it is now. Anything else stays `stale_version`.
 */
const rebasedOntoFlush = async ({
  tool,
  flush,
  inPlace,
  documentPath,
  expected,
  current,
}: RebaseCheck): Promise<Rebased | null> => {
  if (flush.type !== "flushed" || tool.type !== "agentWrite" || !inPlace) return null;
  if (expected === undefined || expected !== flush.versionBefore) return null;
  const latest = await latestCommitFor(documentPath);
  return latest !== null &&
    latest.tool === EDITOR_SAVE_TOOL &&
    latest.fromVersion === expected &&
    latest.toVersion === current
    ? { fromVersion: expected, toVersion: current, flushedBy: flush.holder.owner }
    : null;
};

const DOCX_NAME = /\.docx$/iu;

const invalidDestination = (message: string, hint?: string): FolioCliError =>
  cliError({ code: FOLIO_CLI_ERROR_CODES.invalidDestination, message, hint });

/**
 * A write lands only on a `.docx` whose name does not start with a dot and
 * that is not inside a `.folio` directory, so no tool call can replace a
 * dotfile, a sidecar, or another kind of file.
 */
const checkDestinationName = (candidate: string): Result<void, FolioCliError> => {
  const name = path.basename(candidate);
  if (!DOCX_NAME.test(name) || name.startsWith(".")) {
    return Result.err(
      invalidDestination(
        `${candidate} is not a writable destination: folio writes only .docx files whose name does not start with a dot.`,
      ),
    );
  }
  if (candidate.split(path.sep).includes(SIDECAR_DIRECTORY)) {
    return Result.err(
      invalidDestination(`${candidate} is inside a ${SIDECAR_DIRECTORY} directory.`),
    );
  }
  return Result.ok();
};

type ResolvedPath = { path: string; identity: FileIdentity | null };

const notFound = (message: string): FolioCliError =>
  cliError({ code: FOLIO_CLI_ERROR_CODES.notFound, message });

/**
 * The real path a write goes to: an existing file resolved through its
 * symlinks (it must be a regular file), or a new name in its directory's real
 * path. The name rules apply to the resolved path, so a link cannot redirect
 * a write onto a file the rules refuse.
 */
const resolveWritePath = async (
  requested: string,
): Promise<Result<ResolvedPath, FolioCliError>> => {
  const absolute = path.resolve(requested);
  const named = checkDestinationName(absolute);
  if (named.isErr()) return Result.err(named.error);
  const real = await Result.tryPromise(() => realpath(absolute));
  let resolved: ResolvedPath;
  if (real.isOk()) {
    const info = await Result.tryPromise(() => stat(real.value));
    if (info.isErr() || !info.value.isFile()) {
      return Result.err(invalidDestination(`${absolute} is not a regular file.`));
    }
    resolved = { path: real.value, identity: { dev: info.value.dev, ino: info.value.ino } };
  } else if (errnoCode(real.error.cause) === "ENOENT") {
    const parent = await Result.tryPromise(() => realpath(path.dirname(absolute)));
    if (parent.isErr()) return Result.err(notFound(`No directory ${path.dirname(absolute)}.`));
    resolved = { path: path.join(parent.value, path.basename(absolute)), identity: null };
  } else {
    return Result.err(invalidDestination(`Cannot resolve ${absolute}.`));
  }
  const resolvedName = checkDestinationName(resolved.path);
  return resolvedName.isErr() ? Result.err(resolvedName.error) : Result.ok(resolved);
};

const identityOfInput = async (
  inputPath: string,
): Promise<Result<{ path: string; identity: FileIdentity }, FolioCliError>> => {
  const real = await Result.tryPromise(() => realpath(path.resolve(inputPath)));
  if (real.isErr()) return Result.err(notFound(`No file at ${path.resolve(inputPath)}.`));
  const info = await Result.tryPromise(() => stat(real.value));
  if (info.isErr() || !info.value.isFile()) {
    return Result.err(invalidInput(`${real.value} is not a file.`));
  }
  return Result.ok({ path: real.value, identity: { dev: info.value.dev, ino: info.value.ino } });
};

type ResolveTargetOptions = {
  sourcePath: string;
  destination: WriteDestination;
  /** A redline: never in place, and never onto the revised file either. */
  redline?: { revisedPath: unknown };
};

/**
 * Resolve source and destination to real paths and refuse a destination
 * that is the input under another name (a symlink, a hard link, a
 * case-insensitive alias): files are compared by device and inode.
 */
export const resolveTarget = async ({
  sourcePath,
  destination,
  redline,
}: ResolveTargetOptions): Promise<Result<ResolvedTarget, FolioCliError>> => {
  const source = await identityOfInput(sourcePath);
  if (source.isErr()) return Result.err(source.error);
  if (destination.type === "inPlace") {
    if (redline !== undefined) {
      return Result.err(
        cliError({
          code: FOLIO_CLI_ERROR_CODES.usage,
          message: "A redline is written to a new file, never in place.",
          hint: "Pass -o <redline.docx>.",
        }),
      );
    }
    const named = checkDestinationName(source.value.path);
    if (named.isErr()) return Result.err(named.error);
    return Result.ok({
      sourcePath: source.value.path,
      sourceIdentity: source.value.identity,
      destinationPath: source.value.path,
    });
  }
  const target = await resolveWritePath(destination.path);
  if (target.isErr()) return Result.err(target.error);
  const inputs = [source.value];
  if (redline !== undefined && typeof redline.revisedPath === "string") {
    const revised = await identityOfInput(redline.revisedPath);
    if (revised.isErr()) return Result.err(revised.error);
    inputs.push(revised.value);
  }
  const clash = inputs.find(
    ({ path: inputPath, identity }) =>
      inputPath === target.value.path ||
      (target.value.identity !== null && sameFile(identity, target.value.identity)),
  );
  if (clash !== undefined) {
    return Result.err(
      invalidDestination(
        `The destination is the input ${clash.path}.`,
        redline !== undefined
          ? "Write the redline to a new file."
          : "Use --in-place to change the file itself; it keeps a backup.",
      ),
    );
  }
  return Result.ok({
    sourcePath: source.value.path,
    sourceIdentity: source.value.identity,
    destinationPath: target.value.path,
  });
};

/** Run one write tool as a transaction and return its receipt. */
export const executeWriteTool = async (
  tool: FolioFileToolSpec,
  call: FileToolCall,
  options: WriteOptions,
): Promise<Result<Readonly<Record<string, unknown>>, FolioCliError>> => {
  if (options.txId !== undefined && !TX_ID_PATTERN.test(options.txId)) {
    return Result.err(invalidInput("txId must be 1 to 128 letters, digits, '.', '_' or '-'."));
  }
  if (options.sourcePrecondition === "required" && call.fileVersion === undefined) {
    return Result.err(
      invalidInput(
        "A change needs the fileVersion of the file it was read from.",
        "Pass --expect-version <fileVersion> from your read (MCP: fileVersion), or --no-expect-version to skip the check.",
      ),
    );
  }
  const txId = options.txId ?? randomUUID();
  const target = await resolveTarget({
    sourcePath: call.path,
    destination: options.destination,
    ...(tool.type === "compare" && { redline: { revisedPath: call.args["revisedPath"] } }),
  });
  if (target.isErr()) return Result.err(target.error);
  const { sourcePath, sourceIdentity, destinationPath } = target.value;
  const journalPath = journalPathFor(destinationPath, options.journalPath);
  const inPlace = destinationPath === sourcePath;

  const acquired = await acquireLeaseForWrite({
    documentPath: destinationPath,
    txId,
    force: options.force,
    ...(options.flushWaitMs !== undefined && { flushWaitMs: options.flushWaitMs }),
  });
  if (acquired.isErr()) return Result.err(acquired.error);
  const { lease, flush } = acquired.value;
  try {
    const recovered = await recoverStages({
      documentPath: destinationPath,
      journalPath,
      now: options.date,
    });
    if (recovered.isErr()) return Result.err(recovered.error);

    const requestHash = createHash("sha256")
      .update(
        canonicalJson({
          tool: tool.name,
          args: call.args,
          source: sourcePath,
          destination: destinationPath,
          mode: options.mode,
          repack: options.repack,
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
                message: `Transaction ${txId} already committed a different request.`,
                hint: "Use a new txId for a new change.",
              }),
            );
      }
    }

    // Read after recovery: the version check is against the file as it is now.
    const source = await readDocumentFile(sourcePath);
    if (source.isErr()) return Result.err(source.error);
    // A roll-forward, or the save of an editor that held the lease, renames
    // a new file into place, so an in-place source legitimately has a new
    // identity then. Its version is still checked below.
    const replacedInPlace =
      inPlace &&
      (flush.type !== "none" || recovered.value.some(({ action }) => action === "rolledForward"));
    const expectedIdentity = replacedInPlace ? source.value.identity : sourceIdentity;
    if (!sameFile(source.value.identity, expectedIdentity)) {
      return Result.err(unsafePath(`${sourcePath} was replaced while the write started.`));
    }
    if (inPlace && source.value.links > 1) {
      return Result.err(
        unsafePath(
          `${sourcePath} has other hard links; an in-place write would detach them.`,
          "Write to a new file with -o instead.",
        ),
      );
    }
    const version = checkExpectedVersion(source.value, call.fileVersion);
    const rebased = version.isErr()
      ? await rebasedOntoFlush({
          tool,
          flush,
          inPlace,
          documentPath: destinationPath,
          expected: call.fileVersion,
          current: source.value.fileVersion,
        })
      : null;
    if (version.isErr() && rebased === null) {
      return Result.err(
        recovered.value.length === 0
          ? version.error
          : cliError({
              code: version.error.code,
              message: version.error.message,
              hint: version.error.hint,
              details: {
                ...(isRecord(version.error.details) ? version.error.details : {}),
                recovered: recovered.value,
              },
            }),
      );
    }

    const mutation = await mutate({ tool, source: source.value, args: call.args, options });
    if (mutation.isErr()) return Result.err(mutation.error);
    const { bytes, saveStrategy, repackReason, result, receipts, inputs } = mutation.value;
    const changedParts = await diffPackages(source.value.bytes, bytes);
    if (changedParts.isErr()) return Result.err(changedParts.error);

    const receipt = {
      txId,
      status: "committed",
      tool: tool.name,
      path: destinationPath,
      fromVersion: source.value.fileVersion,
      ...(destinationPath !== sourcePath && {
        source: { path: sourcePath, fileVersion: source.value.fileVersion },
      }),
      ...inputs,
      author: options.author,
      time: options.date,
      saveStrategy,
      ...(repackReason !== undefined && { repackReason }),
      changedParts: changedParts.value,
      ...(recovered.value.length > 0 && { recovered: recovered.value }),
      ...(rebased !== null && { rebased }),
      result,
    };
    const committed = await commitTransaction({
      txId,
      destinationPath,
      destination: options.destination,
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
        tool: tool.name,
        fromVersion: source.value.fileVersion,
        author: options.author,
        time: options.date,
        ops: call.args,
        receipts,
      },
      receipt,
    });
    return committed.isErr() ? Result.err(committed.error) : Result.ok(committed.value.receipt);
  } finally {
    await lease.release();
  }
};
