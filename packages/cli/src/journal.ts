/**
 * The transaction journal: `.folio/journal.jsonl` beside the document (or
 * `--journal`), one JSON line per committed transaction plus one per crash
 * recovery. A `commit` line is written, and synced, after the staged package
 * validates and before it is renamed into place, so the line is the commit
 * point: a stage left behind by a crash is rolled forward when its line
 * names this document, the stage is the validated package it recorded, and
 * the destination is still what the transaction replaced. Anything else is
 * discarded. The journal, like every sidecar file, is never read or written
 * through a symlink or a hard link.
 */

import { Result } from "better-result";
import { constants } from "node:fs";
import { readdir, rename, rm } from "node:fs/promises";
import path from "node:path";

import { errnoCode, fileVersionOf, sameFile } from "./document";
import { cliError, FOLIO_CLI_ERROR_CODES, type FolioCliError } from "./errors";
import { checkPackageIntegrity } from "./package-parts";
import {
  ensureSidecarParent,
  inspectPath,
  openSidecarFile,
  readSidecarFile,
  SIDECAR_DIRECTORY,
  syncDirectory,
} from "./sidecar";

export type JournalCommit = {
  type: "commit";
  txId: string;
  /** Hash of the request, so a retry with the same txId but other arguments is refused. */
  requestHash: string;
  tool: string;
  /** The destination file. */
  path: string;
  fromVersion: string;
  toVersion: string;
  /** The destination's version before the rename; `null` when it did not exist. */
  destinationVersionBefore: string | null;
  /** The staged file's name, beside the destination. */
  stage: string;
  author: string;
  time: string;
  /** The tool arguments the transaction applied. */
  ops: unknown;
  /** Per-operation receipts from the document-operation contract, when the tool has them. */
  receipts: readonly unknown[];
  /** What the command returned, replayed verbatim for a retry with the same txId. */
  receipt: Readonly<Record<string, unknown>>;
};

export const RECOVERY_DISCARD_REASONS = [
  "notJournaled",
  "otherDocument",
  "notRegularFile",
  "stageChanged",
  "invalidPackage",
  "destinationMoved",
  "leaseLost",
  "renameRefused",
] as const;

export type RecoveryDiscardReason = (typeof RECOVERY_DISCARD_REASONS)[number];

export type JournalRecovery = {
  type: "recovery";
  txId: string;
  path: string;
  time: string;
} & ({ action: "rolledForward" } | { action: "discarded"; reason: RecoveryDiscardReason });

export type JournalEntry = JournalCommit | JournalRecovery;

/** The fields of a commit line recovery and replay read back. */
export type CommittedTransaction = {
  txId: string;
  path: string;
  requestHash: string;
  toVersion: string;
  destinationVersionBefore: string | null;
  receipt: Readonly<Record<string, unknown>>;
};

export const journalPathFor = (documentPath: string, override: string | undefined): string =>
  override === undefined
    ? path.join(path.dirname(documentPath), SIDECAR_DIRECTORY, "journal.jsonl")
    : path.resolve(override);

const journalError = (message: string, error: unknown): FolioCliError =>
  cliError({
    code: FOLIO_CLI_ERROR_CODES.internal,
    message: `${message}: ${error instanceof Error ? error.message : String(error)}`,
  });

/** Append one line and sync it to disk before returning. */
export const appendJournal = async (
  journalPath: string,
  entry: JournalEntry,
): Promise<Result<void, FolioCliError>> => {
  const parent = await ensureSidecarParent(journalPath);
  if (parent.isErr()) return Result.err(parent.error);
  const opened = await openSidecarFile(
    journalPath,
    constants.O_WRONLY | constants.O_APPEND | constants.O_CREAT,
  );
  if (opened.isErr()) return Result.err(opened.error);
  const { handle } = opened.value;
  const written = await Result.tryPromise(async () => {
    try {
      await handle.writeFile(`${JSON.stringify(entry)}\n`);
      await handle.sync();
    } finally {
      await handle.close();
    }
  });
  if (written.isErr()) {
    return Result.err(journalError(`Cannot append to ${journalPath}`, written.error.cause));
  }
  return await syncDirectory(path.dirname(journalPath));
};

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

const parseCommit = (line: string): CommittedTransaction | null => {
  const parsed = Result.try((): unknown => JSON.parse(line));
  if (parsed.isErr() || !isRecord(parsed.value)) return null;
  const {
    type,
    txId,
    path: committedPath,
    requestHash,
    toVersion,
    destinationVersionBefore,
    receipt,
  } = parsed.value;
  return type === "commit" &&
    typeof txId === "string" &&
    typeof committedPath === "string" &&
    typeof requestHash === "string" &&
    typeof toVersion === "string" &&
    (typeof destinationVersionBefore === "string" || destinationVersionBefore === null) &&
    isRecord(receipt)
    ? { txId, path: committedPath, requestHash, toVersion, destinationVersionBefore, receipt }
    : null;
};

const isDiscardRecord = (line: string, txId: string): boolean => {
  const parsed = Result.try((): unknown => JSON.parse(line));
  return (
    parsed.isOk() &&
    isRecord(parsed.value) &&
    parsed.value["type"] === "recovery" &&
    parsed.value["txId"] === txId &&
    parsed.value["action"] === "discarded"
  );
};

/**
 * The committed transaction with `txId`, if the journal has one and recovery
 * did not discard it. Torn or foreign lines (a crash mid-append) are skipped.
 * A journal (or `.folio`) that is a symlink or hard link is refused.
 */
export const findCommit = async (
  journalPath: string,
  txId: string,
): Promise<Result<CommittedTransaction | undefined, FolioCliError>> => {
  const journal = await inspectPath(journalPath);
  if (journal.isErr()) return Result.err(journal.error);
  if (journal.value.type === "missing") {
    const parent = await inspectPath(path.dirname(journalPath));
    if (parent.isErr()) return Result.err(parent.error);
    return parent.value.type === "missing" || parent.value.type === "directory"
      ? Result.ok(undefined)
      : Result.err(
          cliError({
            code: FOLIO_CLI_ERROR_CODES.unsafePath,
            message: `${path.dirname(journalPath)} is not a plain directory.`,
          }),
        );
  }
  const parent = await ensureSidecarParent(journalPath);
  if (parent.isErr()) return Result.err(parent.error);
  const bytes = await readSidecarFile(journalPath);
  if (bytes.isErr()) return Result.err(bytes.error);
  let found: CommittedTransaction | undefined;
  for (const line of new TextDecoder().decode(bytes.value).split("\n")) {
    if (!line.includes(txId)) continue;
    if (isDiscardRecord(line, txId)) return Result.ok(undefined);
    const commit = parseCommit(line);
    if (commit?.txId === txId) found ??= commit;
  }
  return Result.ok(found);
};

export const stagePrefixFor = (documentPath: string): string =>
  `.${path.basename(documentPath)}.folio-stage-`;

export const stagePathFor = (documentPath: string, txId: string): string =>
  path.join(path.dirname(documentPath), `${stagePrefixFor(documentPath)}${txId}`);

export type RecoveryAction =
  | { txId: string; action: "rolledForward" }
  | { txId: string; action: "discarded"; reason: RecoveryDiscardReason };

type RecoverStagesOptions = {
  documentPath: string;
  journalPath: string;
  now: string;
};

type StageVerdict = { type: "rollForward" } | { type: "discard"; reason: RecoveryDiscardReason };

/** The destination's current version, read without following a symlink; `null` when missing. */
const destinationVersion = async (
  documentPath: string,
): Promise<Result<string | null, FolioCliError>> => {
  const entry = await inspectPath(documentPath);
  if (entry.isErr()) return Result.err(entry.error);
  if (entry.value.type === "missing") return Result.ok(null);
  const bytes = await readSidecarFile(documentPath);
  return bytes.isErr() ? Result.err(bytes.error) : Result.ok(fileVersionOf(bytes.value));
};

type JudgeStageOptions = {
  documentPath: string;
  stagePath: string;
  commit: CommittedTransaction;
};

/**
 * Whether a journaled stage may complete its rename: its line names this
 * document, it is a regular single-link file (a symlink is never followed),
 * its bytes are the version the line recorded and still parse as a package,
 * and the destination is still what the transaction replaced.
 */
const judgeStage = async ({
  documentPath,
  stagePath,
  commit,
}: JudgeStageOptions): Promise<Result<StageVerdict, FolioCliError>> => {
  if (commit.path !== documentPath) return Result.ok({ type: "discard", reason: "otherDocument" });
  const entry = await inspectPath(stagePath);
  if (entry.isErr()) return Result.err(entry.error);
  if (entry.value.type !== "file" || entry.value.links !== 1) {
    return Result.ok({ type: "discard", reason: "notRegularFile" });
  }
  const opened = await openSidecarFile(stagePath, constants.O_RDONLY);
  if (opened.isErr()) return Result.ok({ type: "discard", reason: "notRegularFile" });
  const { handle, identity } = opened.value;
  const read = await Result.tryPromise(async () => new Uint8Array(await handle.readFile()));
  await handle.close();
  if (read.isErr() || !sameFile(identity, entry.value.identity)) {
    return Result.ok({ type: "discard", reason: "stageChanged" });
  }
  if (fileVersionOf(read.value) !== commit.toVersion) {
    return Result.ok({ type: "discard", reason: "stageChanged" });
  }
  const valid = await checkPackageIntegrity(read.value, []);
  if (valid.isErr()) return Result.ok({ type: "discard", reason: "invalidPackage" });
  const current = await destinationVersion(documentPath);
  if (current.isErr()) return Result.err(current.error);
  return current.value === commit.destinationVersionBefore
    ? Result.ok({ type: "rollForward" })
    : Result.ok({ type: "discard", reason: "destinationMoved" });
};

/**
 * Settle every stage a crashed transaction left beside `documentPath`. The
 * caller holds the document's lease, so no live transaction owns a stage.
 * Removing a discarded stage unlinks the name only; a symlink's target is
 * never touched.
 */
export const recoverStages = async ({
  documentPath,
  journalPath,
  now,
}: RecoverStagesOptions): Promise<Result<RecoveryAction[], FolioCliError>> => {
  const directory = path.dirname(documentPath);
  const prefix = stagePrefixFor(documentPath);
  const listed = await Result.tryPromise(() => readdir(directory));
  if (listed.isErr()) {
    return errnoCode(listed.error.cause) === "ENOENT"
      ? Result.ok([])
      : Result.err(journalError(`Cannot list ${directory}`, listed.error.cause));
  }
  const actions: RecoveryAction[] = [];
  for (const stage of listed.value.filter((name) => name.startsWith(prefix))) {
    const txId = stage.slice(prefix.length);
    const stagePath = path.join(directory, stage);
    const commit = await findCommit(journalPath, txId);
    if (commit.isErr()) return Result.err(commit.error);
    if (commit.value === undefined) {
      // Never journaled, so never committed: nothing to record.
      await rm(stagePath, { force: true });
      continue;
    }
    const verdict = await judgeStage({ documentPath, stagePath, commit: commit.value });
    if (verdict.isErr()) return Result.err(verdict.error);
    let action: RecoveryAction;
    if (verdict.value.type === "rollForward") {
      const renamed = await Result.tryPromise(() => rename(stagePath, documentPath));
      if (renamed.isErr()) {
        return Result.err(journalError(`Cannot complete ${stage}`, renamed.error.cause));
      }
      const synced = await syncDirectory(directory);
      if (synced.isErr()) return Result.err(synced.error);
      action = { txId, action: "rolledForward" };
    } else {
      await rm(stagePath, { force: true });
      action = { txId, action: "discarded", reason: verdict.value.reason };
    }
    const recorded = await appendJournal(journalPath, {
      type: "recovery",
      path: documentPath,
      time: now,
      ...action,
    });
    if (recorded.isErr()) return Result.err(recorded.error);
    actions.push(action);
  }
  return Result.ok(actions);
};
