/**
 * The transaction journal: `.folio/journal.jsonl` beside the document (or
 * `--journal`), one JSON line per committed transaction plus one per crash
 * recovery. A `commit` line is written, and synced, after the staged package
 * validates and before it is renamed into place, so the line is the commit
 * point: a stage left behind by a crash is rolled forward when its line
 * exists and the destination is still what the transaction replaced, and
 * discarded otherwise.
 */

import { Result } from "better-result";
import { mkdir, open, readdir, readFile, rename, rm } from "node:fs/promises";
import path from "node:path";

import { errnoCode, fileVersionOf } from "./document";
import { cliError, FOLIO_CLI_ERROR_CODES, type FolioCliError } from "./errors";

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

export type JournalRecovery = {
  type: "recovery";
  txId: string;
  path: string;
  action: "rolledForward" | "discarded";
  time: string;
};

export type JournalEntry = JournalCommit | JournalRecovery;

/** The fields of a commit line recovery and replay read back. */
export type CommittedTransaction = {
  txId: string;
  requestHash: string;
  toVersion: string;
  destinationVersionBefore: string | null;
  receipt: Readonly<Record<string, unknown>>;
};

export const journalPathFor = (documentPath: string, override: string | undefined): string =>
  override === undefined
    ? path.join(path.dirname(documentPath), ".folio", "journal.jsonl")
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
  const written = await Result.tryPromise(async () => {
    await mkdir(path.dirname(journalPath), { recursive: true });
    const handle = await open(journalPath, "a");
    try {
      await handle.writeFile(`${JSON.stringify(entry)}\n`);
      await handle.sync();
    } finally {
      await handle.close();
    }
  });
  return written.isOk()
    ? Result.ok()
    : Result.err(journalError(`Cannot append to ${journalPath}`, written.error.cause));
};

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

const parseCommit = (line: string): CommittedTransaction | null => {
  const parsed = Result.try((): unknown => JSON.parse(line));
  if (parsed.isErr() || !isRecord(parsed.value)) return null;
  const { type, txId, requestHash, toVersion, destinationVersionBefore, receipt } = parsed.value;
  return type === "commit" &&
    typeof txId === "string" &&
    typeof requestHash === "string" &&
    typeof toVersion === "string" &&
    (typeof destinationVersionBefore === "string" || destinationVersionBefore === null) &&
    isRecord(receipt)
    ? { txId, requestHash, toVersion, destinationVersionBefore, receipt }
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
 */
export const findCommit = async (
  journalPath: string,
  txId: string,
): Promise<CommittedTransaction | undefined> => {
  const text = await Result.tryPromise(() => readFile(journalPath, "utf8"));
  if (text.isErr()) return undefined;
  let found: CommittedTransaction | undefined;
  for (const line of text.value.split("\n")) {
    if (!line.includes(txId)) continue;
    if (isDiscardRecord(line, txId)) return undefined;
    const commit = parseCommit(line);
    if (commit?.txId === txId) found ??= commit;
  }
  return found;
};

export const stagePrefixFor = (documentPath: string): string =>
  `.${path.basename(documentPath)}.folio-stage-`;

export const stagePathFor = (documentPath: string, txId: string): string =>
  path.join(path.dirname(documentPath), `${stagePrefixFor(documentPath)}${txId}`);

const hashIfPresent = async (filePath: string): Promise<string | null> => {
  const bytes = await Result.tryPromise(() => readFile(filePath));
  return bytes.isOk() ? fileVersionOf(bytes.value) : null;
};

export type RecoveryAction = { txId: string; action: JournalRecovery["action"] };

type RecoverStagesOptions = {
  documentPath: string;
  journalPath: string;
  now: string;
};

/**
 * Settle every stage a crashed transaction left beside `documentPath`. The
 * caller holds the document's lease, so no live transaction owns a stage.
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
    if (commit === undefined) {
      // Never journaled, so never committed.
      await rm(stagePath, { force: true });
      continue;
    }
    const stageVersion = await hashIfPresent(stagePath);
    const destinationVersion = await hashIfPresent(documentPath);
    const rollForward =
      stageVersion === commit.toVersion && destinationVersion === commit.destinationVersionBefore;
    if (rollForward) {
      const renamed = await Result.tryPromise(() => rename(stagePath, documentPath));
      if (renamed.isErr()) {
        return Result.err(journalError(`Cannot complete ${stage}`, renamed.error.cause));
      }
    } else {
      await rm(stagePath, { force: true });
    }
    const action = rollForward ? "rolledForward" : "discarded";
    const recorded = await appendJournal(journalPath, {
      type: "recovery",
      txId,
      path: documentPath,
      action,
      time: now,
    });
    if (recorded.isErr()) return Result.err(recorded.error);
    actions.push({ txId, action });
  }
  return Result.ok(actions);
};
