/**
 * Cooperative write lease: `.<name>.folio-lock` beside the document, holding
 * the owner's pid, host, and expiry. A folio write takes it for the whole
 * transaction; a long-lived holder (an editor session) renews it. A lease is
 * stale once it expires or, on the same host, once its process has exited.
 * Reads never consult it.
 */

import { Result } from "better-result";
import { open, readFile, rm, writeFile } from "node:fs/promises";
import { hostname } from "node:os";
import path from "node:path";

import { errnoCode } from "./document";
import { cliError, FOLIO_CLI_ERROR_CODES, type FolioCliError } from "./errors";

/** How long a command-line transaction's lease lasts without renewal. */
export const TRANSACTION_LEASE_MS = 5 * 60 * 1000;

export type LockHolder = {
  owner: string;
  pid: number;
  host: string;
  txId: string;
  acquiredAt: string;
  expiresAt: string;
};

export const lockPathFor = (documentPath: string): string =>
  path.join(path.dirname(documentPath), `.${path.basename(documentPath)}.folio-lock`);

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

const parseHolder = (text: string): LockHolder | null => {
  const parsed = Result.try((): unknown => JSON.parse(text));
  if (parsed.isErr() || !isRecord(parsed.value)) return null;
  const { owner, pid, host, txId, acquiredAt, expiresAt } = parsed.value;
  return typeof owner === "string" &&
    typeof pid === "number" &&
    typeof host === "string" &&
    typeof txId === "string" &&
    typeof acquiredAt === "string" &&
    typeof expiresAt === "string"
    ? { owner, pid, host, txId, acquiredAt, expiresAt }
    : null;
};

const processIsAlive = (pid: number): boolean => {
  const probe = Result.try(() => process.kill(pid, 0));
  return probe.isOk() || errnoCode(probe.error.cause) === "EPERM";
};

/** Whether a lease no longer protects anything. An unreadable lease counts as stale. */
export const isStaleLease = (holder: LockHolder | null, now: Date): boolean => {
  if (holder === null) return true;
  const expires = Date.parse(holder.expiresAt);
  if (Number.isNaN(expires) || expires <= now.getTime()) return true;
  return holder.host === hostname() && !processIsAlive(holder.pid);
};

/** The lease currently beside `documentPath`, if any. */
export const readLease = async (
  documentPath: string,
): Promise<{ type: "free" } | { type: "held"; holder: LockHolder | null }> => {
  const text = await Result.tryPromise(() => readFile(lockPathFor(documentPath), "utf8"));
  if (text.isErr()) return { type: "free" };
  return { type: "held", holder: parseHolder(text.value) };
};

export type AcquiredLease = { release: () => Promise<void>; holder: LockHolder };

type AcquireLeaseOptions = {
  documentPath: string;
  txId: string;
  /** Take the lease over even when another live holder has it. */
  force: boolean;
  now?: Date;
};

const lockedError = (documentPath: string, holder: LockHolder | null): FolioCliError =>
  cliError({
    code: FOLIO_CLI_ERROR_CODES.locked,
    message: `${documentPath} is being written by another process${
      holder === null ? "" : ` (${holder.owner}, pid ${holder.pid} on ${holder.host})`
    }.`,
    hint: "Retry after it finishes, or pass --force to take the lease over.",
    details: holder === null ? undefined : { holder },
  });

const fileSystemError = (message: string, error: unknown): FolioCliError =>
  cliError({
    code: FOLIO_CLI_ERROR_CODES.internal,
    message: `${message}: ${error instanceof Error ? error.message : String(error)}`,
  });

/**
 * Take the lease for one transaction. A stale lease is replaced; a live one
 * refuses with `locked` unless `force`.
 */
export const acquireLease = async ({
  documentPath,
  txId,
  force,
  now = new Date(),
}: AcquireLeaseOptions): Promise<Result<AcquiredLease, FolioCliError>> => {
  const lockPath = lockPathFor(documentPath);
  const holder: LockHolder = {
    owner: "folio-cli",
    pid: process.pid,
    host: hostname(),
    txId,
    acquiredAt: now.toISOString(),
    expiresAt: new Date(now.getTime() + TRANSACTION_LEASE_MS).toISOString(),
  };
  const contents = `${JSON.stringify(holder)}\n`;
  const release = async (): Promise<void> => {
    const current = await readLease(documentPath);
    if (current.type === "held" && current.holder?.txId === txId) {
      await rm(lockPath, { force: true });
    }
  };

  for (let attempt = 0; attempt < 2; attempt++) {
    const created = await Result.tryPromise(async () => {
      const handle = await open(lockPath, "wx");
      try {
        await handle.writeFile(contents);
        await handle.sync();
      } finally {
        await handle.close();
      }
    });
    if (created.isOk()) {
      return Result.ok({ release, holder });
    }
    if (errnoCode(created.error.cause) !== "EEXIST") {
      return Result.err(fileSystemError(`Cannot create ${lockPath}`, created.error.cause));
    }
    const existing = await readLease(documentPath);
    const current = existing.type === "held" ? existing.holder : null;
    if (existing.type === "held" && !isStaleLease(current, now) && !force) {
      return Result.err(lockedError(documentPath, current));
    }
    if (existing.type === "held" && force) {
      const replaced = await Result.tryPromise(() => writeFile(lockPath, contents));
      return replaced.isOk()
        ? Result.ok({ release, holder })
        : Result.err(fileSystemError(`Cannot replace ${lockPath}`, replaced.error.cause));
    }
    await rm(lockPath, { force: true });
  }
  return Result.err(lockedError(documentPath, null));
};
