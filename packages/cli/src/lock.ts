/**
 * Cooperative write lease: `.<name>.folio-lock` beside the document, holding
 * the owner's pid, host, expiry, and a random token. A folio write takes it
 * for the whole transaction and checks, just before it journals and renames,
 * that the lock still carries its token; a writer whose lease was taken over
 * stops there. A long-lived holder (an editor session) renews it, and marks
 * it `acceptsFlush` when it answers the flush requests `editor-lease.ts`
 * describes.
 *
 * The lock appears complete or not at all: it is written to a private
 * temporary file and hard-linked into place, which fails when a lock exists.
 * A lease is stale once it expires or, on the same host, once its process has
 * exited. A lock that cannot be parsed is treated as held until it is older
 * than a lease. Reads never consult it.
 */

import { Result } from "better-result";
import { randomUUID } from "node:crypto";
import { link, rename, rm } from "node:fs/promises";
import { hostname } from "node:os";
import path from "node:path";

import { errnoCode } from "./file-system";
import { cliError, FOLIO_CLI_ERROR_CODES, type FolioCliError } from "./errors";
import { inspectPath, readSidecarFile, writeNewSidecarFile } from "./sidecar";

/** How long a command-line transaction's lease lasts without renewal. */
export const TRANSACTION_LEASE_MS = 5 * 60 * 1000;

/** The owner a lease records when the caller names none. */
export const DEFAULT_LEASE_OWNER = "folio-cli";

const OWNER_PATTERN = /^[\w.-]{1,64}$/u;

/** A lease owner is 1 to 64 letters, digits, '.', '_' or '-'. */
export const isLeaseOwner = (value: string): boolean => OWNER_PATTERN.test(value);

export type LockHolder = {
  owner: string;
  pid: number;
  host: string;
  txId: string;
  /** Random per acquisition; the fencing check compares it. */
  token: string;
  acquiredAt: string;
  expiresAt: string;
  /** The holder saves and releases when a write asks it to (`editor-lease.ts`). */
  acceptsFlush?: true;
};

export const lockPathFor = (documentPath: string): string =>
  path.join(path.dirname(documentPath), `.${path.basename(documentPath)}.folio-lock`);

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

const parseHolder = (text: string): LockHolder | null => {
  const parsed = Result.try((): unknown => JSON.parse(text));
  if (parsed.isErr() || !isRecord(parsed.value)) return null;
  const { owner, pid, host, txId, token, acquiredAt, expiresAt, acceptsFlush } = parsed.value;
  return typeof owner === "string" &&
    typeof pid === "number" &&
    typeof host === "string" &&
    typeof txId === "string" &&
    typeof token === "string" &&
    typeof acquiredAt === "string" &&
    typeof expiresAt === "string"
    ? {
        owner,
        pid,
        host,
        txId,
        token,
        acquiredAt,
        expiresAt,
        ...(acceptsFlush === true && { acceptsFlush: true as const }),
      }
    : null;
};

const processIsAlive = (pid: number): boolean => {
  const probe = Result.try(() => process.kill(pid, 0));
  return probe.isOk() || errnoCode(probe.error.cause) === "EPERM";
};

/** What is beside the document: no lock, or a lock with its parsed holder when it has one. */
export type LeaseState =
  | { type: "free" }
  | { type: "held"; holder: LockHolder; modifiedMs: number }
  | { type: "unreadable"; modifiedMs: number };

/**
 * Whether a lock no longer protects anything. One that cannot be read is
 * held until it is older than a whole lease, so a lock another process is
 * writing is never taken for abandoned.
 */
export const isStaleLease = (state: LeaseState, now: Date): boolean => {
  switch (state.type) {
    case "free":
      return true;
    case "unreadable":
      return state.modifiedMs + TRANSACTION_LEASE_MS <= now.getTime();
    case "held": {
      const expires = Date.parse(state.holder.expiresAt);
      if (Number.isNaN(expires) || expires <= now.getTime()) return true;
      return state.holder.host === hostname() && !processIsAlive(state.holder.pid);
    }
    default: {
      const unreachable: never = state;
      return unreachable;
    }
  }
};

/** The lease currently beside `documentPath`. A symlinked lock reads as unreadable. */
export const readLease = async (documentPath: string): Promise<LeaseState> => {
  const lockPath = lockPathFor(documentPath);
  const entry = await inspectPath(lockPath);
  if (entry.isErr()) return { type: "unreadable", modifiedMs: Date.now() };
  if (entry.value.type === "missing") return { type: "free" };
  const modifiedMs = entry.value.type === "file" ? entry.value.modifiedMs : Date.now();
  const bytes = await readSidecarFile(lockPath);
  const holder = bytes.isOk() ? parseHolder(new TextDecoder().decode(bytes.value)) : null;
  return holder === null
    ? { type: "unreadable", modifiedMs }
    : { type: "held", holder, modifiedMs };
};

export type AcquiredLease = {
  /** The lock's contents as last written by this handle. */
  holder: LockHolder;
  /** Refuse unless the lock still carries this lease's token. */
  verify: () => Promise<Result<void, FolioCliError>>;
  /** Push the expiry a lease length past `now`; refuses once the lease was taken over. */
  renew: (now?: Date) => Promise<Result<LockHolder, FolioCliError>>;
  release: () => Promise<void>;
};

type AcquireLeaseOptions = {
  documentPath: string;
  txId: string;
  /** Take the lease over even when another live holder has it. */
  force: boolean;
  /** Who holds it (default `folio-cli`), shown to a writer the lease refuses. */
  owner?: string;
  /** How long the lease lasts without renewal (default {@link TRANSACTION_LEASE_MS}). */
  leaseMs?: number;
  /** Advertise that the holder answers flush requests. */
  acceptsFlush?: boolean;
  now?: Date;
};

const lockedError = (documentPath: string, state: LeaseState): FolioCliError =>
  cliError({
    code: FOLIO_CLI_ERROR_CODES.locked,
    message: `${documentPath} is being written by another process${
      state.type === "held"
        ? ` (${state.holder.owner}, pid ${state.holder.pid} on ${state.holder.host})`
        : ""
    }.`,
    hint: "Retry after it finishes, or pass --force to take the lease over.",
    details: state.type === "held" ? { holder: state.holder } : undefined,
  });

const lostError = (documentPath: string): FolioCliError =>
  cliError({
    code: FOLIO_CLI_ERROR_CODES.locked,
    message: `Another process took over the write lease on ${documentPath}; nothing was written.`,
    hint: "Re-read the document and retry.",
  });

const fileSystemError = (message: string, error: unknown): FolioCliError =>
  cliError({
    code: FOLIO_CLI_ERROR_CODES.internal,
    message: `${message}: ${error instanceof Error ? error.message : String(error)}`,
  });

type PlaceLockResult = { type: "placed" } | { type: "exists" };

/**
 * Put a complete lock in place: write it to a private temporary file, then
 * `link` it to the lock path (fails when one exists) or, taking over,
 * `rename` it there. The lock path never names a partly written file.
 */
const placeLock = async (
  lockPath: string,
  contents: Uint8Array,
  mode: "create" | "replace",
): Promise<Result<PlaceLockResult, FolioCliError>> => {
  const temporary = `${lockPath}.${randomUUID()}.tmp`;
  const written = await writeNewSidecarFile(temporary, contents);
  if (written.isErr()) return Result.err(written.error);
  try {
    const placed = await Result.tryPromise(() =>
      mode === "create" ? link(temporary, lockPath) : rename(temporary, lockPath),
    );
    if (placed.isOk()) return Result.ok({ type: "placed" });
    return errnoCode(placed.error.cause) === "EEXIST"
      ? Result.ok({ type: "exists" })
      : Result.err(fileSystemError(`Cannot create ${lockPath}`, placed.error.cause));
  } finally {
    await rm(temporary, { force: true });
  }
};

const encodeHolder = (holder: LockHolder): Uint8Array =>
  new TextEncoder().encode(`${JSON.stringify(holder)}\n`);

/** The handle on a lease whose lock carries `initial`'s token. */
const leaseFor = (documentPath: string, initial: LockHolder, leaseMs: number): AcquiredLease => {
  const lockPath = lockPathFor(documentPath);
  const holds = async (): Promise<boolean> => {
    const current = await readLease(documentPath);
    return current.type === "held" && current.holder.token === initial.token;
  };
  const lease: AcquiredLease = {
    holder: initial,
    verify: async () => ((await holds()) ? Result.ok() : Result.err(lostError(documentPath))),
    renew: async (now = new Date()) => {
      if (!(await holds())) return Result.err(lostError(documentPath));
      const expiresAt = new Date(now.getTime() + leaseMs).toISOString();
      const renewed = { ...lease.holder, expiresAt };
      const placed = await placeLock(lockPath, encodeHolder(renewed), "replace");
      if (placed.isErr()) return Result.err(placed.error);
      lease.holder = renewed;
      return Result.ok(renewed);
    },
    release: async () => {
      if (await holds()) await rm(lockPath, { force: true });
    },
  };
  return lease;
};

/**
 * Act under a lease another call acquired, named by its token: `folio save
 * --lease-token` from the editor that holds it. Releasing stays with that
 * holder, so this handle's `release` does nothing.
 */
export const adoptLease = async (
  documentPath: string,
  token: string,
): Promise<Result<AcquiredLease, FolioCliError>> => {
  const current = await readLease(documentPath);
  if (current.type !== "held" || current.holder.token !== token) {
    return Result.err(
      cliError({
        code: FOLIO_CLI_ERROR_CODES.locked,
        message: `${documentPath} is not held under that lease token.`,
        hint: "The lease expired or was taken over; acquire it again and retry.",
        details: current.type === "held" ? { holder: current.holder } : undefined,
      }),
    );
  }
  const { holder } = current;
  const leaseMs = Math.max(Date.parse(holder.expiresAt) - Date.parse(holder.acquiredAt), 0);
  const lease = leaseFor(documentPath, holder, leaseMs);
  lease.release = () => Promise.resolve();
  return Result.ok(lease);
};

/**
 * Take the lease for one transaction. A stale lease is replaced; a live or
 * unreadable one refuses with `locked` unless `force`.
 */
export const acquireLease = async ({
  documentPath,
  txId,
  force,
  owner = DEFAULT_LEASE_OWNER,
  leaseMs = TRANSACTION_LEASE_MS,
  acceptsFlush = false,
  now = new Date(),
}: AcquireLeaseOptions): Promise<Result<AcquiredLease, FolioCliError>> => {
  const lockPath = lockPathFor(documentPath);
  const holder: LockHolder = {
    owner,
    pid: process.pid,
    host: hostname(),
    txId,
    token: randomUUID(),
    acquiredAt: now.toISOString(),
    expiresAt: new Date(now.getTime() + leaseMs).toISOString(),
    ...(acceptsFlush && { acceptsFlush: true as const }),
  };
  const contents = encodeHolder(holder);
  const lease = leaseFor(documentPath, holder, leaseMs);

  for (let attempt = 0; attempt < 2; attempt++) {
    const placed = await placeLock(lockPath, contents, "create");
    if (placed.isErr()) return Result.err(placed.error);
    if (placed.value.type === "placed") return Result.ok(lease);
    const existing = await readLease(documentPath);
    if (force) {
      const replaced = await placeLock(lockPath, contents, "replace");
      if (replaced.isErr()) return Result.err(replaced.error);
      return Result.ok(lease);
    }
    if (!isStaleLease(existing, now)) {
      return Result.err(lockedError(documentPath, existing));
    }
    // Two writers can both judge the same lock stale; the one whose lock is
    // replaced fails its fencing check before it journals or renames.
    await rm(lockPath, { force: true });
  }
  return Result.err(lockedError(documentPath, { type: "free" }));
};
