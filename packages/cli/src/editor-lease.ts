/**
 * The editor lease and the flush handshake: how a live editor (the VS Code
 * custom editor, `folio serve --edit`) and a folio write share one `.docx`
 * so that a write is never refused just because someone is typing, and the
 * user's unsaved edits are saved before the write lands on top of them.
 *
 * Files beside `<dir>/<name>.docx`, all created complete (temporary file,
 * then link or rename) and never followed through a symlink:
 *
 * - `.<name>.docx.folio-lock`: the lease (`lock.ts`). While it has unsaved
 *   edits, an editor holds it long-lived: its own `owner` (`folio-vscode`),
 *   `acceptsFlush: true`, {@link EDITOR_LEASE_MS} long, renewed every
 *   {@link EDITOR_RENEW_MS}. The pid is the editor's long-lived process, so
 *   a crashed editor's lease is stale at once on the same host and within
 *   {@link EDITOR_LEASE_MS} elsewhere.
 * - `.<name>.docx.folio-flush-<id>`: one flush request per waiting writer,
 *   JSON {@link FlushRequest} naming the lease token it asks to let go.
 *
 * A writer (a tool call, or `folio save` without the lease) calls
 * {@link acquireLeaseForWrite}. When a live `acceptsFlush` holder refuses
 * it, it writes a request, then retries the lease every
 * {@link FLUSH_POLL_MS} for up to `flushWaitMs` ({@link DEFAULT_FLUSH_WAIT_MS}).
 * Once the holder has saved and released, the writer takes the lease,
 * removes its request, and re-reads the file: its operations run on the
 * version the editor just saved (block ids and `blockTextHash`
 * preconditions decide whether they still apply). If the holder has not
 * released by the deadline, the writer removes its request and falls back
 * to the ordinary lease rules: an expired lease or one whose process exited
 * is replaced, a live one refuses with `locked` (unless `--force`). Nothing
 * is ever written under a live holder that has not released.
 *
 * The editor:
 * 1. On its first unsaved edit, {@link acquireEditorLease}; keep it with
 *    {@link keepEditorLeaseAlive}. It is refused while a writer's request is
 *    pending (the writer is between the editor's release and its own
 *    acquisition); retry after the writer's commit.
 * 2. {@link watchFlushRequests} with the lease's token. On a request, save
 *    the unsaved edits under the lease (`folio save --lease-token <token>`,
 *    or `saveDocumentBytes` with `leaseToken`), then release it.
 * 3. After a save with no edits pending, release the lease.
 * 4. When the file's version changes on disk (the writer's commit, named by
 *    the journal's newest line), reload it before accepting the next save:
 *    a save names the version it was based on and refuses a stale one.
 */

import { Result } from "better-result";
import { randomUUID } from "node:crypto";
import { watch, type FSWatcher } from "node:fs";
import { readdir, readFile, rename, rm } from "node:fs/promises";
import { hostname } from "node:os";
import path from "node:path";
import { setTimeout as sleep } from "node:timers/promises";

import { cliError, FOLIO_CLI_ERROR_CODES, type FolioCliError } from "./errors";
import { fileVersionOf } from "./file-system";
import {
  acquireLease,
  DEFAULT_LEASE_OWNER,
  readLease,
  type AcquiredLease,
  type LockHolder,
} from "./lock";
import { readSidecarFile, writeNewSidecarFile } from "./sidecar";

/** How long an editor's lease lasts without renewal. */
export const EDITOR_LEASE_MS = 30 * 1000;

/** How often an editor renews its lease while it holds it. */
export const EDITOR_RENEW_MS = 10 * 1000;

/** How long a writer waits for an editor to save and release by default. */
export const DEFAULT_FLUSH_WAIT_MS = 5 * 1000;

/** How often a writer retries the lease, and an editor re-checks for requests. */
export const FLUSH_POLL_MS = 100;

/** A writer's request that the editor holding `leaseToken` save and release. */
export type FlushRequest = {
  id: string;
  /** The token of the lease the writer found; an editor answers only its own. */
  leaseToken: string;
  /** The writer's lease owner, pid, host, and transaction. */
  owner: string;
  pid: number;
  host: string;
  txId: string;
  requestedAt: string;
  /** After this the writer has stopped waiting; the request is void. */
  deadline: string;
};

const flushPrefixFor = (documentPath: string): string =>
  `.${path.basename(documentPath)}.folio-flush-`;

export const flushRequestPathFor = (documentPath: string, id: string): string =>
  path.join(path.dirname(documentPath), `${flushPrefixFor(documentPath)}${id}`);

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

const parseRequest = (text: string): FlushRequest | null => {
  const parsed = Result.try((): unknown => JSON.parse(text));
  if (parsed.isErr() || !isRecord(parsed.value)) return null;
  const { id, leaseToken, owner, pid, host, txId, requestedAt, deadline } = parsed.value;
  return typeof id === "string" &&
    typeof leaseToken === "string" &&
    typeof owner === "string" &&
    typeof pid === "number" &&
    typeof host === "string" &&
    typeof txId === "string" &&
    typeof requestedAt === "string" &&
    typeof deadline === "string"
    ? { id, leaseToken, owner, pid, host, txId, requestedAt, deadline }
    : null;
};

/**
 * The flush requests beside `documentPath` whose writer is still waiting.
 * Requests past their deadline (a writer that crashed) are removed.
 */
export const pendingFlushRequests = async (
  documentPath: string,
  now: Date = new Date(),
): Promise<FlushRequest[]> => {
  const directory = path.dirname(documentPath);
  const prefix = flushPrefixFor(documentPath);
  const listed = await Result.tryPromise(() => readdir(directory));
  if (listed.isErr()) return [];
  const pending: FlushRequest[] = [];
  for (const name of listed.value.filter((entry) => entry.startsWith(prefix))) {
    const requestPath = path.join(directory, name);
    const bytes = await readSidecarFile(requestPath);
    if (bytes.isErr()) continue;
    const request = parseRequest(new TextDecoder().decode(bytes.value));
    if (request === null) continue;
    if (Date.parse(request.deadline) > now.getTime()) {
      pending.push(request);
    } else {
      await rm(requestPath, { force: true });
    }
  }
  return pending;
};

const placeRequest = async (
  documentPath: string,
  request: FlushRequest,
): Promise<Result<void, FolioCliError>> => {
  const target = flushRequestPathFor(documentPath, request.id);
  // Outside the request prefix, so a watcher never reads a partial request.
  const temporary = path.join(
    path.dirname(documentPath),
    `.${path.basename(documentPath)}.folio-flush.${request.id}.tmp`,
  );
  const written = await writeNewSidecarFile(
    temporary,
    new TextEncoder().encode(`${JSON.stringify(request)}\n`),
  );
  if (written.isErr()) return Result.err(written.error);
  const renamed = await Result.tryPromise(() => rename(temporary, target));
  if (renamed.isOk()) return Result.ok();
  await rm(temporary, { force: true });
  return Result.err(
    cliError({
      code: FOLIO_CLI_ERROR_CODES.internal,
      message: `Cannot write ${target}.`,
    }),
  );
};

/** The file's version now, or `null` when it cannot be read. */
const currentVersion = async (documentPath: string): Promise<string | null> => {
  const bytes = await Result.tryPromise(() => readFile(documentPath));
  return bytes.isOk() ? fileVersionOf(bytes.value) : null;
};

/** Whether a writer waited for an editor, and what came of it. */
export type FlushOutcome =
  | { type: "none" }
  /** The holder released within the wait; `versionBefore` is the file's version when asked. */
  | { type: "flushed"; holder: LockHolder; versionBefore: string | null }
  /** The holder never released; the lease was taken by the ordinary rules (stale, or `force`). */
  | { type: "timedOut"; holder: LockHolder };

export type LeaseForWrite = { lease: AcquiredLease; flush: FlushOutcome };

type AcquireLeaseForWriteOptions = {
  documentPath: string;
  txId: string;
  /** After the wait, take the lease over from a holder that did not release. */
  force: boolean;
  owner?: string;
  /** How long to wait for a flush (default {@link DEFAULT_FLUSH_WAIT_MS}); 0 does not ask. */
  flushWaitMs?: number;
  pollMs?: number;
};

const flushableHolder = async (documentPath: string): Promise<LockHolder | null> => {
  const state = await readLease(documentPath);
  return state.type === "held" && state.holder.acceptsFlush === true ? state.holder : null;
};

/**
 * Take the lease for a write. An editor holding it with unsaved edits is
 * asked to save and release first; see the module comment for the rules.
 */
export const acquireLeaseForWrite = async ({
  documentPath,
  txId,
  force,
  owner = DEFAULT_LEASE_OWNER,
  flushWaitMs = DEFAULT_FLUSH_WAIT_MS,
  pollMs = FLUSH_POLL_MS,
}: AcquireLeaseForWriteOptions): Promise<Result<LeaseForWrite, FolioCliError>> => {
  const take = (takeOver: boolean) =>
    acquireLease({ documentPath, txId, owner, force: takeOver, now: new Date() });
  const first = await take(false);
  if (first.isOk()) return Result.ok({ lease: first.value, flush: { type: "none" } });
  if (first.error.code !== FOLIO_CLI_ERROR_CODES.locked) return Result.err(first.error);
  let holder = await flushableHolder(documentPath);
  if (holder === null || flushWaitMs <= 0) {
    if (!force) return Result.err(first.error);
    const forced = await take(true);
    return forced.isOk()
      ? Result.ok({ lease: forced.value, flush: { type: "none" } })
      : Result.err(forced.error);
  }

  const versionBefore = await currentVersion(documentPath);
  const deadline = new Date(Date.now() + flushWaitMs);
  const ask = (target: LockHolder): FlushRequest => ({
    id: randomUUID(),
    leaseToken: target.token,
    owner,
    pid: process.pid,
    host: hostname(),
    txId,
    requestedAt: new Date().toISOString(),
    deadline: deadline.toISOString(),
  });
  let request = ask(holder);
  const placed = await placeRequest(documentPath, request);
  if (placed.isErr()) return Result.err(placed.error);
  try {
    while (Date.now() < deadline.getTime()) {
      await sleep(Math.min(pollMs, Math.max(deadline.getTime() - Date.now(), 0)));
      const attempt = await take(false);
      if (attempt.isOk()) {
        return Result.ok({
          lease: attempt.value,
          flush: { type: "flushed", holder, versionBefore },
        });
      }
      if (attempt.error.code !== FOLIO_CLI_ERROR_CODES.locked) return Result.err(attempt.error);
      const current = await flushableHolder(documentPath);
      if (current !== null && current.token !== request.leaseToken) {
        // Another editor session took the lease meanwhile; ask it instead.
        await rm(flushRequestPathFor(documentPath, request.id), { force: true });
        holder = current;
        request = ask(current);
        const replaced = await placeRequest(documentPath, request);
        if (replaced.isErr()) return Result.err(replaced.error);
      }
    }
    const last = await take(force);
    if (last.isOk()) return Result.ok({ lease: last.value, flush: { type: "timedOut", holder } });
    if (last.error.code !== FOLIO_CLI_ERROR_CODES.locked) return Result.err(last.error);
    return Result.err(
      cliError({
        code: FOLIO_CLI_ERROR_CODES.locked,
        message: `${documentPath} is held by ${holder.owner} (pid ${holder.pid} on ${holder.host}), which did not save and release it within ${String(flushWaitMs)} ms.`,
        hint: "Save or close the document in the editor and retry, or pass --force to take the lease over.",
        details: { holder, flush: "timedOut", flushWaitMs },
      }),
    );
  } finally {
    await rm(flushRequestPathFor(documentPath, request.id), { force: true });
  }
};

type AcquireEditorLeaseOptions = {
  documentPath: string;
  /** The editor's lease owner, such as `folio-vscode`. */
  owner: string;
  /** Recorded in the lock; generated when absent. */
  txId?: string;
  leaseMs?: number;
  now?: Date;
};

/**
 * Take the lease for an editor session with unsaved edits: long-lived,
 * renewable, and answering flush requests. Refused (`locked`) while another
 * holder has it or a writer's flush request is pending.
 */
export const acquireEditorLease = async ({
  documentPath,
  owner,
  txId = `editor-${randomUUID()}`,
  leaseMs = EDITOR_LEASE_MS,
  now = new Date(),
}: AcquireEditorLeaseOptions): Promise<Result<AcquiredLease, FolioCliError>> => {
  const pending = await pendingFlushRequests(documentPath, now);
  if (pending.length > 0) {
    return Result.err(
      cliError({
        code: FOLIO_CLI_ERROR_CODES.locked,
        message: `A write to ${documentPath} is waiting to take the lease.`,
        hint: "Retry after it commits, then reload the document.",
        details: { pending },
      }),
    );
  }
  return await acquireLease({
    documentPath,
    txId,
    owner,
    leaseMs,
    acceptsFlush: true,
    force: false,
    now,
  });
};

export type LeaseKeeper = { stop: () => void };

/**
 * Renew `lease` every `intervalMs` until stopped. When a renewal is refused
 * (the lease was taken over or expired), renewing stops and `onLost` runs:
 * the editor must not save under it again.
 */
export const keepEditorLeaseAlive = (
  lease: AcquiredLease,
  {
    intervalMs = EDITOR_RENEW_MS,
    onLost,
  }: { intervalMs?: number; onLost: (error: FolioCliError) => void },
): LeaseKeeper => {
  let stopped = false;
  const renew = async (): Promise<void> => {
    const renewed = await lease.renew();
    if (renewed.isErr() && !stopped) {
      stopped = true;
      clearInterval(timer);
      onLost(renewed.error);
    }
  };
  const timer = setInterval(() => void renew(), intervalMs);
  timer.unref();
  return {
    stop: () => {
      stopped = true;
      clearInterval(timer);
    },
  };
};

export type FlushWatcher = { close: () => void };

type WatchFlushRequestsOptions = {
  documentPath: string;
  /** The editor's lease token; requests for any other lease are ignored. */
  token: string;
  /** Called once per request: save the unsaved edits, then release the lease. */
  onRequest: (request: FlushRequest) => void;
  pollMs?: number;
};

/**
 * Watch for flush requests addressed to the lease `token`: a directory
 * watch for promptness plus a poll every `pollMs`, since watch events can
 * be dropped or unsupported (network drives).
 */
export const watchFlushRequests = ({
  documentPath,
  token,
  onRequest,
  pollMs = FLUSH_POLL_MS,
}: WatchFlushRequestsOptions): FlushWatcher => {
  const seen = new Set<string>();
  const prefix = flushPrefixFor(documentPath);
  let closed = false;
  let checking = false;
  const check = async (): Promise<void> => {
    if (closed || checking) return;
    checking = true;
    try {
      for (const request of await pendingFlushRequests(documentPath)) {
        if (closed || request.leaseToken !== token || seen.has(request.id)) continue;
        seen.add(request.id);
        onRequest(request);
      }
    } finally {
      checking = false;
    }
  };
  const watcher = Result.try(
    (): FSWatcher =>
      watch(path.dirname(documentPath), (_event, name) => {
        if (name === null || name.startsWith(prefix)) void check();
      }),
  );
  if (watcher.isOk()) watcher.value.on("error", () => undefined);
  const timer = setInterval(() => void check(), pollMs);
  timer.unref();
  void check();
  return {
    close: () => {
      closed = true;
      clearInterval(timer);
      if (watcher.isOk()) watcher.value.close();
    },
  };
};
