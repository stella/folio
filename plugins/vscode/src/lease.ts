/**
 * The editor lease, held from the extension host while a document has
 * unsaved edits (`@stll/folio-cli/editor-lease`): a folio write that finds
 * it asks the editor to save and release instead of failing, and the lease's
 * pid is this long-lived process, so a crashed editor's lease goes stale.
 *
 * One {@link EditorLease} per open document: {@link EditorLease.ensure} on
 * the first unsaved edit takes the lease, renews it, and watches for flush
 * requests; {@link EditorLease.release} after a save, a revert, or a close.
 */

import {
  acquireEditorLease,
  keepEditorLeaseAlive,
  watchFlushRequests,
  type FlushRequest,
} from "../../../packages/cli/src/editor-lease";
import type { AcquiredLease } from "../../../packages/cli/src/lock";
import { EDITOR_LEASE_OWNER } from "./save";

/** The parts of the lease module the controller uses, replaceable in tests. */
export type LeaseBackend = {
  readonly acquire: (
    documentPath: string,
  ) => Promise<
    | { readonly ok: true; readonly lease: HeldLease }
    | { readonly ok: false; readonly message: string }
  >;
  readonly keepAlive: (
    lease: HeldLease,
    onLost: (message: string) => void,
  ) => { readonly stop: () => void };
  readonly watch: (
    documentPath: string,
    token: string,
    onRequest: (request: FlushRequest) => void,
  ) => { readonly close: () => void };
};

/** What the controller needs of an acquired lease. */
export type HeldLease = {
  readonly token: string;
  readonly release: () => Promise<void>;
};

type AcquiredHandle = HeldLease & { readonly acquired: AcquiredLease };

const isAcquiredHandle = (lease: HeldLease): lease is AcquiredHandle => "acquired" in lease;

/** The lease module itself. */
export const folioLeaseBackend: LeaseBackend = {
  acquire: async (documentPath) => {
    const acquired = await acquireEditorLease({ documentPath, owner: EDITOR_LEASE_OWNER });
    if (acquired.isErr()) return { ok: false, message: acquired.error.message };
    const lease = acquired.value;
    const handle: AcquiredHandle = {
      token: lease.holder.token,
      acquired: lease,
      release: () => lease.release(),
    };
    return { ok: true, lease: handle };
  },
  keepAlive: (lease, onLost) => {
    if (!isAcquiredHandle(lease)) throw new Error("Not a lease this backend acquired.");
    return keepEditorLeaseAlive(lease.acquired, { onLost: (error) => onLost(error.message) });
  },
  watch: (documentPath, token, onRequest) => watchFlushRequests({ documentPath, token, onRequest }),
};

export type EditorLeaseEvents = {
  /** A write asks the editor to save its edits and release the lease. */
  readonly onFlushRequest: (request: FlushRequest) => void;
  /** Renewal was refused: the lease expired or another process took it over. */
  readonly onLost: (message: string) => void;
};

type Held = {
  readonly lease: HeldLease;
  readonly keeper: { readonly stop: () => void };
  readonly watcher: { readonly close: () => void };
};

export class EditorLease {
  private held: Held | undefined;
  private acquiring: Promise<void> | undefined;
  private disposed = false;
  /** Why the last attempt was refused, for a save that has to go without the lease. */
  lastRefusal: string | undefined;
  private readonly documentPath: string;
  private readonly events: EditorLeaseEvents;
  private readonly backend: LeaseBackend;

  constructor(documentPath: string, events: EditorLeaseEvents, backend = folioLeaseBackend) {
    this.documentPath = documentPath;
    this.events = events;
    this.backend = backend;
  }

  /** The token of the lease this editor holds now. */
  get token(): string | undefined {
    return this.held?.lease.token;
  }

  /**
   * Take the lease unless it is held already. A refusal (another holder, or
   * a write waiting to take it) leaves it free; the next edit tries again.
   */
  ensure(): Promise<void> {
    if (this.held !== undefined || this.disposed) return Promise.resolve();
    this.acquiring ??= this.acquire().finally(() => {
      this.acquiring = undefined;
    });
    return this.acquiring;
  }

  private async acquire(): Promise<void> {
    let acquired: Awaited<ReturnType<LeaseBackend["acquire"]>>;
    try {
      acquired = await this.backend.acquire(this.documentPath);
    } catch (error) {
      acquired = { ok: false, message: error instanceof Error ? error.message : String(error) };
    }
    if (!acquired.ok) {
      this.lastRefusal = acquired.message;
      return;
    }
    this.lastRefusal = undefined;
    const { lease } = acquired;
    if (this.disposed) {
      await lease.release();
      return;
    }
    const keeper = this.backend.keepAlive(lease, (message) => {
      if (this.held?.lease !== lease) return;
      this.stop();
      this.events.onLost(message);
    });
    const watcher = this.backend.watch(this.documentPath, lease.token, (request) =>
      this.events.onFlushRequest(request),
    );
    this.held = { lease, keeper, watcher };
  }

  /** Stop renewing and watching; the lock itself is left as it is. */
  private stop(): Held | undefined {
    const { held } = this;
    this.held = undefined;
    held?.keeper.stop();
    held?.watcher.close();
    return held;
  }

  /** Forget a lease another process took over, without touching its lock. */
  forget(): void {
    this.stop();
  }

  /** Let the lease go: a waiting write can take it. */
  async release(): Promise<void> {
    await this.acquiring;
    await this.stop()?.lease.release();
  }

  async dispose(): Promise<void> {
    this.disposed = true;
    await this.release();
  }
}
