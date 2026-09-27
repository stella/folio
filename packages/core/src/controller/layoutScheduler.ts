// Framework-agnostic layout scheduler — the first piece of the headless editor
// controller (seam-architecture P4). It owns the "when to lay out" policy:
// coalesce a burst of edits into a single layout pass after a short debounce,
// while enforcing a hard latency cap from the first edit. It handles the editor
// state opaquely (generic `TState`), so it has no ProseMirror or React
// dependency; the adapters wire it to their `runLayoutPipeline`.
//
// A request carries no state. The pass reads the editor's state when it runs,
// so a state replaced in the meantime (a document load, a remount) can never
// be laid out over the current one.

import { LAYOUT_MEASURE, type LayoutRunOptions } from "./layoutRunOptions";

export type RunLayout<TState> = (state: TState, options: LayoutRunOptions) => void;

// Injected timing primitives so the scheduler is environment-agnostic (browser
// rAF today; a host-provided clock for desktop or deterministic tests). Mirrors
// the FolioHost `schedule` seam.
export type SchedulerClock = {
  requestFrame: (callback: () => void) => number;
  cancelFrame: (id: number) => void;
  setTimer: (callback: () => void, delayMs: number) => number;
  clearTimer: (id: number) => void;
  now: () => number;
};

// The browser implementation of the injected clock. window.setTimeout is typed
// to return a number (vs Node's Timeout); window is referenced only inside these
// methods (call-time), so importing the scheduler in a non-browser host is safe
// — that host passes its own SchedulerClock.
export const browserClock: SchedulerClock = {
  requestFrame: (callback) => window.requestAnimationFrame(callback),
  cancelFrame: (id) => window.cancelAnimationFrame(id),
  setTimer: (callback, delayMs) => window.setTimeout(callback, delayMs),
  clearTimer: (id) => window.clearTimeout(id),
  now: () => window.performance.now(),
};

export type LayoutSchedulerConfig<TState> = {
  /**
   * The editor's current state, read when a pass runs; `null` (no editor)
   * skips the pass.
   */
  readState: () => TState | null;
  runLayout: RunLayout<TState>;
  /** Quiet-window debounce before an interactive layout pass. */
  debounceMs: number;
  /** Run the first edit after a max-delay-sized idle period on the next frame. */
  leadingFrame?: boolean;
  /** Hard cap on latency from the first edit in a burst. */
  maxDelayMs: number;
  /**
   * Timing source. Pass `browserClock` in the browser; a host injects its own
   * for desktop/headless. Required (no browser default) so the scheduler never
   * silently depends on `window`.
   */
  clock: SchedulerClock;
};

/** Interactive transaction layout timing shared by framework adapters. */
export const TRANSACTION_LAYOUT_TIMING = {
  debounceMs: 32,
  maxDelayMs: 96,
  leadingFrame: true,
} as const satisfies Pick<
  LayoutSchedulerConfig<unknown>,
  "debounceMs" | "maxDelayMs" | "leadingFrame"
>;

export type LayoutScheduler = {
  /**
   * Request an incremental layout pass after a short coalescing window.
   * Repeated calls in the window join the pending request, so a typing burst
   * paints once while still honoring the max-latency cap.
   */
  schedule: () => void;
  /** Cancel any pending timer/frame. Call on teardown. */
  dispose: () => void;
};

type PendingLayoutRequest = {
  firstScheduledAt: number;
  rafId: number | null;
  timerId: number | null;
};

export const createLayoutScheduler = <TState>(
  config: LayoutSchedulerConfig<TState>,
): LayoutScheduler => {
  const clock = config.clock;
  let lastRunAt: number | null = null;
  let pending: PendingLayoutRequest | null = null;

  const flushPending = (): void => {
    if (!pending || pending.rafId !== null) {
      return;
    }
    pending.timerId = null;
    pending.rafId = clock.requestFrame(() => {
      pending = null;
      lastRunAt = clock.now();
      const state = config.readState();
      if (state === null) {
        return;
      }
      config.runLayout(state, { reason: "transaction", measure: LAYOUT_MEASURE.incremental });
    });
  };

  const armTimer = (request: PendingLayoutRequest): void => {
    if (request.rafId !== null) {
      return;
    }
    if (request.timerId !== null) {
      clock.clearTimer(request.timerId);
    }
    const elapsedMs = clock.now() - request.firstScheduledAt;
    const delayMs =
      elapsedMs >= config.maxDelayMs
        ? 0
        : Math.min(config.debounceMs, config.maxDelayMs - elapsedMs);
    request.timerId = clock.setTimer(flushPending, delayMs);
  };

  return {
    schedule() {
      if (pending) {
        armTimer(pending);
        return;
      }
      const next: PendingLayoutRequest = {
        firstScheduledAt: clock.now(),
        rafId: null,
        timerId: null,
      };
      pending = next;
      if (
        config.leadingFrame &&
        (lastRunAt === null || clock.now() - lastRunAt >= config.maxDelayMs)
      ) {
        flushPending();
        return;
      }
      armTimer(next);
    },
    dispose() {
      if (!pending) {
        return;
      }
      if (pending.timerId !== null) {
        clock.clearTimer(pending.timerId);
      }
      if (pending.rafId !== null) {
        clock.cancelFrame(pending.rafId);
      }
      pending = null;
    },
  };
};
