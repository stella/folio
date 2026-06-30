/**
 * HistoryManager
 *
 * Framework-agnostic undo/redo stack and document-snapshot history, extracted
 * from the React `useHistory` hook. Owns the current state plus the undo/redo
 * stacks and runs push (with rapid-change grouping), undo, redo, redo
 * invalidation, stack-size capping, and bulk state transforms.
 *
 * Adapters bind to it through `Subscribable` (React:
 * `useSyncExternalStore(manager.subscribe, manager.getSnapshot)`). The adapter
 * keeps the genuinely framework-bound concerns: keyboard shortcuts, the
 * render-cycle timing that lowers the undo/redo re-entrancy guard, and routing
 * the optional `onUndo` / `onRedo` host callbacks.
 *
 * `HistoryManager<Document | null>` structurally satisfies the
 * `DocumentLoaderHistory` shape that `DocumentLoaderManager` consumes
 * (`{ readonly state; reset(document) }`).
 */

import { Subscribable } from "./Subscribable";

/** History entry containing state and metadata. */
export type HistoryEntry<T> = {
  /** The state at this point. */
  state: T;
  /** Timestamp when this entry was created. */
  timestamp: number;
  /** Optional description of what changed. */
  description?: string;
};

/** Reactive snapshot adapters render from. */
export type HistorySnapshot<T> = {
  /** Current state. */
  state: T;
  /** Whether undo is available. */
  canUndo: boolean;
  /** Whether redo is available. */
  canRedo: boolean;
  /** Number of entries in the undo stack. */
  undoCount: number;
  /** Number of entries in the redo stack. */
  redoCount: number;
};

export type HistoryManagerOptions<T> = {
  /** Maximum number of entries in the undo stack (default: 100). */
  maxEntries?: number;
  /** Time in ms within which rapid pushes are grouped (default: 500). */
  groupingInterval?: number;
  /** Custom comparison used to skip no-op pushes (default: JSON compare). */
  isEqual?: (a: T, b: T) => boolean;
};

/** Default equality check using JSON stringify. */
export const defaultHistoryIsEqual = <T>(a: T, b: T): boolean =>
  JSON.stringify(a) === JSON.stringify(b);

export class HistoryManager<T> extends Subscribable<HistorySnapshot<T>> {
  private currentState: T;
  private readonly initialState: T;
  private undoStack: HistoryEntry<T>[] = [];
  private redoStack: HistoryEntry<T>[] = [];
  private lastPushTime = 0;
  /**
   * Raised while an undo/redo is settling. Pushes that arrive during this
   * window (e.g. the editor re-syncing to the restored document) adopt the new
   * state without creating a fresh history entry. The adapter lowers it once
   * its render cycle has flushed.
   */
  private undoRedoActive = false;
  private readonly maxEntries: number;
  private readonly groupingInterval: number;
  private readonly isEqual: (a: T, b: T) => boolean;

  constructor(initialState: T, options: HistoryManagerOptions<T> = {}) {
    super({ state: initialState, canUndo: false, canRedo: false, undoCount: 0, redoCount: 0 });
    this.currentState = initialState;
    this.initialState = initialState;
    this.maxEntries = options.maxEntries ?? 100;
    this.groupingInterval = options.groupingInterval ?? 500;
    this.isEqual = options.isEqual ?? defaultHistoryIsEqual;
  }

  /** The current state (satisfies `DocumentLoaderHistory.state`). */
  get state(): T {
    return this.currentState;
  }

  /** Whether the undo/redo re-entrancy guard is currently raised. */
  get isUndoRedoActive(): boolean {
    return this.undoRedoActive;
  }

  /** Push a new state to history. */
  push(newState: T, description?: string): void {
    // Skip if state hasn't changed.
    if (this.isEqual(this.currentState, newState)) {
      return;
    }

    // During an undo/redo settle window, adopt the new state without recording
    // a fresh entry or clearing the redo stack.
    if (this.undoRedoActive) {
      this.currentState = newState;
      this.emit();
      return;
    }

    const now = Date.now();
    const timeSinceLastPush = now - this.lastPushTime;

    if (timeSinceLastPush < this.groupingInterval && this.undoStack.length > 0) {
      // Fold this change into the most recent entry, keeping its stored state
      // (the state from before the grouped run) so a later undo restores to the
      // start of the group, not just one push back.
      const last = this.undoStack.at(-1);
      const desc = description || last?.description;
      this.undoStack[this.undoStack.length - 1] = {
        state: last?.state ?? this.currentState,
        timestamp: now,
        ...(desc !== undefined ? { description: desc } : {}),
      };
    } else {
      this.undoStack.push({
        state: this.currentState,
        timestamp: now,
        ...(description !== undefined ? { description } : {}),
      });
      // Limit stack size.
      if (this.undoStack.length > this.maxEntries) {
        this.undoStack = this.undoStack.slice(this.undoStack.length - this.maxEntries);
      }
    }

    // Clear redo stack on a new change.
    this.redoStack = [];
    this.currentState = newState;
    this.lastPushTime = now;
    this.emit();
  }

  /** Undo to the previous state, raising the re-entrancy guard. */
  undo(): T | undefined {
    if (this.undoStack.length === 0) {
      return undefined;
    }

    this.undoRedoActive = true;

    const prevEntry = this.undoStack.at(-1);
    if (!prevEntry) {
      return undefined;
    }
    this.undoStack = this.undoStack.slice(0, -1);
    this.redoStack = [...this.redoStack, { state: this.currentState, timestamp: Date.now() }];
    this.currentState = prevEntry.state;
    this.emit();
    return prevEntry.state;
  }

  /** Redo to the next state, raising the re-entrancy guard. */
  redo(): T | undefined {
    if (this.redoStack.length === 0) {
      return undefined;
    }

    this.undoRedoActive = true;

    const nextEntry = this.redoStack.at(-1);
    if (!nextEntry) {
      return undefined;
    }
    this.redoStack = this.redoStack.slice(0, -1);
    this.undoStack = [...this.undoStack, { state: this.currentState, timestamp: Date.now() }];
    this.currentState = nextEntry.state;
    this.emit();
    return nextEntry.state;
  }

  /**
   * Lower the undo/redo re-entrancy guard. The adapter calls this once its
   * render cycle (and the effects it triggered) has flushed, so a follow-up
   * edit records a fresh entry again.
   */
  endUndoRedo(): void {
    this.undoRedoActive = false;
  }

  /** Clear both stacks, keeping the current state. */
  clear(): void {
    this.undoStack = [];
    this.redoStack = [];
    this.emit();
  }

  /** Reset to the initial (or a supplied) state and clear history. */
  reset(newInitialState?: T): void {
    this.currentState = newInitialState ?? this.initialState;
    this.undoStack = [];
    this.redoStack = [];
    this.lastPushTime = 0;
    this.emit();
  }

  /** Snapshot copy of the undo stack (for debugging/display). */
  getUndoStack(): HistoryEntry<T>[] {
    return [...this.undoStack];
  }

  /** Snapshot copy of the redo stack (for debugging/display). */
  getRedoStack(): HistoryEntry<T>[] {
    return [...this.redoStack];
  }

  /**
   * Transform every stored state (current + both stacks). Useful for bulk
   * cleanup such as stripping cached snapshots.
   */
  transformAll(fn: (state: T) => T): void {
    this.currentState = fn(this.currentState);
    this.undoStack = this.undoStack.map((entry) => ({ ...entry, state: fn(entry.state) }));
    this.redoStack = this.redoStack.map((entry) => ({ ...entry, state: fn(entry.state) }));
    this.emit();
  }

  private emit(): void {
    this.setSnapshot({
      state: this.currentState,
      canUndo: this.undoStack.length > 0,
      canRedo: this.redoStack.length > 0,
      undoCount: this.undoStack.length,
      redoCount: this.redoStack.length,
    });
  }
}
