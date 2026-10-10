/**
 * History hook for undo/redo functionality
 *
 * Thin React binding around the framework-agnostic HistoryManager (core). The
 * manager owns the undo/redo stacks, rapid-change grouping, redo invalidation,
 * and the snapshot state; this hook keeps the React glue:
 * - the `useSyncExternalStore` subscription,
 * - the keyboard shortcuts (Ctrl+Z / Ctrl+Y / Ctrl+Shift+Z),
 * - the render-cycle timing that lowers the undo/redo re-entrancy guard,
 * - routing the optional `onUndo` / `onRedo` host callbacks.
 */

import {
  useLayoutEffect,
  useCallback,
  useEffect,
  useRef,
  useState,
  useSyncExternalStore,
} from "react";

import { HistoryManager, classifyHistoryShortcut } from "@stll/folio-core/managers/HistoryManager";
import type { HistoryEntry } from "@stll/folio-core/managers/HistoryManager";

export type { HistoryEntry } from "@stll/folio-core/managers/HistoryManager";

// ============================================================================
// TYPES
// ============================================================================

/**
 * Options for the useHistory hook
 */
export type UseHistoryOptions<T> = {
  /** Maximum number of entries in history (default: 100) */
  maxEntries?: number;
  /** Time in ms to group rapid changes (default: 500) */
  groupingInterval?: number;
  /** Whether to enable keyboard shortcuts (default: true) */
  enableKeyboardShortcuts?: boolean;
  /** Custom comparison function for detecting changes */
  isEqual?: (a: T, b: T) => boolean;
  /** Callback when undo is triggered */
  onUndo?: (state: T) => void;
  /** Callback when redo is triggered */
  onRedo?: (state: T) => void;
  /** Ref to the container element for keyboard events */
  containerRef?: React.RefObject<HTMLElement>;
};

/**
 * Return type of the useHistory hook
 */
export type UseHistoryReturn<T> = {
  /** Current state */
  state: T;
  /** Whether undo is available */
  canUndo: boolean;
  /** Whether redo is available */
  canRedo: boolean;
  /** Number of entries in undo stack */
  undoCount: number;
  /** Number of entries in redo stack */
  redoCount: number;
  /** Push a new state to history */
  push: (newState: T, description?: string) => void;
  /** Undo to previous state */
  undo: () => T | undefined;
  /** Redo to next state */
  redo: () => T | undefined;
  /** Clear all history */
  clear: () => void;
  /** Reset to initial state and clear history */
  reset: (newInitialState?: T) => void;
  /** Get all undo entries (for debugging/display) */
  getUndoStack: () => HistoryEntry<T>[];
  /** Get all redo entries (for debugging/display) */
  getRedoStack: () => HistoryEntry<T>[];
  /** Transform all stored states (current + undo/redo stacks) */
  transformAll: (fn: (state: T) => T) => void;
};

// ============================================================================
// HOOK IMPLEMENTATION
// ============================================================================

/**
 * Custom hook for managing undo/redo history
 */
export function useHistory<T>(
  initialState: T,
  options: UseHistoryOptions<T> = {},
): UseHistoryReturn<T> {
  const {
    maxEntries = 100,
    groupingInterval = 500,
    enableKeyboardShortcuts = true,
    isEqual,
    onUndo,
    onRedo,
    containerRef,
  } = options;

  const [manager] = useState(
    () =>
      new HistoryManager<T>(initialState, {
        maxEntries,
        groupingInterval,
        ...(isEqual ? { isEqual } : {}),
      }),
  );

  const snapshot = useSyncExternalStore(manager.subscribe, manager.getSnapshot);

  // Latest host callbacks + initial state, read by the stable api wrappers so
  // they never need to change identity across renders.
  const onUndoRef = useRef(onUndo);
  useLayoutEffect(() => {
    onUndoRef.current = onUndo;
  });
  const onRedoRef = useRef(onRedo);
  useLayoutEffect(() => {
    onRedoRef.current = onRedo;
  });
  const initialStateRef = useRef(initialState);
  useLayoutEffect(() => {
    initialStateRef.current = initialState;
  });

  // Stable api: the manager is the single source of truth, so these wrappers
  // keep a constant identity. undo/redo layer in the render-cycle timing that
  // lowers the manager's re-entrancy guard plus the host callbacks.
  const [api] = useState(() => {
    const settleUndoRedo = () => {
      // Lower the guard after React has flushed the render + effects triggered
      // by the restored state, so a follow-up edit records a fresh entry again.
      setTimeout(() => manager.endUndoRedo(), 0);
    };

    const undo = (): T | undefined => {
      const result = manager.undo();
      if (result !== undefined) {
        settleUndoRedo();
        onUndoRef.current?.(result);
      }
      return result;
    };

    const redo = (): T | undefined => {
      const result = manager.redo();
      if (result !== undefined) {
        settleUndoRedo();
        onRedoRef.current?.(result);
      }
      return result;
    };

    return {
      push: (newState: T, description?: string) => manager.push(newState, description),
      undo,
      redo,
      clear: () => manager.clear(),
      reset: (newInitialState?: T) => manager.reset(newInitialState ?? initialStateRef.current),
      getUndoStack: () => manager.getUndoStack(),
      getRedoStack: () => manager.getRedoStack(),
      transformAll: (fn: (state: T) => T) => manager.transformAll(fn),
    };
  });

  // Keyboard shortcuts
  useEffect(() => {
    if (!enableKeyboardShortcuts) {
      return;
    }

    const handleKeyDown = (event: KeyboardEvent) => {
      // Ctrl/Cmd+Z (undo), Ctrl/Cmd+Shift+Z and Ctrl/Cmd+Y (redo). The shared
      // classifier matches `key` case-insensitively, so a held Shift reporting
      // the uppercase "Z" still routes to redo.
      const shortcut = classifyHistoryShortcut(event);
      if (!shortcut) {
        return;
      }
      event.preventDefault();
      if (shortcut === "undo") {
        api.undo();
      } else {
        api.redo();
      }
    };

    // Add listener to container or document
    const target = containerRef?.current || document;
    target.addEventListener("keydown", handleKeyDown as EventListener);

    return () => {
      target.removeEventListener("keydown", handleKeyDown as EventListener);
    };
  }, [enableKeyboardShortcuts, containerRef, api]);

  return {
    state: snapshot.state,
    canUndo: snapshot.canUndo,
    canRedo: snapshot.canRedo,
    undoCount: snapshot.undoCount,
    redoCount: snapshot.redoCount,
    ...api,
  };
}

// ============================================================================
// CONVENIENCE HOOKS
// ============================================================================

/**
 * Simplified hook that just tracks state changes automatically
 */
export function useAutoHistory<T>(
  value: T,
  options: UseHistoryOptions<T> = {},
): Omit<UseHistoryReturn<T>, "push"> {
  const history = useHistory(value, options);
  const { push } = history;

  // Automatically push when value changes
  useEffect(() => {
    push(value);
  }, [value, push]);

  return history;
}

/**
 * Hook for document history with specialized comparison
 */
export function useDocumentHistory<
  T extends {
    package?: {
      document?: unknown;
      headers?: unknown;
      footers?: unknown;
    } | null;
  } | null,
>(document: T, options: Omit<UseHistoryOptions<T>, "isEqual"> = {}): UseHistoryReturn<T> {
  // Hot path for editor typing: callers produce fresh references for real
  // document edits, so avoid serializing the whole document to prove equality.
  const isEqual = useCallback((a: T, b: T): boolean => {
    if (a === b) {
      return true;
    }

    const aPackage = a?.package;
    const bPackage = b?.package;
    if (!aPackage || !bPackage) {
      return aPackage === bPackage;
    }

    if (aPackage.document !== bPackage.document) {
      return false;
    }
    if (aPackage.headers !== bPackage.headers) {
      return false;
    }
    if (aPackage.footers !== bPackage.footers) {
      return false;
    }
    return true;
  }, []);

  return useHistory(document, { ...options, isEqual });
}
