/**
 * AutoSaveManager
 *
 * Framework-agnostic class for auto-saving documents to localStorage with
 * crash recovery. Ported from the upstream docx-editor `managers/AutoSaveManager`.
 *
 * Usage with React:
 * ```ts
 * const snapshot = useSyncExternalStore(manager.subscribe, manager.getSnapshot);
 * ```
 *
 * The snapshot/option types are co-located here (and re-exported) because our
 * `managers/types` does not define them.
 */

import type { Document } from "../types/document";
import { Subscribable } from "./Subscribable";

// ============================================================================
// TYPES
// ============================================================================

/** Auto-save status. */
export type AutoSaveStatus = "idle" | "saving" | "saved" | "error";

/** Configuration for {@link AutoSaveManager}. */
export type AutoSaveManagerOptions = {
  /** Storage key for localStorage (default: 'docx-editor-autosave'). */
  storageKey?: string;
  /** Save interval in milliseconds (default: 30000 - 30 seconds). */
  interval?: number;
  /** Maximum age of auto-save before it is considered stale (default: 24 hours). */
  maxAge?: number;
  /** Whether to save on document change with debounce (default: true). */
  saveOnChange?: boolean;
  /** Debounce delay for saveOnChange in milliseconds (default: 2000). */
  debounceDelay?: number;
  /** Callback when save succeeds. */
  onSave?: (timestamp: Date) => void;
  /** Callback when save fails. */
  onError?: (error: Error) => void;
  /** Callback when recovery data is found. */
  onRecoveryAvailable?: (savedDocument: SavedDocumentData) => void;
};

/** Saved document data structure. */
export type SavedDocumentData = {
  /** The document. */
  document: Document;
  /** When the document was saved (ISO string). */
  savedAt: string;
  /** Version for format compatibility. */
  version: number;
  /** Optional document identifier. */
  documentId?: string;
};

/** AutoSaveManager snapshot for UI consumption. */
export type AutoSaveSnapshot = {
  status: AutoSaveStatus;
  lastSaveTime: Date | null;
  hasRecoveryData: boolean;
  isEnabled: boolean;
};

// ============================================================================
// CONSTANTS
// ============================================================================

const DEFAULT_STORAGE_KEY = "docx-editor-autosave";
const DEFAULT_INTERVAL = 30_000; // 30 seconds
const DEFAULT_MAX_AGE = 24 * 60 * 60 * 1000; // 24 hours
const DEFAULT_DEBOUNCE_DELAY = 2000; // 2 seconds
const SAVE_VERSION = 1;

// ============================================================================
// HELPERS
// ============================================================================

function isLocalStorageAvailable(): boolean {
  try {
    const testKey = "__docx_editor_test__";
    localStorage.setItem(testKey, "test");
    localStorage.removeItem(testKey);
    return true;
  } catch {
    return false;
  }
}

function serializeForStorage(document: Document): string {
  return JSON.stringify({ ...document, originalBuffer: null });
}

function isDocumentLike(value: unknown): value is Document {
  return typeof value === "object" && value !== null && "package" in value;
}

function parseSavedData(json: string): SavedDocumentData | null {
  let parsed: unknown;
  try {
    parsed = JSON.parse(json);
  } catch {
    return null;
  }

  if (typeof parsed !== "object" || parsed === null) {
    return null;
  }
  if (!("document" in parsed) || !("savedAt" in parsed) || !("version" in parsed)) {
    return null;
  }

  const { document, savedAt, version } = parsed;
  if (typeof savedAt !== "string" || typeof version !== "number") {
    return null;
  }
  if (!isDocumentLike(document)) {
    return null;
  }

  return { document, savedAt, version };
}

function isStale(savedAt: string, maxAge: number): boolean {
  const savedTime = new Date(savedAt).getTime();
  return Date.now() - savedTime > maxAge;
}

// ============================================================================
// MANAGER
// ============================================================================

export class AutoSaveManager extends Subscribable<AutoSaveSnapshot> {
  private readonly storageKey: string;
  private readonly interval: number;
  private readonly maxAge: number;
  private readonly saveOnChange: boolean;
  private readonly debounceDelay: number;
  private readonly onSaveCallback: ((timestamp: Date) => void) | undefined;
  private readonly onErrorCallback: ((error: Error) => void) | undefined;
  private readonly onRecoveryAvailableCallback: ((saved: SavedDocumentData) => void) | undefined;

  private readonly storageAvailable: boolean;
  private currentDocument: Document | null = null;
  private lastSavedJson: string | null = null;
  private intervalTimer: ReturnType<typeof setInterval> | null = null;
  private debounceTimer: ReturnType<typeof setTimeout> | null = null;

  private status: AutoSaveStatus = "idle";
  private lastSaveTime: Date | null = null;
  private hasRecoveryData = false;
  private isEnabled: boolean;

  constructor(options: AutoSaveManagerOptions = {}) {
    super({
      status: "idle",
      lastSaveTime: null,
      hasRecoveryData: false,
      isEnabled: true,
    });

    this.storageKey = options.storageKey ?? DEFAULT_STORAGE_KEY;
    this.interval = options.interval ?? DEFAULT_INTERVAL;
    this.maxAge = options.maxAge ?? DEFAULT_MAX_AGE;
    this.saveOnChange = options.saveOnChange ?? true;
    this.debounceDelay = options.debounceDelay ?? DEFAULT_DEBOUNCE_DELAY;
    this.onSaveCallback = options.onSave;
    this.onErrorCallback = options.onError;
    this.onRecoveryAvailableCallback = options.onRecoveryAvailable;
    this.isEnabled = true;
    this.storageAvailable = isLocalStorageAvailable();

    this.checkRecoveryData();
  }

  // --------------------------------------------------------------------------
  // PUBLIC API
  // --------------------------------------------------------------------------

  /** Update the current document. Triggers a debounced save if enabled. */
  onDocumentChanged(document: Document | null): void {
    this.currentDocument = document;

    if (this.isEnabled && this.saveOnChange && document && this.storageAvailable) {
      this.debounceSave();
    }
  }

  /** Manually trigger a save. */
  save(): boolean {
    if (!this.storageAvailable) {
      this.onErrorCallback?.(new Error("localStorage is not available"));
      return false;
    }

    const doc = this.currentDocument;
    if (!doc) {
      return false;
    }

    this.updateStatus("saving");

    try {
      const serialized = serializeForStorage(doc);

      // Skip if unchanged
      if (serialized === this.lastSavedJson) {
        this.updateStatus("saved");
        return true;
      }

      this.persistToStorage(doc);
      this.lastSavedJson = serialized;

      const saveTime = new Date();
      this.lastSaveTime = saveTime;
      this.updateStatus("saved");
      this.onSaveCallback?.(saveTime);
      return true;
    } catch (error) {
      this.updateStatus("error");
      this.onErrorCallback?.(error instanceof Error ? error : new Error(String(error)));
      return false;
    }
  }

  /** Clear auto-saved data from storage. */
  clear(): void {
    if (!this.storageAvailable) {
      return;
    }
    try {
      localStorage.removeItem(this.storageKey);
      this.hasRecoveryData = false;
      this.lastSavedJson = null;
      this.emitSnapshot();
    } catch (error) {
      this.onErrorCallback?.(error instanceof Error ? error : new Error(String(error)));
    }
  }

  /** Get recovery data from storage. */
  getRecoveryData(): SavedDocumentData | null {
    if (!this.storageAvailable) {
      return null;
    }
    const savedJson = localStorage.getItem(this.storageKey);
    if (!savedJson) {
      return null;
    }

    const savedData = parseSavedData(savedJson);
    if (!savedData) {
      return null;
    }

    if (isStale(savedData.savedAt, this.maxAge)) {
      this.clear();
      return null;
    }
    return savedData;
  }

  /** Accept recovery and return the document. */
  acceptRecovery(): Document | null {
    const data = this.getRecoveryData();
    if (!data) {
      return null;
    }
    this.hasRecoveryData = false;
    this.emitSnapshot();
    return data.document;
  }

  /** Dismiss recovery and clear saved data. */
  dismissRecovery(): void {
    this.clear();
    this.hasRecoveryData = false;
    this.emitSnapshot();
  }

  /** Enable auto-save and start the interval timer. */
  enable(): void {
    this.isEnabled = true;
    this.startInterval();
    this.emitSnapshot();
  }

  /** Disable auto-save and stop all timers. */
  disable(): void {
    this.isEnabled = false;
    this.stopTimers();
    this.emitSnapshot();
  }

  /** Start the interval timer. Call after enabling or on init. */
  startInterval(): void {
    this.stopTimers();
    if (!this.isEnabled || !this.storageAvailable) {
      return;
    }

    this.intervalTimer = setInterval(() => {
      this.save();
    }, this.interval);
  }

  /** Save synchronously on destroy (best-effort). */
  destroy(): void {
    this.stopTimers();

    if (this.isEnabled && this.currentDocument && this.storageAvailable) {
      try {
        this.persistToStorage(this.currentDocument);
      } catch (error) {
        this.onErrorCallback?.(error instanceof Error ? error : new Error(String(error)));
      }
    }
  }

  // --------------------------------------------------------------------------
  // PRIVATE
  // --------------------------------------------------------------------------

  private checkRecoveryData(): void {
    if (!this.storageAvailable) {
      return;
    }
    const data = this.getRecoveryData();
    if (data) {
      this.hasRecoveryData = true;
      this.emitSnapshot();
      this.onRecoveryAvailableCallback?.(data);
    }
  }

  private persistToStorage(document: Document): void {
    const dataToSave = {
      document: { ...document, originalBuffer: null },
      savedAt: new Date().toISOString(),
      version: SAVE_VERSION,
    };
    localStorage.setItem(this.storageKey, JSON.stringify(dataToSave));
  }

  private debounceSave(): void {
    if (this.debounceTimer) {
      clearTimeout(this.debounceTimer);
    }
    this.debounceTimer = setTimeout(() => {
      this.save();
    }, this.debounceDelay);
  }

  private stopTimers(): void {
    if (this.intervalTimer) {
      clearInterval(this.intervalTimer);
      this.intervalTimer = null;
    }
    if (this.debounceTimer) {
      clearTimeout(this.debounceTimer);
      this.debounceTimer = null;
    }
  }

  private updateStatus(status: AutoSaveStatus): void {
    this.status = status;
    this.emitSnapshot();
  }

  private emitSnapshot(): void {
    this.setSnapshot({
      status: this.status,
      lastSaveTime: this.lastSaveTime,
      hasRecoveryData: this.hasRecoveryData,
      isEnabled: this.isEnabled,
    });
  }
}

// ============================================================================
// UTILITY FUNCTIONS
// ============================================================================

/** Format last save time for display. */
export function formatLastSaveTime(date: Date | null): string {
  if (!date) {
    return "Never";
  }

  const now = new Date();
  const diffMs = now.getTime() - date.getTime();
  const diffSec = Math.floor(diffMs / 1000);
  const diffMin = Math.floor(diffSec / 60);
  const diffHour = Math.floor(diffMin / 60);

  if (diffSec < 10) {
    return "Just now";
  }
  if (diffSec < 60) {
    return `${diffSec} seconds ago`;
  }
  if (diffMin < 60) {
    return `${diffMin} minute${diffMin === 1 ? "" : "s"} ago`;
  }
  if (diffHour < 24) {
    return `${diffHour} hour${diffHour === 1 ? "" : "s"} ago`;
  }

  return date.toLocaleDateString();
}

/** Get auto-save status label. */
export function getAutoSaveStatusLabel(status: AutoSaveStatus): string {
  const labels: Record<AutoSaveStatus, string> = {
    idle: "Ready",
    saving: "Saving...",
    saved: "Saved",
    error: "Save failed",
  };
  return labels[status];
}

/** Get storage size used by auto-save (bytes). */
export function getAutoSaveStorageSize(storageKey: string = DEFAULT_STORAGE_KEY): number {
  try {
    const data = localStorage.getItem(storageKey);
    if (!data) {
      return 0;
    }
    return new Blob([data]).size;
  } catch {
    return 0;
  }
}

/** Format storage size for display. */
export function formatStorageSize(bytes: number): string {
  if (bytes < 1024) {
    return `${bytes} B`;
  }
  if (bytes < 1024 * 1024) {
    return `${(bytes / 1024).toFixed(1)} KB`;
  }
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

/** Check whether auto-save is supported. */
export function isAutoSaveSupported(): boolean {
  return isLocalStorageAvailable();
}
