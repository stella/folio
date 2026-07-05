/**
 * useAutoSave — Vue composable wrapping AutoSaveManager from core.
 *
 * Persists document to localStorage with configurable interval. Tracks the
 * manager's snapshot (status, lastSaveTime, hasRecoveryData, isEnabled) as
 * Vue refs so UI components can react to it. Recovery detection is exposed
 * for crash-recovery UX.
 *
 * The snapshot/option types are co-located in the core manager module and
 * re-exported here. Manager `save()` is synchronous in our fork (returns a
 * boolean), so the composable's `save` returns `boolean` rather than a
 * promise.
 */

import { onBeforeUnmount, ref, unref, watch, type MaybeRef, type Ref } from "vue";
import {
  AutoSaveManager,
  formatLastSaveTime,
  getAutoSaveStatusLabel,
  getAutoSaveStorageSize,
  formatStorageSize,
  isAutoSaveSupported,
  type AutoSaveManagerOptions,
  type AutoSaveStatus,
  type SavedDocumentData,
} from "@stll/folio-core/managers/AutoSaveManager";
import type { Document } from "@stll/folio-core/types/document";

export type { AutoSaveStatus, SavedDocumentData };
export {
  formatLastSaveTime,
  getAutoSaveStatusLabel,
  getAutoSaveStorageSize,
  formatStorageSize,
  isAutoSaveSupported,
};

export interface UseAutoSaveOptions {
  /** localStorage key (default: 'docx-editor-autosave') */
  storageKey?: string;
  /** Auto-save interval in ms (default: 30000) */
  interval?: number;
  /** Whether auto-save starts enabled (default: true). */
  enabled?: boolean;
  /** Maximum age of auto-save before it is considered stale. */
  maxAge?: number;
  /** Callback when save succeeds. */
  onSave?: (timestamp: Date) => void;
  /** Callback when save fails. */
  onError?: (error: Error) => void;
  /** Callback when recovery data is found. */
  onRecoveryAvailable?: (savedDocument: SavedDocumentData) => void;
  /** Whether document changes trigger a debounced save. */
  saveOnChange?: boolean;
  /** Debounce delay for saveOnChange in milliseconds. */
  debounceDelay?: number;
}

export interface UseAutoSaveReturn {
  status: Ref<AutoSaveStatus>;
  lastSaveTime: Ref<Date | null>;
  save: () => boolean;
  clearAutoSave: () => void;
  hasRecoveryData: Ref<boolean>;
  getRecoveryData: () => SavedDocumentData | null;
  acceptRecovery: () => Document | null;
  dismissRecovery: () => void;
  isEnabled: Ref<boolean>;
  enable: () => void;
  disable: () => void;
}

export function useAutoSave(
  document: MaybeRef<Document | null | undefined>,
  options: UseAutoSaveOptions = {},
): UseAutoSaveReturn {
  const status = ref<AutoSaveStatus>("idle");
  const lastSaveTime = ref<Date | null>(null);
  const hasRecoveryData = ref(false);
  const isEnabled = ref(true);
  const {
    storageKey,
    interval,
    enabled: initialEnabled = true,
    maxAge,
    onSave,
    onError,
    onRecoveryAvailable,
    saveOnChange,
    debounceDelay,
  } = options;

  if (!isAutoSaveSupported()) {
    return {
      status,
      lastSaveTime,
      hasRecoveryData,
      isEnabled,
      save: () => false,
      clearAutoSave: () => {},
      getRecoveryData: () => null,
      acceptRecovery: () => null,
      dismissRecovery: () => {},
      enable: () => {},
      disable: () => {},
    };
  }

  // exactOptionalPropertyTypes: only forward keys the caller actually set.
  const managerOptions: AutoSaveManagerOptions = {
    ...(storageKey !== undefined && { storageKey }),
    ...(interval !== undefined && { interval }),
    ...(maxAge !== undefined && { maxAge }),
    ...(saveOnChange !== undefined && { saveOnChange }),
    ...(debounceDelay !== undefined && { debounceDelay }),
    ...(onSave !== undefined && { onSave }),
    ...(onError !== undefined && { onError }),
    ...(onRecoveryAvailable !== undefined && { onRecoveryAvailable }),
  };
  const manager = new AutoSaveManager(managerOptions);

  const sync = () => {
    const snapshot = manager.getSnapshot();
    status.value = snapshot.status;
    lastSaveTime.value = snapshot.lastSaveTime;
    hasRecoveryData.value = snapshot.hasRecoveryData;
    isEnabled.value = snapshot.isEnabled;
  };

  const unsubscribe = manager.subscribe(sync);
  if (initialEnabled) manager.enable();
  else manager.disable();
  sync();

  watch(
    () => unref(document),
    (doc) => manager.onDocumentChanged(doc ?? null),
    { immediate: true },
  );

  onBeforeUnmount(() => {
    unsubscribe();
    manager.destroy();
  });

  return {
    status,
    lastSaveTime,
    hasRecoveryData,
    isEnabled,
    save: () => manager.save(),
    clearAutoSave: () => manager.clear(),
    getRecoveryData: () => manager.getRecoveryData(),
    acceptRecovery: () => manager.acceptRecovery(),
    dismissRecovery: () => manager.dismissRecovery(),
    enable: () => manager.enable(),
    disable: () => manager.disable(),
  };
}
