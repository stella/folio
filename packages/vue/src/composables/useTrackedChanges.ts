import type { EditorView } from "prosemirror-view";
/**
 * useTrackedChanges — Vue composable wrapping the framework-agnostic
 * extractTrackedChanges from core. Returns a computed result that
 * re-derives whenever the editor view's transaction count changes.
 */
import { computed, type Ref } from "vue";

import {
  lockHeldRevisions,
  trackedChangeEntryRevisionIds,
} from "@stll/folio-core/prosemirror/commands/comments";
import {
  extractTrackedChanges,
  type TrackedChangesResult,
} from "@stll/folio-core/prosemirror/utils/extractTrackedChanges";

import type { TrackedChangeEntry } from "../utils/comments";

export type { TrackedChangesResult };
export { extractTrackedChanges };
export type { TrackedChangeEntry } from "../utils/comments";

/**
 * Pass the editor view ref + a stateTick that bumps on every PM
 * transaction (the same tick the Toolbar listens to). The computed
 * re-runs whenever the tick advances.
 */
export function useTrackedChanges(
  view: Ref<EditorView | null>,
  stateTick: Ref<number>,
) {
  return computed<TrackedChangesResult & { entries: TrackedChangeEntry[] }>(
    () => {
      void stateTick.value;
      const state = view.value?.state ?? null;
      const result = extractTrackedChanges(state);
      if (state === null) return result;
      const held = {
        accept: lockHeldRevisions(state, "accept"),
        reject: lockHeldRevisions(state, "reject"),
      };
      if (held.accept.size === 0 && held.reject.size === 0) return result;
      // A card resolves every revision it stands for, so one held revision
      // holds the card. The entries are this reading's own.
      const entries: TrackedChangeEntry[] = result.entries;
      for (const entry of entries) {
        const ids = trackedChangeEntryRevisionIds(entry);
        entry.heldByLock = {
          accept: ids.some((id) => held.accept.has(id)),
          reject: ids.some((id) => held.reject.has(id)),
        };
      }
      return { ...result, entries };
    },
  );
}
