/**
 * Tracked-revision id minting (`w:id` on `<w:ins>` / `<w:del>`).
 *
 * Uniqueness comes from `seedRevisionIdsAbove` (called on load with the
 * document's own max id), NOT from the counter's starting value. A clock seed
 * bought uniqueness at the cost of ids no consumer could read — see
 * `MAX_REVISION_ID`. Port of eigenpal/docx-editor#1093.
 *
 * Kept free of `suggestionMode` imports so the suggestion plugin can mint
 * without a load-order cycle.
 */

import type { Node as PmNode } from "prosemirror-model";

import { MAX_REVISION_ID } from "@stll/docx-core/model";

/**
 * Next id to hand out. Starts at 1 (0 is the serializer's "unusable metadata"
 * fallback) and is raised past a loaded document's existing ids by
 * `seedRevisionIdsAbove`.
 *
 * NOT seeded from a clock. `Date.now()` (~1.8e12) overflows the signed 32-bit
 * int consumers read `w:id` into.
 */
let counter = 1;

/**
 * Mint the next tracked-revision id (`w:id`). Strictly monotonic per realm,
 * and always inside the range serialized `w:id` may occupy.
 */
export function mintRevisionId(): number {
  // Wrap rather than emit an id no consumer can read.
  if (counter > MAX_REVISION_ID) {
    counter = 1;
  }
  return counter++;
}

/**
 * Raise the counter above `maxId` so freshly minted ids cannot collide with
 * ids already present in a loaded document. Only ever raises.
 *
 * `maxId` is derived from file content, so it is untrusted: a malformed or
 * out-of-range value is ignored outright rather than folded into range.
 */
export function seedRevisionIdsAbove(maxId: number): void {
  if (!Number.isInteger(maxId) || maxId < 0 || maxId >= MAX_REVISION_ID) {
    return;
  }
  if (maxId >= counter) {
    counter = maxId + 1;
  }
}

/** Node attrs that carry a revision triple directly. */
const REVISION_ATTR_KEYS = ["pPrIns", "pPrDel", "trIns", "trDel"] as const;

/**
 * Raise the counter above every `revisionId` already present in `doc`.
 * Called from the suggestion-mode plugin's `state.init`.
 */
export function seedRevisionIdsFromDoc(doc: PmNode): void {
  let max = 0;

  const consider = (id: unknown): void => {
    if (typeof id === "number" && Number.isInteger(id) && id > max && id <= MAX_REVISION_ID) {
      max = id;
    }
  };

  doc.descendants((node) => {
    for (const mark of node.marks) {
      consider(mark.attrs["revisionId"]);
    }

    const attrs = node.attrs;
    for (const key of REVISION_ATTR_KEYS) {
      const value = attrs[key] as { revisionId?: unknown } | null | undefined;
      if (value) {
        consider(value.revisionId);
      }
    }
    const cellMarker = attrs["cellMarker"] as
      | { info?: { revisionId?: unknown } }
      | null
      | undefined;
    if (cellMarker?.info) {
      consider(cellMarker.info.revisionId);
    }
  });

  seedRevisionIdsAbove(max);
}
