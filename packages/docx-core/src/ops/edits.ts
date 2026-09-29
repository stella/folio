/**
 * Operations applied one after another, told as one edit: the document the
 * last produced, the inverses in reverse order, and each block's net change.
 */

import type { Document } from "../model/document";
import type { DocumentOp, TouchedBlocks } from "./types";

/** What applying operations produced: the document, its inverse, and the blocks it changed. */
export type DocumentEdit = {
  document: Document;
  /** Apply in order to restore the input document. */
  inverse: readonly DocumentOp[];
  touched: TouchedBlocks;
};

type TouchedState = "modified" | "inserted" | "removed";

const TOUCHED_STATES = [
  "modified",
  "inserted",
  "removed",
] as const satisfies readonly TouchedState[];

/** A block's net change after two changes in turn; `undefined` when it came and went. */
const TOUCH_TRANSITIONS = {
  modified: { modified: "modified", inserted: "modified", removed: "removed" },
  inserted: { modified: "inserted", inserted: "inserted", removed: undefined },
  removed: { modified: "modified", inserted: "modified", removed: "removed" },
} as const satisfies Record<TouchedState, Record<TouchedState, TouchedState | undefined>>;

/** Edits made in turn, starting from `document`, as one edit. */
export const combineEdits = (document: Document, edits: readonly DocumentEdit[]): DocumentEdit => {
  const touched = new Map<string, TouchedState>();
  for (const edit of edits) {
    for (const state of TOUCHED_STATES) {
      for (const id of edit.touched[state]) {
        const previous = touched.get(id);
        const next = previous === undefined ? state : TOUCH_TRANSITIONS[previous][state];
        if (next === undefined) {
          touched.delete(id);
        } else {
          touched.set(id, next);
        }
      }
    }
  }
  const ids = (state: TouchedState): string[] =>
    [...touched].flatMap(([id, value]) => (value === state ? [id] : []));
  return {
    document: edits.at(-1)?.document ?? document,
    inverse: edits.toReversed().flatMap(({ inverse }) => inverse),
    touched: { modified: ids("modified"), inserted: ids("inserted"), removed: ids("removed") },
  };
};
