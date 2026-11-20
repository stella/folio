/**
 * Operations applied one after another, told as one edit: the document the
 * last produced, the inverses in reverse order, and each block's net change.
 */

import type { Document } from "../model/document";
import type { DocumentOp, ParagraphAbsorption, TouchedBlocks } from "./types";

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
} as const satisfies Record<
  TouchedState,
  Record<TouchedState, TouchedState | undefined>
>;

/** Edits made in turn, starting from `document`, as one edit. */
export const combineEdits = (
  document: Document,
  edits: readonly DocumentEdit[],
): DocumentEdit => {
  const touched = new Map<string, TouchedState>();
  for (const edit of edits) {
    for (const state of TOUCHED_STATES) {
      for (const id of edit.touched[state]) {
        const previous = touched.get(id);
        const next =
          previous === undefined ? state : TOUCH_TRANSITIONS[previous][state];
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
  const absorbed = combinedAbsorptions(
    edits.map(({ touched: { absorbed: list } }) => list ?? []),
    new Set(ids("removed")),
  );
  return {
    document: edits.at(-1)?.document ?? document,
    inverse: edits.toReversed().flatMap(({ inverse }) => inverse),
    touched: {
      modified: ids("modified"),
      inserted: ids("inserted"),
      removed: ids("removed"),
      ...(absorbed.length === 0 ? {} : { absorbed }),
    },
  };
};

/**
 * Absorptions made in turn, as one: a paragraph joined into one that a later
 * edit joined on is reported against the paragraph that holds both now. Only
 * paragraphs the edits removed in the end are reported, and only into a
 * survivor that stands.
 */
export const combinedAbsorptions = (
  lists: readonly (readonly ParagraphAbsorption[])[],
  removed: ReadonlySet<string>,
): ParagraphAbsorption[] => {
  const into = new Map<string, string>();
  for (const list of lists) {
    for (const { absorbed, survivor } of list) {
      for (const [earlier, held] of into)
        if (held === absorbed) into.set(earlier, survivor);
      into.set(absorbed, survivor);
    }
  }
  return [...into].flatMap(([absorbed, survivor]) =>
    removed.has(absorbed) && !removed.has(survivor)
      ? [{ absorbed, survivor }]
      : [],
  );
};
