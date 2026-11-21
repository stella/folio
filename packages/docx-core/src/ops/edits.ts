/**
 * Operations applied one after another, told as one edit: the document the
 * last produced, the inverses in reverse order, and each block's net change.
 */

import type { Document } from "../model/document";
import { idKey, type KeyMembership, packageParagraphIdIndex } from "./ids";
import type {
  DocumentOp,
  ParagraphAbsorption,
  RevisionAbsorption,
  TouchedBlocks,
} from "./types";

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

export type DocumentEditEffects = Pick<DocumentEdit, "inverse" | "touched">;

/** Effects of edits made in turn, with their final document. */
export const combineEditEffects = (
  final: Document,
  edits: readonly DocumentEditEffects[],
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
  const lists = edits.map(({ touched: { absorbed: list } }) => list ?? []);
  const absorbed = lists.some((list) => list.length > 0)
    ? combinedAbsorptionsIn(
        lists,
        new Set(ids("removed")),
        packageParagraphIdIndex(final.package),
      )
    : [];
  const absorbedRevisions = combinedRevisionAbsorptions(edits);
  return {
    document: final,
    inverse: edits.toReversed().flatMap(({ inverse }) => inverse),
    touched: {
      modified: ids("modified"),
      inserted: ids("inserted"),
      removed: ids("removed"),
      ...(absorbed.length === 0 ? {} : { absorbed }),
      ...(absorbedRevisions.length === 0 ? {} : { absorbedRevisions }),
    },
  };
};

/**
 * Revision absorptions made in turn, as one: a change joined into one that a
 * later edit joined on is reported against the change that holds both now.
 */
const combinedRevisionAbsorptions = (
  edits: readonly DocumentEditEffects[],
): RevisionAbsorption[] => {
  const survivorOf = new Map<number, number>();
  for (const { touched } of edits)
    for (const { absorbed, survivor } of touched.absorbedRevisions ?? []) {
      for (const [earlier, into] of survivorOf)
        if (into === absorbed) survivorOf.set(earlier, survivor);
      survivorOf.set(absorbed, survivor);
    }
  return [...survivorOf].map(([absorbed, survivor]) => ({
    absorbed,
    survivor,
  }));
};

/** Edits made in turn, starting from `document`, as one edit. */
export const combineEdits = (
  document: Document,
  edits: readonly DocumentEdit[],
): DocumentEdit =>
  combineEditEffects(edits.at(-1)?.document ?? document, edits);

/**
 * Absorptions made in turn, as one: a paragraph joined into one that a later
 * edit joined on is reported against the paragraph that holds both now. Only
 * paragraphs the edits removed in the end are reported, and only into a
 * survivor the final document holds. Ids compare as addresses do, ignoring
 * case.
 */
export const combinedAbsorptions = (
  lists: readonly (readonly ParagraphAbsorption[])[],
  removed: ReadonlySet<string>,
  present: ReadonlySet<string>,
): ParagraphAbsorption[] =>
  combinedAbsorptionsIn(lists, removed, new Set([...present].map(idKey)));

const combinedAbsorptionsIn = (
  lists: readonly (readonly ParagraphAbsorption[])[],
  removed: ReadonlySet<string>,
  presentKeys: KeyMembership,
): ParagraphAbsorption[] => {
  // Walk backwards: the final destination of a survivor is already known
  // when an earlier paragraph joins it. A later reuse of the absorbed ID
  // replaces its own report, but earlier joins still follow their own path.
  const finalDestination = new Map<string, string>();
  const latest = new Map<string, ParagraphAbsorption>();
  for (const list of lists.toReversed()) {
    for (const { absorbed, survivor } of list.toReversed()) {
      const key = idKey(absorbed);
      const destination = finalDestination.get(idKey(survivor)) ?? survivor;
      finalDestination.set(key, destination);
      if (!latest.has(key))
        latest.set(key, { absorbed, survivor: destination });
    }
  }
  const removedKeys = new Set([...removed].map(idKey));
  const result: ParagraphAbsorption[] = [];
  // Report in first-absorption order, with the spelling of the last record.
  for (const list of lists) {
    for (const { absorbed } of list) {
      const key = idKey(absorbed);
      const entry = latest.get(key);
      if (entry === undefined) continue;
      latest.delete(key);
      if (
        removedKeys.has(key) &&
        !presentKeys.has(key) &&
        presentKeys.has(idKey(entry.survivor))
      )
        result.push(entry);
    }
  }
  return result;
};
