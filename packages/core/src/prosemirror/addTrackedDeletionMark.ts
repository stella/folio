import { panic } from "better-result";
import type { Mark } from "prosemirror-model";
import type { Transform } from "prosemirror-transform";

import { expectTrackedChangeMarkAttrs } from "./attrs";
import { trackedRevisionLayerOf } from "./trackedRevisionPath";
import { canCarryTrackedRunMark } from "./trackedRunInlineAtoms";

type AddTrackedDeletionMarkOptions<T extends Transform> = {
  tr: T;
  from: number;
  to: number;
  mark: Mark;
  insertionPolicy: "preserve-pending" | "retract-own";
};

/** Already deleted runs keep their revision identity and enclosing insertion path. */
export const addTrackedDeletionMark = <T extends Transform>({
  tr,
  from,
  to,
  mark,
  insertionPolicy,
}: AddTrackedDeletionMarkOptions<T>): T => {
  if (mark.type.name !== "deletion") panic("A tracked deletion requires a deletion mark");
  // AddMarkStep also marks descendants of an inline atom (a structured field).
  // Restore their existing deletion marks after marking the enclosing carrier.
  const preservedDeletions: { from: number; to: number; mark: Mark }[] = [];
  tr.doc.nodesBetween(from, to, (node, position) => {
    const existing = node.marks.find(({ type }) => type.name === "deletion");
    if (node.isInline && existing) {
      const start = Math.max(from, position);
      const end = Math.min(to, position + node.nodeSize);
      if (start < end) preservedDeletions.push({ from: start, to: end, mark: existing });
    }
    return true;
  });
  const mapFrom = tr.mapping.maps.length;
  const ranges: { from: number; to: number; mark: Mark; disposition: "mark" | "retract" }[] = [];
  tr.doc.nodesBetween(from, to, (node, position) => {
    if (!canCarryTrackedRunMark(node)) return true;
    if (node.marks.some(({ type }) => type.name === "deletion")) return false;
    const start = Math.max(from, position);
    const end = Math.min(to, position + node.nodeSize);
    if (start >= end) return false;
    const insertion = node.marks.find(({ type }) => type.name === "insertion");
    const ownInsertion =
      insertion !== undefined && insertion.attrs["author"] === mark.attrs["author"];
    const ancestors = insertion
      ? expectTrackedChangeMarkAttrs(insertion)._docxRevisionAncestors
      : null;
    const deletion =
      insertion && ancestors?.length
        ? mark.type.create({
            ...mark.attrs,
            _docxRevisionAncestors: [...ancestors, trackedRevisionLayerOf(insertion, node)],
          })
        : mark;
    const disposition = insertionPolicy === "retract-own" && ownInsertion ? "retract" : "mark";
    const previous = ranges.at(-1);
    if (
      previous?.to === start &&
      previous.disposition === disposition &&
      previous.mark.eq(deletion)
    )
      previous.to = end;
    else ranges.push({ from: start, to: end, mark: deletion, disposition });
    return false;
  });
  // Retractions change positions: process the original ranges right to left.
  for (const range of ranges.toReversed()) {
    if (range.disposition === "retract") tr.delete(range.from, range.to);
    else tr.addMark(range.from, range.to, range.mark);
  }
  const mapping = tr.mapping.slice(mapFrom);
  for (const preserved of preservedDeletions) {
    const start = mapping.map(preserved.from);
    const end = mapping.map(preserved.to);
    if (start < end) tr.addMark(start, end, preserved.mark);
  }
  return tr;
};
