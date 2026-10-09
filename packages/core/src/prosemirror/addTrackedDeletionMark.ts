import { panic } from "better-result";
import type { Mark, Node as PMNode } from "prosemirror-model";
import type { Transform } from "prosemirror-transform";

import { expectTrackedChangeMarkAttrs } from "./attrs";
import { trackedRevisionLayerOf } from "./trackedRevisionPath";
import { canCarryTrackedRunMark } from "./trackedRunInlineAtoms";

type AddTrackedDeletionMarkOptions = {
  tr: Transform;
  from: number;
  to: number;
  mark: Mark;
  insertionPolicy: "preserve-pending" | "retract-own";
};

/**
 * Already deleted runs keep their revision identity and enclosing insertion path.
 * An own insertion containing deleted descendants is marked rather than retracted:
 * removing its carrier would also remove another pending revision.
 */
export const addTrackedDeletionMark = ({
  tr,
  from,
  to,
  mark,
  insertionPolicy,
}: AddTrackedDeletionMarkOptions): void => {
  if (mark.type.name !== "deletion") panic("A tracked deletion requires a deletion mark");
  // AddMarkStep also marks descendants of an inline atom (a structured field).
  // Restore their existing deletion marks after marking the enclosing carrier.
  const preservedDeletions: { from: number; to: number; mark: Mark }[] = [];
  const protectedCarriers = new Set<PMNode>();
  tr.doc.nodesBetween(from, to, (node, position) => {
    const existing = node.marks.find(({ type }) => type.name === "deletion");
    if (node.isInline && existing) {
      const start = Math.max(from, position);
      const end = Math.min(to, position + node.nodeSize);
      if (start < end) preservedDeletions.push({ from: start, to: end, mark: existing });
      const at = tr.doc.resolve(position);
      for (let depth = 1; depth <= at.depth; depth++) {
        const ancestor = at.node(depth);
        if (canCarryTrackedRunMark(ancestor)) protectedCarriers.add(ancestor);
      }
    }
    return true;
  });
  const mapFrom = tr.mapping.maps.length;
  const ranges: { from: number; to: number; mark: Mark; disposition: "mark" | "retract" }[] = [];
  const descendantDeletions: { from: number; to: number; mark: Mark }[] = [];
  const deletionFor = (node: PMNode): Mark => {
    const insertion = node.marks.find(({ type }) => type.name === "insertion");
    const ancestors = insertion
      ? expectTrackedChangeMarkAttrs(insertion)._docxRevisionAncestors
      : null;
    return insertion && ancestors?.length
      ? mark.type.create({
          ...mark.attrs,
          _docxRevisionAncestors: [...ancestors, trackedRevisionLayerOf(insertion, node)],
        })
      : mark;
  };
  tr.doc.nodesBetween(from, to, (node, position) => {
    if (!canCarryTrackedRunMark(node)) return true;
    if (node.marks.some(({ type }) => type.name === "deletion")) return false;
    const start = Math.max(from, position);
    const end = Math.min(to, position + node.nodeSize);
    if (start >= end) return false;
    const insertion = node.marks.find(({ type }) => type.name === "insertion");
    const ownInsertion =
      insertion !== undefined && insertion.attrs["author"] === mark.attrs["author"];
    const deletion = deletionFor(node);
    const disposition =
      insertionPolicy === "retract-own" && ownInsertion && !protectedCarriers.has(node)
        ? "retract"
        : "mark";
    const previous = ranges.at(-1);
    if (
      previous?.to === start &&
      previous.disposition === disposition &&
      previous.mark.eq(deletion)
    )
      previous.to = end;
    else ranges.push({ from: start, to: end, mark: deletion, disposition });
    if (disposition === "mark") {
      node.descendants((child, offset) => {
        if (child.marks.some(({ type }) => type.name === "deletion")) return false;
        const childDeletion = deletionFor(child);
        if (childDeletion === mark) return true;
        const childStart = Math.max(from, position + 1 + offset);
        const childEnd = Math.min(to, position + 1 + offset + child.nodeSize);
        if (childStart < childEnd)
          descendantDeletions.push({ from: childStart, to: childEnd, mark: childDeletion });
        return true;
      });
    }
    return false;
  });
  // Retractions change positions: process the original ranges right to left.
  for (const range of ranges.toReversed()) {
    if (range.disposition === "retract") tr.delete(range.from, range.to);
    else tr.addMark(range.from, range.to, range.mark);
  }
  const mapping = tr.mapping.slice(mapFrom);
  // AddMarkStep propagates the carrier's mark. Restore each child's own path,
  // then restore prior deletions so existing ownership always takes precedence.
  for (const preserved of [...descendantDeletions, ...preservedDeletions]) {
    const start = mapping.map(preserved.from);
    const end = mapping.map(preserved.to);
    if (start < end) tr.addMark(start, end, preserved.mark);
  }
};
