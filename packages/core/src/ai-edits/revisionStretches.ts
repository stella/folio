/**
 * One id per stretch of an inline revision.
 *
 * An edit inside a pending insertion (or a replacement whose words differ in
 * several places) leaves one revision in stretches with other content between
 * them. A saved package writes each stretch as a wrapper of its own, and
 * `w:id` is unique per wrapper, so the save would give every stretch after the
 * first a fresh id: once reopened they are separate changes, each accepted or
 * rejected alone. The batch gives them those ids when it makes them, so the
 * reviewer lists, resolves, saves and reopens the same changes.
 *
 * A stretch is what the reader lists as one change: the revision's content in
 * one paragraph, uninterrupted but for comment references.
 */

import type { Mark } from "prosemirror-model";
import type { Transaction } from "prosemirror-state";

import { expectTrackedChangeMarkAttrs } from "../prosemirror/attrs";

type Stretch = { end: number; id: number };

type Retarget = { from: number; to: number; mark: Mark; ids: ReadonlyMap<number, number> };

type SeparateRevisionStretchesOptions = {
  tr: Transaction;
  revisionSeed: number;
};

export type SeparatedRevisionStretches = {
  nextRevisionId: number;
  /** Each fresh id, with the revision whose stretch it now names. */
  minted: { revisionId: number; from: number }[];
};

const INLINE_REVISION_MARKS = new Set(["insertion", "deletion"]);

export const separateRevisionStretches = ({
  tr,
  revisionSeed,
}: SeparateRevisionStretchesOptions): SeparatedRevisionStretches => {
  let nextRevisionId = revisionSeed;
  const minted: SeparatedRevisionStretches["minted"] = [];
  const retargets: Retarget[] = [];
  let stretches = new Map<string, Stretch>();
  const transparentEnds = new Map<number, number>();
  const transparentGap = (start: number, end: number): boolean => {
    let position = start;
    while (position < end) {
      const next = transparentEnds.get(position);
      if (next === undefined) return false;
      position = next;
    }
    return position === end;
  };

  tr.doc.descendants((node, pos) => {
    if (node.isTextblock) {
      stretches = new Map();
      return true;
    }
    if (node.type.name === "commentReference") {
      transparentEnds.set(pos, pos + node.nodeSize);
    }
    if (!node.isInline) return true;
    const end = pos + node.nodeSize;
    /** The id this node's content of `kind` revision `revisionId` belongs to. */
    const stretchId = (kind: string, revisionId: number): number => {
      const key = `${kind}:${String(revisionId)}`;
      const stretch = stretches.get(key);
      if (!stretch) {
        stretches.set(key, { end, id: revisionId });
      } else if (stretch.end === end) {
        // Already claimed by this node, through another of its marks.
      } else if (transparentGap(stretch.end, pos)) {
        stretch.end = end;
      } else {
        const fresh = nextRevisionId++;
        minted.push({ revisionId: fresh, from: revisionId });
        stretches.set(key, { end, id: fresh });
      }
      return stretches.get(key)?.id ?? revisionId;
    };
    for (const mark of node.marks) {
      if (!INLINE_REVISION_MARKS.has(mark.type.name)) continue;
      const attrs = expectTrackedChangeMarkAttrs(mark);
      if (attrs.moveKind !== undefined) continue;
      const ids = new Map<number, number>();
      const layers = [
        ...(attrs._docxRevisionAncestors ?? []).flatMap((ancestor) =>
          ancestor.type === "insertion" || ancestor.type === "deletion"
            ? [{ kind: ancestor.type, revisionId: ancestor.revisionId }]
            : [],
        ),
        { kind: mark.type.name, revisionId: attrs.revisionId },
      ];
      for (const { kind, revisionId } of layers) {
        const id = stretchId(kind, revisionId);
        if (id !== revisionId) ids.set(revisionId, id);
      }
      if (ids.size > 0) {
        retargets.push({ from: pos, to: end, mark, ids });
      }
    }
    return true;
  });

  for (const { from, to, mark, ids } of retargets) {
    const attrs = expectTrackedChangeMarkAttrs(mark);
    const ancestors = attrs._docxRevisionAncestors;
    tr.removeMark(from, to, mark);
    tr.addMark(
      from,
      to,
      mark.type.create({
        ...mark.attrs,
        revisionId: ids.get(attrs.revisionId) ?? attrs.revisionId,
        ...(ancestors && {
          _docxRevisionAncestors: ancestors.map((ancestor) =>
            Object.assign(structuredClone(ancestor), {
              revisionId: ids.get(ancestor.revisionId) ?? ancestor.revisionId,
            }),
          ),
        }),
      }),
    );
  }
  return { nextRevisionId, minted };
};
