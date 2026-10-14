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
 * A formatting revision instead belongs to one serialized run carrier. A
 * later edit can split that carrier without changing its property mark, so
 * each resulting carrier gets its own id through the same allocator.
 */

import type { Mark } from "prosemirror-model";
import type { Transaction } from "prosemirror-state";
import { panic } from "better-result";

import {
  expectRunPropertyChangeMarkAttrs,
  expectTrackedChangeMarkAttrs,
} from "../prosemirror/attrs";
import {
  applyMarksToRunFormattingRepresentation,
  expandRunFormattingCarrier,
  type RunFormattingCarrierRepresentation,
} from "../prosemirror/runFormattingInlineCarriers";

type Stretch = { id: number; interrupted: boolean };

type Retarget = { from: number; to: number; mark: Mark; ids: ReadonlyMap<number, number> };

type FormattingRetarget = {
  representation: RunFormattingCarrierRepresentation;
  mark: Mark;
  ids: ReadonlyMap<number, number>;
};

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
  const formattingRetargets: FormattingRetarget[] = [];
  const formattingIds = new Set<number>();
  const formattingRepresentations = new Set<number>();
  // The open stretches of the current paragraph.
  let stretches = new Map<string, Stretch>();

  tr.doc.descendants((node, pos) => {
    if (node.isTextblock) {
      stretches = new Map();
      return true;
    }
    // A property revision lives on one w:r, rather than the wrapper around
    // several runs. Later text/revision edits can split that run: give every
    // resulting carrier its id now, before the save's uniqueness pass would.
    const carrier = node.marks.some(({ type }) => type.name === "runPropertyChange")
      ? expandRunFormattingCarrier(node, pos)
      : null;
    for (const representation of carrier?.representations ?? []) {
      if (carrier?.disposition === "structured-field" && representation.role === "owner") {
        // The outer field projects its result runs; it emits no w:rPr itself.
        continue;
      }
      const mark = representation.node.marks.find(({ type }) => type.name === "runPropertyChange");
      if (!mark) continue;
      if (formattingRepresentations.has(representation.position)) continue;
      formattingRepresentations.add(representation.position);
      const ids = new Map<number, number>();
      for (const { info } of expectRunPropertyChangeMarkAttrs(mark).changes) {
        if (!formattingIds.has(info.id)) {
          formattingIds.add(info.id);
          continue;
        }
        const fresh = nextRevisionId++;
        minted.push({ revisionId: fresh, from: info.id });
        ids.set(info.id, fresh);
      }
      if (ids.size > 0) formattingRetargets.push({ representation, mark, ids });
    }
    // Content, not the inline containers around it: a revision that runs
    // into a content control and out again is one wrapper when saved.
    if (!node.isInline || !(node.isLeaf || node.isText)) return true;
    const layersOf = (mark: Mark) => {
      const attrs = expectTrackedChangeMarkAttrs(mark);
      return [
        ...(attrs._docxRevisionAncestors ?? []).flatMap((ancestor) =>
          ancestor.type === "insertion" || ancestor.type === "deletion"
            ? [{ kind: ancestor.type, revisionId: ancestor.revisionId }]
            : [],
        ),
        { kind: mark.type.name, revisionId: attrs.revisionId },
      ];
    };
    const revisionMarks = node.marks.filter(
      (mark) =>
        INLINE_REVISION_MARKS.has(mark.type.name) &&
        expectTrackedChangeMarkAttrs(mark).moveKind === undefined,
    );
    const carried = new Set(
      revisionMarks.flatMap((mark) =>
        layersOf(mark).map(({ kind, revisionId }) => `${kind}:${String(revisionId)}`),
      ),
    );
    // Content outside a revision ends the stretch it interrupts; a comment
    // reference sits inside a wrapper and ends nothing.
    if (node.type.name !== "commentReference") {
      for (const [key, stretch] of stretches) {
        if (!carried.has(key)) stretch.interrupted = true;
      }
    }
    for (const mark of revisionMarks) {
      const ids = new Map<number, number>();
      for (const { kind, revisionId } of layersOf(mark)) {
        const key = `${kind}:${String(revisionId)}`;
        let stretch = stretches.get(key);
        if (!stretch) {
          stretch = { id: revisionId, interrupted: false };
          stretches.set(key, stretch);
        } else if (stretch.interrupted) {
          const fresh = nextRevisionId++;
          minted.push({ revisionId: fresh, from: revisionId });
          stretch = { id: fresh, interrupted: false };
          stretches.set(key, stretch);
        }
        if (stretch.id !== revisionId) ids.set(revisionId, stretch.id);
      }
      if (ids.size > 0) {
        retargets.push({ from: pos, to: pos + node.nodeSize, mark, ids });
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
  for (const { representation, mark, ids } of formattingRetargets) {
    const live = tr.doc.nodeAt(representation.position);
    if (!live) panic("A formatting revision carrier vanished during normalization");
    const changes = expectRunPropertyChangeMarkAttrs(mark).changes.map((change) =>
      Object.assign({}, change, {
        info: Object.assign({}, change.info, { id: ids.get(change.info.id) ?? change.info.id }),
      }),
    );
    const replacement = mark.type.create({ ...mark.attrs, changes });
    applyMarksToRunFormattingRepresentation({
      tr,
      representation: {
        node: live,
        position: representation.position,
        from: representation.position,
        to: representation.position + representation.node.nodeSize,
      },
      marks: replacement.addToSet(live.marks),
    });
  }
  return { nextRevisionId, minted };
};
