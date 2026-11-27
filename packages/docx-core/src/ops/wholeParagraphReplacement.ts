import { panic, Result } from "better-result";

import type { Document, TrackedChangeInfo } from "../model/document";
import { sourceCommentAnchorsIn } from "./comments";
import { idKey } from "./ids";
import {
  childNodes,
  type CommentAnchorCarrier,
  isCommentAnchor,
  leafSpans,
  nestedRecords,
  zeroWidthLeavesAt,
} from "./leaves";
import { selectedParagraphRuns } from "./plan";
import type { DocumentOpRefusal } from "./refusal";
import { isTrackedWrapper } from "./review";
import { isOwnRecord, type RevisionOwner } from "./revisionOwner";
import type { TextPosition } from "./types";

type WholeParagraphReplacementOptions = {
  document: Document;
  from: TextPosition;
  to: TextPosition;
  /**
   * Who a suggesting replacement records, absent for a direct edit. A
   * suggestion retracts only its own pending insertions; a host with
   * another owner's stays with the ordinary range planner.
   */
  suggesting?: RevisionOwner | undefined;
};

type WholeParagraphReplacementPlan = {
  blockIds: readonly string[];
  commentIds: ReadonlySet<number>;
};

/**
 * Whole source hosts before a surviving paragraph can retire through range
 * deletion and joins. Their complete comment ownership needs no alignment
 * inside a surviving host. Partial first hosts, section endpoints and pending
 * removals, moves or property changes in a host stay with the ordinary range
 * planner.
 */
export const planWholeParagraphReplacement = ({
  document,
  from,
  to,
  suggesting,
}: WholeParagraphReplacementOptions): Result<
  WholeParagraphReplacementPlan | null,
  DocumentOpRefusal
> => {
  if (
    idKey(from.blockId) === idKey(to.blockId) ||
    from.offset !== 0 ||
    (from.zeroWidthBefore !== undefined && from.zeroWidthBefore !== 0)
  )
    return Result.ok(null);
  const runs = selectedParagraphRuns(document, from, to);
  if (runs.isErr()) return Result.err(runs.error);
  if (runs.value.length !== 1) return Result.ok(null);
  const hosts = runs.value.at(0)?.slice(0, -1) ?? [];
  if (hosts.length === 0) return Result.ok(null);
  const first = hosts.at(0);
  if (first === undefined)
    return panic("A nonempty source prefix has its first host.");
  if (
    from.zeroWidthBefore === undefined &&
    zeroWidthLeavesAt(first.paragraph.content, 0).length > 0
  )
    return Result.ok(null);
  // A pending insertion, words or break, goes with its host: the direct edit
  // removes it, and a suggestion retracts its author's own. Pending removals
  // and moves outlive a direct edit, and an empty wrapper is a leaf of its own.
  const retires = (info: TrackedChangeInfo) =>
    suggesting === undefined || isOwnRecord(info, suggesting);
  const blockIds: string[] = [];
  const starts = new Set<number>();
  const ends = new Set<number>();
  const references = new Set<number>();
  for (const { paragraph } of hosts) {
    if (
      paragraph.paraId === undefined ||
      paragraph.sectionProperties !== undefined ||
      (paragraph.pPrMark !== undefined &&
        (paragraph.pPrMark.kind !== "ins" ||
          !retires(paragraph.pPrMark.info))) ||
      paragraph.markPropertyChange !== undefined ||
      (paragraph.propertyChanges?.length ?? 0) > 0
    )
      return Result.ok(null);
    blockIds.push(paragraph.paraId);
    for (const { node, ancestors } of leafSpans(paragraph.content)) {
      if (
        isTrackedWrapper(node) ||
        ancestors.some(
          (ancestor) =>
            isTrackedWrapper(ancestor) &&
            (ancestor.type !== "insertion" || !retires(ancestor.info)),
        )
      )
        return Result.ok(null);
      // Atomic field results have source owners outside the addressable leaf walk.
      if (childNodes(node) === undefined && nestedRecords(node) !== undefined)
        return Result.ok(null);
      if (!isCommentAnchor(node)) continue;
      switch (node.type) {
        case "commentRangeStart":
          starts.add(node.id);
          break;
        case "commentRangeEnd":
          ends.add(node.id);
          break;
        case "commentReference":
          references.add(node.id);
          break;
        default: {
          const unreachable: never = node;
          return unreachable;
        }
      }
    }
  }
  const candidates = new Set([...starts, ...ends, ...references]);
  if (candidates.size === 0) return Result.ok(null);
  const sourceCounts = {
    commentRangeStart: new Map<number, number>(),
    commentRangeEnd: new Map<number, number>(),
    commentReference: new Map<number, number>(),
  } as const satisfies Record<
    CommentAnchorCarrier["type"],
    Map<number, number>
  >;
  for (const anchor of sourceCommentAnchorsIn(document)) {
    const counts = sourceCounts[anchor.type];
    counts.set(anchor.id, (counts.get(anchor.id) ?? 0) + 1);
  }
  const definitions = new Map(
    (document.package.document.comments ?? []).map((comment) => [
      comment.id,
      comment,
    ]),
  );
  const knownPoint = (id: number): boolean => {
    const seen = new Set<number>();
    let current = definitions.get(id);
    while (current !== undefined) {
      if (seen.has(current.id))
        return panic("A validated comment thread cannot contain a cycle.");
      seen.add(current.id);
      if (current.anchorKind === "point") return true;
      if (current.parentId === undefined) return false;
      current = definitions.get(current.parentId);
    }
    return false;
  };
  const commentIds = new Set<number>();
  for (const id of candidates) {
    if (
      (sourceCounts.commentReference.get(id) ?? 0) !==
      (references.has(id) ? 1 : 0)
    )
      continue;
    if (
      starts.has(id) &&
      ends.has(id) &&
      sourceCounts.commentRangeStart.get(id) === 1 &&
      sourceCounts.commentRangeEnd.get(id) === 1
    )
      commentIds.add(id);
    else if (
      !starts.has(id) &&
      !ends.has(id) &&
      !sourceCounts.commentRangeStart.has(id) &&
      !sourceCounts.commentRangeEnd.has(id) &&
      references.has(id) &&
      knownPoint(id)
    )
      commentIds.add(id);
  }
  return Result.ok(commentIds.size === 0 ? null : { blockIds, commentIds });
};
