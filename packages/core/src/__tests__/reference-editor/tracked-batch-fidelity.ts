import { Result } from "better-result";
import { Fragment, type Node as PMNode } from "prosemirror-model";
import { EditorState } from "prosemirror-state";

import {
  commentAnchorIdsForNode,
  ownCommentAnchorIds,
} from "../../ai-edits/comment-lifecycle";
import {
  comparisonApplying,
  noteReferenceOf,
} from "../../ai-edits/note-references";
import { changedTopLevelWindow } from "../../ai-edits/result-validation";
import {
  type Inventory,
  missingProtectedContent,
  newlyDeletedProtectedContent,
  sameOriginalBlocks,
} from "../../ai-edits/tracked-batch-fidelity";
import { visitCommentMarkers } from "../../docx/commentAnchorIndex";
import { proseDocToBlocks } from "../../prosemirror/conversion/fromProseDoc";
import { isBlockBoundaryNode } from "../../prosemirror/keptRangeBoundaries";
import { readRangeAnchorAttrs } from "../../prosemirror/rangeAnchorAttrs";
import {
  getSuggestions,
  projectAllChangesInHeadlessState,
  rejectEverySuggestion,
} from "./comments";
import { pluginsForHeadlessRevisionResolution } from "./headlessRevisionResolutionGuard";

type TrackedBatchFidelityOptions = {
  before: EditorState;
  after: PMNode;
  createdCommentIds: readonly number[];
  retiredCommentIds?: readonly number[];
};

const copyRange = (doc: PMNode, from: number, to: number): PMNode => {
  if (from === 0 && to === doc.childCount) return doc;
  const nodes: PMNode[] = [];
  for (let index = from; index < to; index++) nodes.push(doc.child(index));
  return doc.copy(Fragment.fromArray(nodes));
};

const hasStoryBoundary = (doc: PMNode): boolean => {
  let found = false;
  doc.descendants((node) => {
    found ||=
      isBlockBoundaryNode(node) ||
      node.type.name === "rangeAnchor" ||
      node.type.name === "bookmarkBoundary" ||
      node.type.name === "moveRangeBoundary" ||
      node.attrs["_sectionProperties"] != null ||
      noteReferenceOf(node) !== null ||
      commentAnchorIdsForNode(node).length > 0;
    return !found;
  });
  return found;
};

const isStableBoundary = (node: PMNode): boolean =>
  node.type.name === "paragraph" && node.attrs["pPrMark"] == null;

const hasParagraphMark = (node: PMNode): boolean =>
  node.type.name === "paragraph" && node.attrs["pPrMark"] != null;

/** Include the neighbors a paragraph join or boundary relocation may reach. */
const resolutionWindow = (before: PMNode, after: PMNode) => {
  const window = changedTopLevelWindow(before, after);
  if (window === null) return null;
  let { from, beforeTo, afterTo } = window;
  while (from > 0) {
    from--;
    if (isStableBoundary(before.child(from))) {
      // A pending predecessor can join the otherwise stable left neighbor.
      const predecessor = before.maybeChild(from - 1);
      if (
        predecessor === null ||
        (!hasParagraphMark(predecessor) && !isBlockBoundaryNode(predecessor))
      ) {
        break;
      }
    }
  }
  while (beforeTo < before.childCount && afterTo < after.childCount) {
    const boundary = before.child(beforeTo);
    beforeTo++;
    afterTo++;
    if (isStableBoundary(boundary)) break;
  }
  const beforeWindow = copyRange(before, from, beforeTo);
  const afterWindow = copyRange(after, from, afterTo);
  // Ranges, sections and note references can depend on other blocks in the
  // story. A note stays reachable while any one of its references survives.
  if (hasStoryBoundary(beforeWindow) || hasStoryBoundary(afterWindow))
    return { before, after };
  return { before: beforeWindow, after: afterWindow };
};

/**
 * Comments added by an edit batch also exist outside its tracked revisions.
 * Only the added comments' own anchors go; an older comment's reference
 * inside an added comment's range stays, without the added comment's mark.
 */
const withoutAddedComments = (
  node: PMNode,
  added: ReadonlySet<number>,
): PMNode | null => {
  const name = node.type.name;
  if (
    (name === "commentReference" || name === "commentRangeBoundary") &&
    ownCommentAnchorIds(node).some((id) => added.has(id))
  ) {
    return null;
  }
  if (name === "rangeAnchor") {
    const parsed = readRangeAnchorAttrs(node);
    if (parsed.ok) {
      const markers = parsed.value.markers.filter(
        (marker) =>
          (marker.type !== "commentRangeStart" &&
            marker.type !== "commentRangeEnd") ||
          !added.has(marker.id),
      );
      if (markers.length === 0) return null;
      if (markers.length !== parsed.value.markers.length)
        return node.type.create({ ...node.attrs, markers }, null, node.marks);
    }
  }
  const marks = node.marks.filter(
    (mark) =>
      mark.type.name !== "comment" || !added.has(mark.attrs["commentId"]),
  );
  if (node.isLeaf) return node.mark(marks);
  const children: PMNode[] = [];
  node.forEach((child) => {
    const kept = withoutAddedComments(child, added);
    if (kept !== null) children.push(kept);
  });
  return node.copy(Fragment.fromArray(children)).mark(marks);
};

const add = (items: Map<string, number>, key: string): void => {
  items.set(key, (items.get(key) ?? 0) + 1);
};

const isTrackedDeletion = (node: PMNode): boolean =>
  node.marks.some(({ type }) => type.name === "deletion");

/**
 * The protected content of `doc`; with `deletedOnly`, only the notes,
 * drawings and fields that are themselves pending deletion, which Accept All
 * removes as the reviewer sees it.
 */
const protectedContent = (doc: PMNode, deletedOnly = false): Inventory => {
  const inventory: Inventory = {
    note: new Map(),
    drawing: new Map(),
    field: new Map(),
    comment: new Map(),
  };
  const comments = new Set<number>();
  let hasComment = false;
  doc.descendants((node) => {
    if (deletedOnly && !isTrackedDeletion(node)) return true;
    // A text box stays a block of its own; its anchor in the carrier
    // paragraph is what a deletion marks.
    if (deletedOnly && node.type.name === "textBoxAnchor")
      add(
        inventory.drawing,
        JSON.stringify(["textBox", node.attrs["anchorId"]]),
      );
    const note = noteReferenceOf(node);
    if (note !== null)
      inventory.note.set(JSON.stringify([note.noteType, note.noteId]), 1);
    switch (node.type.name) {
      case "image":
        add(
          inventory.drawing,
          JSON.stringify([
            "image",
            node.attrs["rId"] ?? node.attrs["src"],
            node.attrs["docPrId"],
          ]),
        );
        break;
      case "shape":
        add(
          inventory.drawing,
          JSON.stringify(["shape", node.attrs["shapeId"]]),
        );
        break;
      case "textBox":
        add(
          inventory.drawing,
          JSON.stringify(["textBox", node.attrs["_docxAnchorId"]]),
        );
        break;
      case "field":
      case "structuredField":
        add(inventory.field, String(node.attrs["instruction"]));
        break;
    }
    if (deletedOnly) return true;
    for (const id of commentAnchorIdsForNode(node)) {
      comments.add(id);
      hasComment = true;
    }
    return true;
  });
  for (const id of comments) add(inventory.comment, String(id));
  if (hasComment) {
    // A surviving start alone does not keep a comment visible on save.
    visitCommentMarkers(proseDocToBlocks(doc, []), ({ item }) => {
      if (item.type === "commentReference")
        add(inventory.comment, `reference:${String(item.id)}`);
    });
  }
  return inventory;
};

const sameOriginal = (before: PMNode, after: PMNode): boolean =>
  before.content.eq(after.content) ||
  sameOriginalBlocks(
    proseDocToBlocks(before, []),
    proseDocToBlocks(after, []),
    new Set(),
  );

const checkTrackedBatchFidelity = ({
  before,
  after,
  createdCommentIds,
  retiredCommentIds = [],
}: TrackedBatchFidelityOptions): string | null => {
  // Comparisons verify against both input documents after matching all stories.
  if (comparisonApplying()) return null;
  const window = resolutionWindow(before.doc, after);
  if (window === null) return null;
  const plugins = pluginsForHeadlessRevisionResolution(before.plugins);
  let original = EditorState.create({ doc: window.before, plugins });
  let candidate = EditorState.create({ doc: window.after, plugins });
  // A suggestion can span separate paragraphs. Peel it from the whole story
  // before checking either ordinary review outcome.
  if (
    window.before !== before.doc &&
    (getSuggestions(original).length > 0 ||
      getSuggestions(candidate).length > 0)
  ) {
    original = EditorState.create({ doc: before.doc, plugins });
    candidate = EditorState.create({ doc: after, plugins });
  }
  original = rejectEverySuggestion(original);
  candidate = rejectEverySuggestion(candidate);
  const rejectedBefore = projectAllChangesInHeadlessState(original, "reject");
  const rejectedAfter = projectAllChangesInHeadlessState(candidate, "reject");
  if (rejectedBefore === null || rejectedAfter === null)
    return "Reject All could not resolve the tracked result.";
  const rejected =
    createdCommentIds.length === 0
      ? rejectedAfter.doc
      : withoutAddedComments(rejectedAfter.doc, new Set(createdCommentIds));
  if (rejected === null || !sameOriginal(rejectedBefore.doc, rejected))
    return "Reject All would change the original content or formatting.";

  const acceptedBefore = projectAllChangesInHeadlessState(original, "accept");
  const acceptedAfter = projectAllChangesInHeadlessState(candidate, "accept");
  if (acceptedBefore === null || acceptedAfter === null)
    return "Accept All could not resolve the tracked result.";
  const lost = missingProtectedContent(
    protectedContent(acceptedBefore.doc),
    protectedContent(acceptedAfter.doc),
    retiredCommentIds,
    newlyDeletedProtectedContent(
      protectedContent(original.doc, true),
      protectedContent(candidate.doc, true),
    ),
  );
  return lost === null ? null : `Accept All would lose a ${lost}.`;
};

/** Check both clean readings before any tracked batch can publish its changes. */
export const trackedBatchFidelityError = (
  options: TrackedBatchFidelityOptions,
): string | null => {
  const verdict = Result.try({
    try: () => checkTrackedBatchFidelity(options),
    catch: () => "The tracked result could not be verified.",
  });
  return verdict.isOk() ? verdict.value : verdict.error;
};
