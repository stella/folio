import { Result } from "better-result";
import { Fragment, type Node as PMNode } from "prosemirror-model";
import { EditorState } from "prosemirror-state";

import { visitCommentMarkers } from "../docx/commentAnchorIndex";
import { getCachedNumberingMap } from "../docx/numberingParser";
import {
  BLOCK_TREE_DESCENT,
  visitBlockTreeRecords,
} from "../docx/paragraphTraversal";
import { serializeDocumentBody } from "../docx/serializer/documentSerializer";
import type { RevisionResolutionMode } from "../internal/revisionResolutionInline";
import { resolveWholeStory } from "../internal/wholeStoryRevisionResolution";
import { proseDocToBlocks } from "../prosemirror/conversion/fromProseDoc";
import {
  isBlockBoundaryNode,
  normalizeKeptRangeBoundaries,
} from "../prosemirror/keptRangeBoundaries";
import { completeNumberingForDoc } from "../prosemirror/listInstanceReferences";
import {
  getDocumentNumbering,
  getStatedDocumentNumbering,
} from "../prosemirror/plugins/documentNumbering";
import { getDocumentStyleResolver } from "../prosemirror/plugins/documentStyleState";
import { readRangeAnchorAttrs } from "../prosemirror/rangeAnchorAttrs";
import type { NumberingDefinitions } from "../types/document";
import { commentAnchorIdsForNode } from "./comment-lifecycle";
import { comparisonApplying, noteReferenceOf } from "./note-references";
import { changedTopLevelWindow } from "./result-validation";

type TrackedBatchFidelityOptions = {
  before: EditorState;
  after: PMNode;
  numbering?: NumberingDefinitions | null;
  createdCommentIds: readonly number[];
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
  // Joint range repair and section ownership depend on carriers outside the
  // affected blocks when a boundary or section record reaches the window.
  if (hasStoryBoundary(beforeWindow) || hasStoryBoundary(afterWindow))
    return { before, after };
  return { before: beforeWindow, after: afterWindow };
};

/** Comments added by an edit batch also exist outside its tracked revisions. */
const withoutAddedComments = (
  node: PMNode,
  added: ReadonlySet<number>,
): PMNode | null => {
  const name = node.type.name;
  if (
    (name === "commentReference" || name === "commentRangeBoundary") &&
    commentAnchorIdsForNode(node).some((id) => added.has(id))
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

const protectionKinds = ["note", "drawing", "field", "comment"] as const;
type ProtectionKind = (typeof protectionKinds)[number];
type Inventory = Record<ProtectionKind, Map<string, number>>;

const add = (items: Map<string, number>, key: string): void => {
  items.set(key, (items.get(key) ?? 0) + 1);
};

const protectedContent = (doc: PMNode): Inventory => {
  const inventory: Inventory = {
    note: new Map(),
    drawing: new Map(),
    field: new Map(),
    comment: new Map(),
  };
  const comments = new Set<number>();
  let hasComment = false;
  doc.descendants((node) => {
    const note = noteReferenceOf(node);
    if (note !== null)
      add(inventory.note, JSON.stringify([note.noteType, note.noteId]));
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
    for (const id of commentAnchorIdsForNode(node)) {
      comments.add(id);
      hasComment = true;
    }
    return true;
  });
  for (const id of comments) add(inventory.comment, String(id));
  if (hasComment) {
    // A surviving start alone does not keep a comment visible on save.
    visitCommentMarkers(proseDocToBlocks(doc), ({ item }) => {
      if (item.type === "commentReference")
        add(inventory.comment, `reference:${String(item.id)}`);
    });
  }
  return inventory;
};

const lostContent = (before: PMNode, after: PMNode): ProtectionKind | null => {
  const expected = protectedContent(before);
  const actual = protectedContent(after);
  return (
    protectionKinds.find((kind) => {
      for (const [key, count] of expected[kind]) {
        if ((actual[kind].get(key) ?? 0) < count) return true;
      }
      return false;
    }) ?? null
  );
};

const originalXml = (doc: PMNode): string => {
  const content = proseDocToBlocks(doc, []);
  visitBlockTreeRecords(content, (record) => {
    if (record.type === "paragraph") {
      // Paragraph identifiers and layout hints can be regenerated on replay.
      delete record.paraId;
      delete record.textId;
      delete record.renderedPageBreakBefore;
    }
    return BLOCK_TREE_DESCENT.descend;
  });
  return serializeDocumentBody({ content });
};

const sameOriginal = (before: PMNode, after: PMNode): boolean =>
  before.content.eq(after.content) ||
  originalXml(before) === originalXml(after);

type ResolveProjectionOptions = {
  doc: PMNode;
  mode: RevisionResolutionMode;
  styleResolver: ReturnType<typeof getDocumentStyleResolver>;
  numbering: ReturnType<typeof getDocumentNumbering>;
};

const resolveProjection = (
  options: ResolveProjectionOptions,
): PMNode | null => {
  const { resolved, failed, keptBoundaries } = resolveWholeStory(options);
  if (failed) return null;
  if (resolved === options.doc) return resolved;
  const tr = EditorState.create({ doc: resolved }).tr;
  normalizeKeptRangeBoundaries(tr, keptBoundaries);
  return tr.doc;
};

const checkTrackedBatchFidelity = ({
  before,
  after,
  numbering,
  createdCommentIds,
}: TrackedBatchFidelityOptions): string | null => {
  // Comparisons verify against both input documents after matching all stories.
  if (comparisonApplying()) return null;
  const window = resolutionWindow(before.doc, after);
  if (window === null) return null;
  const styleResolver = getDocumentStyleResolver(before);
  const original = {
    doc: window.before,
    styleResolver,
    numbering: getDocumentNumbering(before),
  };
  const definitions =
    numbering === undefined
      ? completeNumberingForDoc(
          getStatedDocumentNumbering(before) ?? undefined,
          window.after,
        )
      : numbering;
  const candidate = {
    doc: window.after,
    styleResolver,
    numbering: definitions ? getCachedNumberingMap(definitions) : null,
  };
  const rejectedBefore = resolveProjection({ ...original, mode: "reject" });
  const rejectedAfter = resolveProjection({ ...candidate, mode: "reject" });
  if (rejectedBefore === null || rejectedAfter === null)
    return "Reject All could not resolve the tracked result.";
  const rejected =
    createdCommentIds.length === 0
      ? rejectedAfter
      : withoutAddedComments(rejectedAfter, new Set(createdCommentIds));
  if (rejected === null || !sameOriginal(rejectedBefore, rejected))
    return "Reject All would change the original content or formatting.";

  const acceptedBefore = resolveProjection({ ...original, mode: "accept" });
  const acceptedAfter = resolveProjection({ ...candidate, mode: "accept" });
  if (acceptedBefore === null || acceptedAfter === null)
    return "Accept All could not resolve the tracked result.";
  const lost = lostContent(acceptedBefore, acceptedAfter);
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
