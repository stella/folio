import { Result } from "better-result";
import { Fragment, type Node as PMNode } from "prosemirror-model";
import { EditorState } from "prosemirror-state";

import {
  removeCommentMarkers,
  visitCommentMarkers,
} from "../docx/commentAnchorIndex";
import {
  BLOCK_TREE_DESCENT,
  visitBlockTreeRecords,
  visitInlineContentSlots,
  visitParagraphRuns,
} from "../docx/paragraphTraversal";
import { consolidateParagraphContent } from "../docx/runConsolidator";
import { pluginsForHeadlessRevisionResolution } from "../internal/headlessRevisionResolutionGuard";
import { projectAllChangesInHeadlessState } from "../prosemirror/commands/comments";
import { proseDocToBlocks } from "../prosemirror/conversion/fromProseDoc";
import { isBlockBoundaryNode } from "../prosemirror/keptRangeBoundaries";
import {
  createDocumentNumberingPlugin,
  getStatedDocumentNumbering,
} from "../prosemirror/plugins/documentNumbering";
import {
  createDocumentStylesPlugin,
  getDocumentStyleResolver,
} from "../prosemirror/plugins/documentStyles";
import { readRangeAnchorAttrs } from "../prosemirror/rangeAnchorAttrs";
import { singletonManager } from "../prosemirror/schema";
import type { BlockContent, NumberingDefinitions } from "../types/document";
import { canonicalJson } from "../utils/canonicalJson";
import { commentAnchorIdsForNode } from "./comment-lifecycle";
import { comparisonApplying, noteReferenceOf } from "./note-references";
import { changedTopLevelWindow } from "./result-validation";

type TrackedBatchFidelityOptions = {
  before: EditorState;
  after: PMNode;
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
    visitCommentMarkers(proseDocToBlocks(doc, []), ({ item }) => {
      if (item.type === "commentReference")
        add(inventory.comment, `reference:${String(item.id)}`);
    });
  }
  return inventory;
};

const missingProtectedContent = (
  expected: Inventory,
  actual: Inventory,
): ProtectionKind | null => {
  return (
    protectionKinds.find((kind) => {
      for (const [key, count] of expected[kind]) {
        if ((actual[kind].get(key) ?? 0) < count) return true;
      }
      return false;
    }) ?? null
  );
};

const authoredBlock = (block: BlockContent): string => {
  const authored = structuredClone(block);
  visitBlockTreeRecords([authored], (record) => {
    if (record.type === "paragraph") {
      delete record.paraId;
      delete record.textId;
      delete record.renderedPageBreakBefore;
      delete record.listRendering;
      delete record.reviewCarrier;
      if (record.formatting && Object.keys(record.formatting).length === 0)
        delete record.formatting;
      record.content = consolidateParagraphContent(record.content);
    }
    return BLOCK_TREE_DESCENT.descend;
  });
  return canonicalJson(authored);
};

export const sameOriginalBlocks = (
  before: BlockContent[],
  after: BlockContent[],
  addedCommentIds: ReadonlySet<number>,
): boolean => {
  const restored = removeCommentMarkers(after, addedCommentIds);
  return (
    before.length === restored.length &&
    before.every((block, index) => {
      const other = restored.at(index);
      return (
        other !== undefined &&
        (block === other || authoredBlock(block) === authoredBlock(other))
      );
    })
  );
};

const sameOriginal = (before: PMNode, after: PMNode): boolean =>
  before.content.eq(after.content) ||
  sameOriginalBlocks(
    proseDocToBlocks(before, []),
    proseDocToBlocks(after, []),
    new Set(),
  );

const modelInventory = (blocks: readonly BlockContent[]): Inventory => {
  const inventory: Inventory = {
    note: new Map(),
    drawing: new Map(),
    field: new Map(),
    comment: new Map(),
  };
  visitBlockTreeRecords(blocks, (record) => {
    if (record.type !== "paragraph") return BLOCK_TREE_DESCENT.descend;
    visitInlineContentSlots(record, ({ item }) => {
      if (item.type === "simpleField" || item.type === "complexField")
        add(inventory.field, item.instruction);
    });
    visitParagraphRuns(record, ({ content }) => {
      for (const item of content) {
        switch (item.type) {
          case "footnoteRef":
          case "endnoteRef":
            add(inventory.note, JSON.stringify([item.type, item.id]));
            break;
          case "drawing":
            add(
              inventory.drawing,
              JSON.stringify([
                "drawing",
                item.image.id,
                item.image.rId ?? item.rawXml ?? item.image.src,
              ]),
            );
            break;
          case "shape":
            add(
              inventory.drawing,
              JSON.stringify([
                "shape",
                item.shape.id,
                item.shape.name,
                item.shape.shapeType,
              ]),
            );
            break;
        }
      }
    });
    return BLOCK_TREE_DESCENT.descend;
  });
  visitCommentMarkers(blocks, ({ item }) => {
    inventory.comment.set(String(item.id), 1);
    if (item.type === "commentReference")
      inventory.comment.set(`reference:${String(item.id)}`, 1);
  });
  return inventory;
};

export const protectedBlockLoss = (
  before: readonly BlockContent[],
  after: readonly BlockContent[],
): string | null => {
  const lost = missingProtectedContent(
    modelInventory(before),
    modelInventory(after),
  );
  return lost === null ? null : `Accept All would lose a ${lost}.`;
};

type ResolutionStateOptions = {
  doc: PMNode;
  styleResolver: ReturnType<typeof getDocumentStyleResolver>;
  numbering: NumberingDefinitions | null | undefined;
};

const resolutionState = ({
  doc,
  styleResolver,
  numbering,
}: ResolutionStateOptions): EditorState =>
  EditorState.create({
    doc,
    plugins: [
      createDocumentNumberingPlugin(numbering),
      ...pluginsForHeadlessRevisionResolution(singletonManager.getPlugins()),
      createDocumentStylesPlugin(styleResolver),
    ],
  });

const checkTrackedBatchFidelity = ({
  before,
  after,
  createdCommentIds,
}: TrackedBatchFidelityOptions): string | null => {
  // Comparisons verify against both input documents after matching all stories.
  if (comparisonApplying()) return null;
  const window = resolutionWindow(before.doc, after);
  if (window === null) return null;
  const styleResolver = getDocumentStyleResolver(before);
  const numbering = getStatedDocumentNumbering(before);
  const original = resolutionState({
    doc: window.before,
    styleResolver,
    numbering,
  });
  const candidate = resolutionState({
    doc: window.after,
    styleResolver,
    numbering,
  });
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
