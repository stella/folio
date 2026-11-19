import { Result } from "better-result";
import { Fragment, type Node as PMNode } from "prosemirror-model";
import { EditorState } from "prosemirror-state";

import { mergeRunContent } from "@stll/docx-core/ops";

import {
  removeCommentMarkers,
  visitCommentMarkers,
} from "../docx/commentAnchorIndex";
import { paragraphFormattingWithAuthoredIndentation } from "../docx/paragraphPropertySource";
import {
  BLOCK_TREE_DESCENT,
  visitBlockTreeRecords,
  visitInlineContentSlots,
  visitParagraphRuns,
} from "../docx/paragraphTraversal";
import {
  consolidateParagraphContent,
  formattingEquals,
} from "../docx/runConsolidator";
import { serializeDocumentPart } from "../docx/serializer/documentSerializer";
import { WORDPROCESSINGML_NAMESPACE_URIS } from "../docx/xmlParser";
import { pluginsForHeadlessRevisionResolution } from "../internal/headlessRevisionResolutionGuard";
import {
  getSuggestions,
  projectAllChangesInHeadlessState,
  rejectEverySuggestion,
} from "../prosemirror/commands/comments";
import { proseDocToBlocks } from "../prosemirror/conversion/fromProseDoc";
import { isBlockBoundaryNode } from "../prosemirror/keptRangeBoundaries";
import { readRangeAnchorAttrs } from "../prosemirror/rangeAnchorAttrs";
import type {
  BlockContent,
  Paragraph,
  ParagraphFormatting,
} from "../types/document";
import { commentAnchorIdsForNode } from "./comment-lifecycle";
import { comparisonApplying, noteReferenceOf } from "./note-references";
import { changedTopLevelWindow } from "./result-validation";

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
  retiredCommentIds: readonly number[],
): ProtectionKind | null => {
  const retiredCommentKeys = new Set(
    retiredCommentIds.flatMap((id) => [String(id), `reference:${String(id)}`]),
  );
  return (
    protectionKinds.find((kind) => {
      for (const [key, count] of expected[kind]) {
        if (kind === "comment" && retiredCommentKeys.has(key)) continue;
        if ((actual[kind].get(key) ?? 0) < count) return true;
      }
      return false;
    }) ?? null
  );
};

// A resolved paragraph break can change the package's editing-session owner.
const PARAGRAPH_SESSION_ATTRIBUTES: ReadonlySet<string> = new Set([
  "rsidR",
  "rsidRPr",
  "rsidDel",
  "rsidP",
  "rsidRDefault",
]);

/**
 * `content` with adjacent runs of the same properties as one run: where a
 * run boundary falls is no content or formatting of the paragraph, so a
 * comment's markers taken out between two runs (around a line break, say)
 * leave the paragraph the original was. A run carrying a property revision
 * keeps its boundaries, which delimit the revision.
 */
const withoutRunBoundaries = (
  content: Paragraph["content"],
): Paragraph["content"] => {
  const joined: Paragraph["content"] = [];
  for (const item of content) {
    const previous = joined.at(-1);
    if (
      previous?.type === "run" &&
      item.type === "run" &&
      previous.propertyChanges === undefined &&
      item.propertyChanges === undefined &&
      formattingEquals(previous.formatting, item.formatting) &&
      JSON.stringify(previous.preservedAttributes) ===
        JSON.stringify(item.preservedAttributes)
    ) {
      joined[joined.length - 1] = {
        ...previous,
        content: mergeRunContent(previous.content, item.content),
      };
      continue;
    }
    joined.push(item);
  }
  return joined;
};

const authoredBlock = (block: BlockContent): string => {
  // A native reject restores authored properties, while the live model still
  // projects list indentation. Read the authored values before cloning drops
  // the provenance that distinguishes inherited indentation from direct edits.
  const formatting: (ParagraphFormatting | undefined)[] = [];
  visitBlockTreeRecords([block], (record) => {
    if (record.type === "paragraph")
      formatting.push(paragraphFormattingWithAuthoredIndentation(record));
    return BLOCK_TREE_DESCENT.descend;
  });
  const authored = structuredClone(block);
  let paragraphIndex = 0;
  visitBlockTreeRecords([authored], (record) => {
    if (record.type === "paragraph") {
      const properties = formatting.at(paragraphIndex);
      paragraphIndex += 1;
      if (properties === undefined) delete record.formatting;
      else record.formatting = properties;
      delete record.paraId;
      delete record.textId;
      delete record.renderedPageBreakBefore;
      delete record.listRendering;
      delete record.reviewCarrier;
      if (record.preservedAttributes !== undefined) {
        record.preservedAttributes = record.preservedAttributes.filter(
          ({ name, namespace }) =>
            !WORDPROCESSINGML_NAMESPACE_URIS.has(namespace ?? "") ||
            !PARAGRAPH_SESSION_ATTRIBUTES.has(name),
        );
      }
      if (record.formatting && Object.keys(record.formatting).length === 0)
        delete record.formatting;
      record.content = withoutRunBoundaries(
        consolidateParagraphContent(record.content),
      );
    }
    return BLOCK_TREE_DESCENT.descend;
  });
  return serializeDocumentPart({
    package: { document: { content: [authored] } },
  }).xml;
};

export const sameOriginalBlocks = (
  before: readonly BlockContent[],
  after: readonly BlockContent[],
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
            inventory.note.set(JSON.stringify([item.type, item.id]), 1);
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

/** The kind of protected content missing from the accepted blocks. */
export const missingProtectedBlockKind = (
  before: readonly BlockContent[],
  after: readonly BlockContent[],
  retiredCommentIds: readonly number[] = [],
): ProtectionKind | null =>
  missingProtectedContent(
    modelInventory(before),
    modelInventory(after),
    retiredCommentIds,
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
