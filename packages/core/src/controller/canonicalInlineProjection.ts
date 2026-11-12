import {
  panic,
  Result,
  TaggedError,
  type TaggedErrorClass,
} from "better-result";
import type { Node as PMNode } from "prosemirror-model";

import { hasIllegalXmlCharacters } from "@stll/docx-core";
import {
  defaultInsertionGap,
  inlineLeafSpans,
  paragraphLogicalText,
} from "@stll/docx-core/ops";

import { readNoteMarkerAttrs } from "../internal/noteMarkerAttrs";
import {
  readCommentMarkAttrs,
  readFootnoteRefMarkAttrs,
  readPreservedXmlAttrs,
} from "../prosemirror/attrs/index";
import {
  bookmarkMarkerFromAttrs,
  readBookmarkBoundaryAttrs,
} from "../prosemirror/bookmarkBoundaryAttrs";
import { readCommentRangeBoundaryAttrs } from "../prosemirror/commentRangeBoundaryAttrs";
import { HYPHEN_TEXT_CARRIERS } from "../prosemirror/conversion/hyphenTextCarriers";
import { COMMENT_RANGE_BOUNDARY_NODE_NAME } from "../prosemirror/extensions/nodes/CommentRangeBoundaryExtension";
import { RENDERED_PAGE_BREAK_NODE_NAME } from "../prosemirror/extensions/nodes/RenderedPageBreakExtension";
import { readMoveRangeBoundaryAttrs } from "../prosemirror/moveRangeBoundaryAttrs";
import { readRangeAnchorAttrs } from "../prosemirror/rangeAnchorAttrs";
import { PRESERVED_XML_LEVELS } from "../prosemirror/schema/nodes";
import { TRACKED_RUN_INLINE_ATOM_DISPOSITIONS } from "../prosemirror/trackedRunInlineAtoms";
import type {
  NoteReferenceContent,
  Paragraph,
  ParagraphContent,
  RunContent,
} from "../types/document";

type SourceKind = ParagraphContent["type"] | RunContent["type"];
type SourcePolicy = "container" | "leaf" | "text" | "rawField";
export const CANONICAL_SOURCE_INLINE_POLICIES: {
  readonly run: "container";
  readonly hyperlink: "container";
  readonly insertion: "container";
  readonly deletion: "container";
  readonly moveFrom: "container";
  readonly moveTo: "container";
  readonly inlineWrapper: "container";
  readonly inlineSdt: "container";
  readonly simpleField: "leaf";
  readonly complexField: "leaf";
  readonly mathEquation: "leaf";
  readonly preservedInline: "leaf";
  readonly bookmarkStart: "leaf";
  readonly bookmarkEnd: "leaf";
  readonly commentRangeStart: "leaf";
  readonly commentRangeEnd: "leaf";
  readonly commentReference: "leaf";
  readonly moveFromRangeStart: "leaf";
  readonly moveFromRangeEnd: "leaf";
  readonly moveToRangeStart: "leaf";
  readonly moveToRangeEnd: "leaf";
  readonly text: "text";
  readonly tab: "leaf";
  readonly break: "leaf";
  readonly symbol: "leaf";
  readonly footnoteRef: "leaf";
  readonly endnoteRef: "leaf";
  readonly noteMarker: "leaf";
  readonly fieldChar: "rawField";
  readonly instrText: "rawField";
  readonly softHyphen: "leaf";
  readonly noBreakHyphen: "leaf";
  readonly renderedPageBreak: "leaf";
  readonly preservedXml: "leaf";
  readonly drawing: "leaf";
  readonly shape: "leaf";
} = {
  run: "container",
  hyperlink: "container",
  insertion: "container",
  deletion: "container",
  moveFrom: "container",
  moveTo: "container",
  inlineWrapper: "container",
  inlineSdt: "container",
  simpleField: "leaf",
  complexField: "leaf",
  mathEquation: "leaf",
  preservedInline: "leaf",
  bookmarkStart: "leaf",
  bookmarkEnd: "leaf",
  commentRangeStart: "leaf",
  commentRangeEnd: "leaf",
  commentReference: "leaf",
  moveFromRangeStart: "leaf",
  moveFromRangeEnd: "leaf",
  moveToRangeStart: "leaf",
  moveToRangeEnd: "leaf",
  text: "text",
  tab: "leaf",
  break: "leaf",
  symbol: "leaf",
  footnoteRef: "leaf",
  endnoteRef: "leaf",
  noteMarker: "leaf",
  fieldChar: "rawField",
  instrText: "rawField",
  softHyphen: "leaf",
  noBreakHyphen: "leaf",
  renderedPageBreak: "leaf",
  preservedXml: "leaf",
  drawing: "leaf",
  shape: "leaf",
} as const satisfies Record<SourceKind, SourcePolicy>;

type NativePolicy =
  | "text"
  | "unit"
  | "markers"
  | "noteMarker"
  | "capture"
  | "control"
  | "refused";
export const CANONICAL_NATIVE_INLINE_POLICIES: {
  readonly text: "text";
  readonly bookmarkBoundary: "markers";
  readonly commentRangeBoundary: "markers";
  readonly moveRangeBoundary: "markers";
  readonly rangeAnchor: "markers";
  readonly noteMarker: "noteMarker";
  readonly preservedXml: "capture";
  readonly field: "unit";
  readonly structuredField: "unit";
  readonly hardBreak: "unit";
  readonly image: "unit";
  readonly math: "unit";
  readonly pageBreakRun: "unit";
  readonly shape: "unit";
  readonly symbol: "unit";
  readonly tab: "unit";
  readonly textBoxAnchor: "unit";
  readonly commentReference: "unit";
  readonly renderedPageBreak: "markers";
  readonly sdt: "control";
} = {
  text: "text",
  bookmarkBoundary: "markers",
  commentRangeBoundary: "markers",
  moveRangeBoundary: "markers",
  rangeAnchor: "markers",
  noteMarker: "noteMarker",
  preservedXml: "capture",
  field: "unit",
  structuredField: "unit",
  hardBreak: "unit",
  image: "unit",
  math: "unit",
  pageBreakRun: "unit",
  shape: "unit",
  symbol: "unit",
  tab: "unit",
  textBoxAnchor: "unit",
  commentReference: "unit",
  renderedPageBreak: "markers",
  sdt: "control",
} as const satisfies Record<
  keyof typeof TRACKED_RUN_INLINE_ATOM_DISPOSITIONS | "sdt",
  NativePolicy
>;
const nativePolicies = new Map<string, NativePolicy>(
  Object.entries(CANONICAL_NATIVE_INLINE_POLICIES),
);

const InlineProjectionErrorBase: TaggedErrorClass<"InlineProjectionError"> =
  TaggedError("InlineProjectionError");

class InlineProjectionError extends InlineProjectionErrorBase<{
  message: string;
}> {}
const refuse = (message: string) =>
  Result.err(new InlineProjectionError({ message }));

/** Every source variant makes a decision before conversion can discard its positions. */
export const canonicalInlineSourceIssue = (
  items: readonly (ParagraphContent | RunContent)[],
  context: "paragraph" | "fieldPayload" = "paragraph",
): string | null => {
  for (const item of items) {
    const policy = CANONICAL_SOURCE_INLINE_POLICIES[item.type];
    switch (policy) {
      case "rawField":
        if (context === "fieldPayload") break;
        return "Canonical editing cannot preserve unstructured field instructions; a structured field is required.";
      case "text":
        if (
          item.type === "text" &&
          (hasIllegalXmlCharacters(item.text) || /[\t\r\n]/u.test(item.text))
        )
          return "Canonical text requires valid XML characters and explicit tab or break elements.";
        break;
      case "container": {
        const children = sourceContainerChildren(item);
        const issue = canonicalInlineSourceIssue(children, context);
        if (issue !== null) return issue;
        break;
      }
      case "leaf": {
        // Fields are one positional unit, but their cached payload still needs validation.
        let children: readonly (ParagraphContent | RunContent)[] = [];
        if (item.type === "simpleField") children = item.content;
        if (item.type === "complexField")
          children = [...item.fieldCode, ...item.fieldResult];
        const issue = canonicalInlineSourceIssue(children, "fieldPayload");
        if (issue !== null) return issue;
        break;
      }
      default:
        policy satisfies never;
    }
  }
  return null;
};

const sourceContainerChildren = (item: ParagraphContent | RunContent) => {
  switch (item.type) {
    case "hyperlink":
      return item.children;
    case "run":
    case "insertion":
    case "deletion":
    case "moveFrom":
    case "moveTo":
    case "inlineWrapper":
    case "inlineSdt":
      return item.content;
    default:
      return panic(
        `Source policy incorrectly classified ${item.type} as a container.`,
      );
  }
};

type NativeCell =
  | { type: "text"; text: string; from: number; to: number; node: PMNode }
  | { type: "unit"; from: number; to: number; node: PMNode }
  | {
      type: "markedUnit";
      sourceType: NoteReferenceContent["type"];
      label: string;
      from: number;
      to: number;
      node: PMNode;
    }
  | {
      type: "zeroWidth";
      sources: readonly (ParagraphContent | RunContent)[];
      from: number;
      to: number;
      node: PMNode;
    };

/** Native positions and deletion geometry read the same inline-cell classification. */
export const canonicalNativeCells = (
  paragraph: PMNode,
): Result<NativeCell[], InlineProjectionError> => {
  const cells: NativeCell[] = [];
  let issue: string | undefined;
  const visit = (node: PMNode, from: number): void => {
    if (issue !== undefined) return;
    const to = from + node.nodeSize;
    const policy = nativePolicies.get(node.type.name);
    switch (policy) {
      case "text": {
        const reference = node.marks.find(
          ({ type }) => type.name === "footnoteRef",
        );
        if (!reference) {
          cells.push({ type: "text", text: node.text ?? "", from, to, node });
          return;
        }
        const attrs = readFootnoteRefMarkAttrs(reference);
        if (!attrs.ok) {
          issue =
            "Canonical editing requires valid note-reference mark attributes.";
          return;
        }
        const label = String(attrs.value.id);
        const text = node.text ?? "";
        if (
          label.length === 0 ||
          text.length % label.length !== 0 ||
          text !== label.repeat(text.length / label.length)
        ) {
          issue =
            "Canonical editing cannot map the rendered note-reference labels.";
          return;
        }
        // Identical marks merge adjacent labels into one native text node.
        for (let offset = 0; offset < text.length; offset += label.length) {
          cells.push({
            type: "markedUnit",
            sourceType:
              attrs.value.noteType === "endnote" ? "endnoteRef" : "footnoteRef",
            label,
            from: from + offset,
            to: from + offset + label.length,
            node,
          });
        }
        return;
      }
      case "unit":
        cells.push({ type: "unit", from, to, node });
        return;
      case "markers": {
        if (node.type.name === "bookmarkBoundary") {
          const attrs = readBookmarkBoundaryAttrs(node);
          if (!attrs.ok) {
            issue =
              "Canonical editing requires valid bookmark boundary attributes.";
            return;
          }
          cells.push({
            type: "zeroWidth",
            sources: [bookmarkMarkerFromAttrs(attrs.value)],
            from,
            to,
            node,
          });
        } else if (node.type.name === RENDERED_PAGE_BREAK_NODE_NAME) {
          // A cached pagination boundary: a zero-width leaf in its run.
          cells.push({
            type: "zeroWidth",
            sources: [{ type: "renderedPageBreak" }],
            from,
            to,
            node,
          });
        } else if (node.type.name === "moveRangeBoundary") {
          const attrs = readMoveRangeBoundaryAttrs(node);
          if (!attrs.ok) {
            issue =
              "Canonical editing requires valid move-range boundary attributes.";
            return;
          }
          cells.push({
            type: "zeroWidth",
            sources: [attrs.value],
            from,
            to,
            node,
          });
        } else if (node.type.name === "commentRangeBoundary") {
          const attrs = readCommentRangeBoundaryAttrs(node);
          if (!attrs.ok) {
            issue =
              "Canonical editing requires valid comment range boundary attributes.";
            return;
          }
          cells.push({
            type: "zeroWidth",
            sources: [attrs.value.marker],
            from,
            to,
            node,
          });
        } else {
          const attrs = readRangeAnchorAttrs(node);
          if (!attrs.ok) {
            issue =
              "Canonical editing requires a valid collapsed range anchor.";
            return;
          }
          cells.push({
            type: "zeroWidth",
            sources: [attrs.value.start, attrs.value.end],
            from,
            to,
            node,
          });
        }
        return;
      }
      case "noteMarker": {
        const attrs = readNoteMarkerAttrs(node);
        if (!attrs.ok) {
          issue = "Canonical editing requires a valid automatic note marker.";
          return;
        }
        cells.push({
          type: "zeroWidth",
          sources: [{ type: "noteMarker", kind: attrs.value.kind }],
          from,
          to,
          node,
        });
        return;
      }
      case "capture": {
        const attrs = readPreservedXmlAttrs(node);
        if (!attrs.ok) {
          issue = "Canonical editing requires valid preserved inline markup.";
          return;
        }
        if (
          attrs.value.level === PRESERVED_XML_LEVELS.inline &&
          attrs.value.text === ""
        )
          cells.push({
            type: "zeroWidth",
            sources: [
              { type: "preservedInline", xml: attrs.value.xml, text: "" },
            ],
            from,
            to,
            node,
          });
        else cells.push({ type: "unit", from, to, node });
        return;
      }
      case "control":
        // An inline content control's edges are editor positions with no
        // source leaf: each stands for the gap at that edge of its content.
        cells.push({
          type: "zeroWidth",
          sources: [],
          from,
          to: from + 1,
          node,
        });
        node.forEach((child, offset) => visit(child, from + 1 + offset));
        cells.push({ type: "zeroWidth", sources: [], from: to - 1, to, node });
        return;
      case "refused":
        issue = `Canonical editing cannot preserve ${node.type.name} boundary positions.`;
        return;
      case undefined:
        issue = `Canonical editing has no position mapping for inline ${node.type.name}.`;
        return;
      default:
        policy satisfies never;
    }
  };
  paragraph.forEach((node, from) => visit(node, from));
  return issue === undefined ? Result.ok(cells) : refuse(issue);
};

export type CanonicalInlineGap = { position: number; zeroWidthBefore: number };
/** The source ordinals one editor position stands for: from `first` to `last`. */
export type CanonicalSeamExtent = { first: number; last: number };
type InlineProjection = {
  text: string;
  boundaries: CanonicalInlineGap[][];
  /** Parallel to `boundaries`: each gap's whole seam. */
  seams: CanonicalSeamExtent[][];
};
const sameSourceMarker = (
  left: ParagraphContent | RunContent,
  right: ParagraphContent | RunContent,
) => {
  switch (left.type) {
    case "bookmarkStart":
    case "bookmarkEnd":
    case "commentRangeStart":
    case "commentRangeEnd":
    case "moveFromRangeStart":
    case "moveFromRangeEnd":
    case "moveToRangeStart":
    case "moveToRangeEnd":
      return right.type === left.type && left.id === right.id;
    case "noteMarker":
      return right.type === "noteMarker" && left.kind === right.kind;
    case "renderedPageBreak":
      return right.type === "renderedPageBreak";
    case "preservedInline":
      return (
        right.type === "preservedInline" &&
        left.xml === right.xml &&
        left.text === right.text
      );
    default:
      return false;
  }
};

/** Whether a story holds an id: all a paragraph's projection asks of a story id set. */
export type StoryIdMembership = Pick<ReadonlySet<number>, "has">;

/** The story's comments as one paragraph's projection reads them. */
export type CanonicalInlineCommentContext = {
  rangedIds: StoryIdMembership;
  pointIds: StoryIdMembership;
  /** The comment ranges open where the paragraph starts. */
  openIds: ReadonlySet<number>;
};

type CanonicalInlineProjectionArgs = {
  source: Paragraph;
  paragraph: PMNode;
  pairedBookmarkIds: StoryIdMembership;
  commentContext?: CanonicalInlineCommentContext;
};

/**
 * What each input of a paragraph's projection depends on: the paragraph
 * alone (its source records and its editor node), or the story around it.
 * A cached projection is reused only for the same paragraph inputs, and only
 * where every story fact it read still holds. Total over the inputs: an input
 * added to them does not compile until it is placed in one half or the other.
 */
const INLINE_INPUT_DEPENDENCY = {
  source: "paragraph",
  paragraph: "paragraph",
  pairedBookmarkIds: "story",
  commentContext: "story",
} as const satisfies Record<
  keyof CanonicalInlineProjectionArgs,
  "paragraph" | "story"
>;
type InlineDependency = typeof INLINE_INPUT_DEPENDENCY;
type ParagraphInputField = {
  [K in keyof InlineDependency]: InlineDependency[K] extends "paragraph"
    ? K
    : never;
}[keyof InlineDependency];
type StoryInputField = Exclude<
  keyof CanonicalInlineProjectionArgs,
  ParagraphInputField
>;
const PARAGRAPH_INPUT_FIELDS = (
  Object.keys(INLINE_INPUT_DEPENDENCY) as (keyof InlineDependency)[]
).filter(
  (field): field is ParagraphInputField =>
    INLINE_INPUT_DEPENDENCY[field] === "paragraph",
);
/** The story facts one paragraph's projection reads, each through its recorder. */
type StoryInlineInputs = Pick<CanonicalInlineProjectionArgs, StoryInputField>;

type CommentTransition = { starts: Set<number>; ends: Set<number> };

type CommentMarkTransitionsArgs = Pick<
  CanonicalInlineProjectionArgs,
  "source" | "paragraph" | "commentContext"
> & { explicitCommentIds: ReadonlySet<number> };

/** Marks carry nonempty ranges; explicit range anchors carry empty ranges. */
const commentMarkTransitions = ({
  source,
  paragraph,
  commentContext,
  explicitCommentIds,
}: CommentMarkTransitionsArgs) => {
  const spans = inlineLeafSpans(source.content);
  // The story's sets are asked about ids, never walked: a paragraph's own
  // markers extend them here.
  const ownRangedIds = new Set<number>();
  const ownPointIds = new Set<number>();
  const finalIds = new Set(commentContext?.openIds);
  for (const { node } of spans) {
    if (node.type === "commentRangeStart" || node.type === "commentRangeEnd")
      ownRangedIds.add(node.id);
    if (node.type === "commentReference") ownPointIds.add(node.id);
    if (node.type === "commentRangeStart") finalIds.add(node.id);
    if (node.type === "commentRangeEnd") finalIds.delete(node.id);
  }
  const transitions = new Map<number, CommentTransition>();
  let previous = new Set(
    [...(commentContext?.openIds ?? [])].filter(
      (id) => !explicitCommentIds.has(id),
    ),
  );
  let issue: string | undefined;
  const record = (position: number, current: Set<number>) => {
    const starts = new Set([...current].filter((id) => !previous.has(id)));
    const ends = new Set([...previous].filter((id) => !current.has(id)));
    if (starts.size > 0 || ends.size > 0)
      transitions.set(position, { starts, ends });
    previous = current;
  };
  paragraph.forEach((node, position) => {
    if (issue !== undefined) return;
    const current = new Set<number>();
    for (const mark of node.marks) {
      if (mark.type.name !== "comment") continue;
      const attrs = readCommentMarkAttrs(mark);
      if (!attrs.ok) {
        issue = "Canonical editing requires valid comment mark attributes.";
        return;
      }
      const id = attrs.value.commentId;
      if (
        ownRangedIds.has(id) ||
        (commentContext?.rangedIds.has(id) ?? false)
      ) {
        if (!explicitCommentIds.has(id)) current.add(id);
      } else if (
        !ownPointIds.has(id) &&
        !(commentContext?.pointIds.has(id) ?? false)
      ) {
        issue =
          "Canonical editing cannot attribute a comment mark to its source.";
        return;
      }
    }
    record(position, current);
  });
  // A range may continue into the next paragraph, whose marks verify this context.
  record(
    paragraph.content.size,
    new Set([...finalIds].filter((id) => !explicitCommentIds.has(id))),
  );
  return issue === undefined ? Result.ok(transitions) : refuse(issue);
};

/** Bind rendered gaps to the operation owner's source leaf ordinals, including collapsed pairs. */
export const projectCanonicalInline = ({
  source,
  paragraph,
  pairedBookmarkIds,
  commentContext,
}: CanonicalInlineProjectionArgs): Result<
  InlineProjection,
  InlineProjectionError
> => {
  const native = canonicalNativeCells(paragraph);
  if (native.isErr()) return native;
  const explicitCommentIds = new Set<number>();
  for (const cell of native.value) {
    if (
      cell.type !== "zeroWidth" ||
      cell.node.type.name !== COMMENT_RANGE_BOUNDARY_NODE_NAME
    )
      continue;
    for (const marker of cell.sources) {
      if (
        marker.type === "commentRangeStart" ||
        marker.type === "commentRangeEnd"
      )
        explicitCommentIds.add(marker.id);
    }
  }
  const openExplicitCommentIds = new Set(
    [...(commentContext?.openIds ?? [])].filter((id) =>
      explicitCommentIds.has(id),
    ),
  );
  const comments = commentMarkTransitions({
    source,
    paragraph,
    explicitCommentIds,
    ...(commentContext === undefined ? {} : { commentContext }),
  });
  if (comments.isErr()) return comments;
  const spans = inlineLeafSpans(source.content);
  const erasedBookmarks = new Set<ParagraphContent | RunContent>();
  const collectErasedBookmarks = (
    items: readonly (ParagraphContent | RunContent)[],
    context: "paragraph" | "nested",
  ): void => {
    for (const item of items) {
      if (item.type === "bookmarkStart" || item.type === "bookmarkEnd") {
        if (context === "paragraph" && !pairedBookmarkIds.has(item.id))
          erasedBookmarks.add(item);
        continue;
      }
      if (CANONICAL_SOURCE_INLINE_POLICIES[item.type] !== "container") continue;
      collectErasedBookmarks(
        sourceContainerChildren(item),
        item.type === "run" || item.type === "inlineWrapper"
          ? context
          : "nested",
      );
    }
  };
  collectErasedBookmarks(source.content, "paragraph");
  // Comment ranges render as marks; only a range collapsed inside this paragraph
  // keeps native range-anchor markers. A boundary of a range the native text
  // carries as a comment mark is a zero-width source leaf with no native cell,
  // like an unpaired bookmark; any other boundary stays unmappable and refuses.
  const markedCommentIds = new Set<number>();
  paragraph.descendants((node) => {
    for (const mark of node.marks) {
      if (mark.type.name !== "comment") continue;
      const attrs = readCommentMarkAttrs(mark);
      if (attrs.ok) markedCommentIds.add(attrs.value.commentId);
    }
  });
  const commentRangeOffsets = new Map<
    number,
    { start?: number; end?: number }
  >();
  for (const { node, before } of spans) {
    if (node.type !== "commentRangeStart" && node.type !== "commentRangeEnd")
      continue;
    const offsets = commentRangeOffsets.get(node.id) ?? {};
    if (node.type === "commentRangeStart") offsets.start = before.offset;
    else offsets.end = before.offset;
    commentRangeOffsets.set(node.id, offsets);
  }
  const erasedCommentMarkers = new Set<ParagraphContent | RunContent>();
  for (const { node } of spans) {
    if (node.type !== "commentRangeStart" && node.type !== "commentRangeEnd")
      continue;
    const offsets = commentRangeOffsets.get(node.id);
    const collapsed =
      offsets?.start !== undefined &&
      offsets.end !== undefined &&
      offsets.start === offsets.end;
    if (
      !collapsed &&
      markedCommentIds.has(node.id) &&
      !explicitCommentIds.has(node.id)
    )
      erasedCommentMarkers.add(node);
  }
  const text = paragraphLogicalText(source);
  const boundaries: CanonicalInlineGap[][] = Array.from(
    { length: text.length + 1 },
    () => [],
  );
  let index = 0;
  let consumed = 0;
  const insertionOrdinals = new Map<number, number>();
  const insertionOrdinalAt = (offset: number): number => {
    const known = insertionOrdinals.get(offset);
    if (known !== undefined) return known;
    const ordinal = defaultInsertionGap(source.content, offset).zeroWidthBefore;
    insertionOrdinals.set(offset, ordinal);
    return ordinal;
  };
  // Parallel to `boundaries`: a gap at an offset and editor position is one
  // entry of its offset's list, and its seam sits at the same index, so a
  // gap finds its seam by the entry it already looked up.
  const seams: CanonicalSeamExtent[][] = Array.from(
    { length: text.length + 1 },
    () => [],
  );
  const addGap = (
    position: number,
    gap: { offset: number; zeroWidthBefore: number },
  ) => {
    const gaps = boundaries.at(gap.offset);
    const extents = seams.at(gap.offset);
    if (!gaps || !extents)
      panic(
        `Source leaf offset ${gap.offset} exceeds paragraph ${source.paraId}.`,
      );
    const existing = gaps.findIndex((entry) => entry.position === position);
    const entry = { position, zeroWidthBefore: gap.zeroWidthBefore };
    if (existing < 0) {
      gaps.push(entry);
      extents.push({ first: gap.zeroWidthBefore, last: gap.zeroWidthBefore });
      return;
    }
    const extent = extents[existing] ?? panic("Missing canonical seam extent");
    extent.first = Math.min(extent.first, gap.zeroWidthBefore);
    extent.last = Math.max(extent.last, gap.zeroWidthBefore);
    // Erased source leaves share a native seam. The seam stands for the gap
    // typed text takes there (`defaultInsertionGap`: after closing markers,
    // before opening ones), so a caret the editor places and a caret an
    // edit returns name the same source gap.
    const preferred = insertionOrdinalAt(gap.offset);
    const kept = gaps[existing] ?? panic("Missing canonical seam gap");
    if (
      Math.abs(entry.zeroWidthBefore - preferred) <
      Math.abs(kept.zeroWidthBefore - preferred)
    )
      gaps[existing] = entry;
  };
  const consumeErasedLeaves = (position: number | null) => {
    while (consumed === 0) {
      const span = spans.at(index);
      if (
        position !== null &&
        span &&
        (span.node.type === "commentRangeStart" ||
          span.node.type === "commentRangeEnd") &&
        !explicitCommentIds.has(span.node.id)
      ) {
        const transition = comments.value.get(position);
        const ids =
          span.node.type === "commentRangeStart"
            ? transition?.starts
            : transition?.ends;
        if (ids?.delete(span.node.id)) {
          index += 1;
          addGap(position, span.after);
          continue;
        }
      }
      if (
        !span ||
        span.before.offset !== span.after.offset ||
        !(
          erasedBookmarks.has(span.node) ||
          erasedCommentMarkers.has(span.node) ||
          (CANONICAL_SOURCE_INLINE_POLICIES[span.node.type] === "container" &&
            sourceContainerChildren(span.node).length === 0)
        )
      )
        break;
      index += 1;
      if (position !== null) addGap(position, span.after);
    }
  };
  addGap(0, { offset: 0, zeroWidthBefore: 0 });
  for (const cell of native.value) {
    consumeErasedLeaves(cell.from);
    if (cell.type === "zeroWidth" && cell.sources.length === 0) {
      // A control edge: both of its editor positions name the current gap.
      if (consumed !== 0)
        return refuse(
          "Canonical editing cannot place a content control edge inside text.",
        );
      const next = spans.at(index);
      const gap = next?.before ??
        spans.at(index - 1)?.after ?? { offset: 0, zeroWidthBefore: 0 };
      addGap(cell.from, gap);
      addGap(cell.to, gap);
      continue;
    }
    let span = spans.at(index);
    if (!span)
      return refuse(
        `Canonical editing cannot map native ${cell.node.type.name} to a source leaf.`,
      );
    addGap(
      cell.from,
      consumed === 0
        ? span.before
        : { offset: span.before.offset + consumed, zeroWidthBefore: 0 },
    );
    if (cell.type !== "zeroWidth" && explicitCommentIds.size > 0) {
      const coveredExplicitCommentIds = new Set<number>();
      for (const mark of cell.node.marks) {
        if (mark.type.name !== "comment") continue;
        const attrs = readCommentMarkAttrs(mark);
        if (!attrs.ok)
          return refuse(
            "Canonical editing requires valid comment mark attributes.",
          );
        if (explicitCommentIds.has(attrs.value.commentId))
          coveredExplicitCommentIds.add(attrs.value.commentId);
      }
      if (
        coveredExplicitCommentIds.size !== openExplicitCommentIds.size ||
        [...coveredExplicitCommentIds].some(
          (id) => !openExplicitCommentIds.has(id),
        )
      )
        return refuse(
          "Canonical comment marks do not match their explicit source boundaries.",
        );
    }
    switch (cell.type) {
      case "zeroWidth":
        for (const marker of cell.sources) {
          consumeErasedLeaves(null);
          span = spans.at(index);
          if (
            !span ||
            consumed !== 0 ||
            span.before.offset !== span.after.offset ||
            !sameSourceMarker(span.node, marker)
          )
            return refuse(
              `Canonical editing cannot preserve ${marker.type} source boundary positions.`,
            );
          if (
            marker.type === "commentRangeStart" &&
            explicitCommentIds.has(marker.id)
          )
            openExplicitCommentIds.add(marker.id);
          if (
            marker.type === "commentRangeEnd" &&
            explicitCommentIds.has(marker.id)
          )
            openExplicitCommentIds.delete(marker.id);
          index += 1;
        }
        if (!span)
          return refuse("Canonical editing lost a source range boundary.");
        addGap(cell.to, span.after);
        break;
      case "unit":
      case "markedUnit":
        if (
          cell.type === "markedUnit" &&
          ((span.node.type !== "footnoteRef" &&
            span.node.type !== "endnoteRef") ||
            span.node.type !== cell.sourceType ||
            String(span.node.id) !== cell.label)
        )
          return refuse(
            "Canonical editing cannot preserve note-reference identity.",
          );
        if (
          consumed !== 0 ||
          span.node.type === "text" ||
          span.after.offset - span.before.offset !== 1
        )
          return refuse(
            `Canonical editing cannot map ${span.node.type} to one native inline unit.`,
          );
        index += 1;
        addGap(cell.to, span.after);
        break;
      case "text":
        for (let unit = 0; unit < cell.text.length; unit += 1) {
          consumeErasedLeaves(cell.from + unit);
          span = spans.at(index);
          if (!span)
            return refuse(
              "Canonical editing cannot map missing source text to native text positions.",
            );
          let value: string;
          switch (span.node.type) {
            case "text":
              value = span.node.text;
              break;
            case "softHyphen":
            case "noBreakHyphen":
              value = HYPHEN_TEXT_CARRIERS[span.node.type];
              break;
            default:
              return refuse(
                `Canonical editing cannot map ${span.node.type} to native text positions.`,
              );
          }
          if (value.charAt(consumed) !== cell.text.charAt(unit))
            return refuse(
              "Canonical editing cannot preserve the source text carrier.",
            );
          consumed += 1;
          addGap(cell.from + unit + 1, {
            offset: span.before.offset + consumed,
            zeroWidthBefore: 0,
          });
          if (consumed === value.length) {
            index += 1;
            consumed = 0;
          }
        }
        break;
      default:
        cell satisfies never;
    }
  }
  consumeErasedLeaves(paragraph.content.size);
  if (
    [...comments.value.values()].some(
      ({ starts, ends }) => starts.size > 0 || ends.size > 0,
    )
  )
    return refuse(
      "Canonical editing cannot attribute a comment mark transition to its source.",
    );
  const missing = spans.at(index);
  if (missing || consumed !== 0)
    return refuse(
      `Canonical editing cannot display the source positions of ${missing?.node.type ?? "incomplete text"}.`,
    );
  if (boundaries.some((gaps) => gaps.length === 0))
    return refuse("Canonical editing lost a source character boundary.");
  return Result.ok({ text, boundaries, seams });
};

/** A membership whose answers one paragraph's projection read, recorded to check later. */
class RecordedMembership implements StoryIdMembership {
  readonly answers = new Map<number, boolean>();
  private readonly source: StoryIdMembership;
  constructor(source: StoryIdMembership) {
    this.source = source;
  }
  has(value: number): boolean {
    const answer = this.source.has(value);
    this.answers.set(value, answer);
    return answer;
  }
}

const answersHold = (
  answers: ReadonlyMap<number, boolean>,
  set: StoryIdMembership,
): boolean => {
  for (const [value, answer] of answers)
    if (set.has(value) !== answer) return false;
  return true;
};

const sameIds = (left: readonly number[], right: Iterable<number>) => {
  let index = 0;
  for (const value of right) {
    if (left[index] !== value) return false;
    index += 1;
  }
  return index === left.length;
};

/** Every story fact one paragraph's projection read. */
type StoryInlineReads = {
  pairedBookmarkIds: ReadonlyMap<number, boolean>;
  commentContext:
    | {
        rangedIds: ReadonlyMap<number, boolean>;
        pointIds: ReadonlyMap<number, boolean>;
        /** In the set's own order. */
        openIds: readonly number[];
      }
    | undefined;
};

/** The story inputs behind recorders, and what they read once the projection ran. */
const recordStory = (
  story: StoryInlineInputs,
): { recorded: StoryInlineInputs; reads: () => StoryInlineReads } => {
  const pairedBookmarkIds = new RecordedMembership(story.pairedBookmarkIds);
  const context = story.commentContext;
  if (context === undefined)
    return {
      recorded: { pairedBookmarkIds },
      reads: () => ({
        pairedBookmarkIds: pairedBookmarkIds.answers,
        commentContext: undefined,
      }),
    };
  const rangedIds = new RecordedMembership(context.rangedIds);
  const pointIds = new RecordedMembership(context.pointIds);
  const openIds = [...context.openIds];
  return {
    recorded: {
      pairedBookmarkIds,
      commentContext: { rangedIds, pointIds, openIds: new Set(openIds) },
    },
    reads: () => ({
      pairedBookmarkIds: pairedBookmarkIds.answers,
      commentContext: {
        rangedIds: rangedIds.answers,
        pointIds: pointIds.answers,
        openIds,
      },
    }),
  };
};

const storyReadsHold = (
  reads: StoryInlineReads,
  story: StoryInlineInputs,
): boolean => {
  if (!answersHold(reads.pairedBookmarkIds, story.pairedBookmarkIds))
    return false;
  const read = reads.commentContext;
  const context = story.commentContext;
  if (read === undefined || context === undefined)
    return read === undefined && context === undefined;
  return (
    sameIds(read.openIds, context.openIds) &&
    answersHold(read.rangedIds, context.rangedIds) &&
    answersHold(read.pointIds, context.pointIds)
  );
};

type CachedInlineProjection = {
  inputs: Pick<CanonicalInlineProjectionArgs, ParagraphInputField>;
  reads: StoryInlineReads;
  projection: InlineProjection;
};

/**
 * Paragraph projections kept across projections of a story.
 *
 * A paragraph an edit left keeps its source records, and the block cache
 * keeps its editor node, so its gaps can be reused where the story facts its
 * projection read are the same: the comments open at its start and the
 * comment and bookmark ids it asked about. Gap positions are relative to the
 * paragraph, so a reused projection holds wherever the paragraph now starts.
 */
export class InlineProjectionCache {
  private paragraphs = new WeakMap<Paragraph, CachedInlineProjection>();

  project(
    args: CanonicalInlineProjectionArgs,
  ): Result<InlineProjection, InlineProjectionError> {
    const known = this.paragraphs.get(args.source);
    if (
      known !== undefined &&
      PARAGRAPH_INPUT_FIELDS.every(
        (field) => known.inputs[field] === args[field],
      ) &&
      storyReadsHold(known.reads, args)
    )
      return Result.ok(known.projection);
    const { recorded, reads } = recordStory(args);
    const projected = projectCanonicalInline({ ...args, ...recorded });
    if (projected.isOk())
      this.paragraphs.set(args.source, {
        inputs: { source: args.source, paragraph: args.paragraph },
        reads: reads(),
        projection: projected.value,
      });
    return projected;
  }
}
