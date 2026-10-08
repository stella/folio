import { panic, Result, TaggedError } from "better-result";
import { inlineLeafSpans, paragraphLogicalText } from "@stll/docx-core/ops";
import { hasIllegalXmlCharacters } from "@stll/docx-core";
import type { Node as PMNode } from "prosemirror-model";
import type {
  Paragraph,
  ParagraphContent,
  RunContent,
  NoteReferenceContent,
} from "../types/document";
import {
  readBookmarkBoundaryAttrs,
  bookmarkMarkerFromAttrs,
} from "../prosemirror/bookmarkBoundaryAttrs";
import { readMoveRangeBoundaryAttrs } from "../prosemirror/moveRangeBoundaryAttrs";
import { readRangeAnchorAttrs } from "../prosemirror/rangeAnchorAttrs";
import { readNoteMarkerAttrs } from "../internal/noteMarkerAttrs";
import {
  readCommentMarkAttrs,
  readFootnoteRefMarkAttrs,
  readPreservedXmlAttrs,
} from "../prosemirror/attrs/index";
import { PRESERVED_XML_LEVELS } from "../prosemirror/schema/nodes";
import { TRACKED_RUN_INLINE_ATOM_DISPOSITIONS } from "../prosemirror/trackedRunInlineAtoms";
import { HYPHEN_TEXT_CARRIERS } from "../prosemirror/conversion/hyphenTextCarriers";

type SourceKind = ParagraphContent["type"] | RunContent["type"];
type SourcePolicy = "container" | "leaf" | "text" | "rawField" | "control" | "layout";
export const CANONICAL_SOURCE_INLINE_POLICIES = {
  run: "container",
  hyperlink: "container",
  insertion: "container",
  deletion: "container",
  moveFrom: "container",
  moveTo: "container",
  inlineWrapper: "container",
  inlineSdt: "control",
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
  renderedPageBreak: "layout",
  preservedXml: "leaf",
  drawing: "leaf",
  shape: "leaf",
} as const satisfies Record<SourceKind, SourcePolicy>;

type NativePolicy = "text" | "unit" | "markers" | "noteMarker" | "capture" | "refused";
export const CANONICAL_NATIVE_INLINE_POLICIES = {
  text: "text",
  bookmarkBoundary: "markers",
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
  renderedPageBreak: "refused",
  sdt: "refused",
} as const satisfies Record<
  keyof typeof TRACKED_RUN_INLINE_ATOM_DISPOSITIONS | "sdt",
  NativePolicy
>;
const nativePolicies = new Map<string, NativePolicy>(
  Object.entries(CANONICAL_NATIVE_INLINE_POLICIES),
);

class InlineProjectionError extends TaggedError("InlineProjectionError")<{ message: string }> {}
const refuse = (message: string) => Result.err(new InlineProjectionError({ message }));

/** Every source variant makes a decision before conversion can discard its positions. */
export const canonicalInlineSourceIssue = (
  items: readonly (ParagraphContent | RunContent)[],
  context: "paragraph" | "fieldPayload" = "paragraph",
): string | null => {
  for (const item of items) {
    const policy = CANONICAL_SOURCE_INLINE_POLICIES[item.type];
    switch (policy) {
      case "control":
        if (context === "fieldPayload" && item.type === "inlineSdt") {
          const issue = canonicalInlineSourceIssue(item.content, context);
          if (issue !== null) return issue;
          break;
        }
        return "Canonical editing cannot preserve inline content-control boundary positions.";
      case "rawField":
        if (context === "fieldPayload") break;
        return "Canonical editing cannot preserve unstructured field instructions; a structured field is required.";
      case "layout":
        return "Canonical editing cannot preserve cached rendered page-break positions.";
      case "text":
        if (
          item.type === "text" &&
          (hasIllegalXmlCharacters(item.text) || /[\t\r\n]/u.test(item.text))
        )
          return "Canonical text requires valid XML characters and explicit tab or break elements.";
        break;
      case "container": {
        if (
          (item.type === "insertion" || item.type === "deletion") &&
          item.resolutionJoins !== undefined &&
          item.content.some((child) => child.type === "hyperlink")
        )
          return "Canonical editing cannot preserve hyperlink wrapper review provenance.";
        const children = sourceContainerChildren(item);
        const issue = canonicalInlineSourceIssue(children, context);
        if (issue !== null) return issue;
        break;
      }
      case "leaf": {
        // Fields are one positional unit, but their cached payload still needs validation.
        let children: readonly (ParagraphContent | RunContent)[] = [];
        if (item.type === "simpleField") children = item.content;
        if (item.type === "complexField") children = [...item.fieldCode, ...item.fieldResult];
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
      return item.content;
    default:
      return panic(`Source policy incorrectly classified ${item.type} as a container.`);
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
  paragraph.forEach((node, from) => {
    if (issue !== undefined) return;
    const to = from + node.nodeSize;
    const policy = nativePolicies.get(node.type.name);
    switch (policy) {
      case "text": {
        const reference = node.marks.find(({ type }) => type.name === "footnoteRef");
        if (!reference) {
          cells.push({ type: "text", text: node.text ?? "", from, to, node });
          return;
        }
        const attrs = readFootnoteRefMarkAttrs(reference);
        if (!attrs.ok) {
          issue = "Canonical editing requires valid note-reference mark attributes.";
          return;
        }
        const label = String(attrs.value.id);
        const text = node.text ?? "";
        if (
          label.length === 0 ||
          text.length % label.length !== 0 ||
          text !== label.repeat(text.length / label.length)
        ) {
          issue = "Canonical editing cannot map the rendered note-reference labels.";
          return;
        }
        // Identical marks merge adjacent labels into one native text node.
        for (let offset = 0; offset < text.length; offset += label.length) {
          cells.push({
            type: "markedUnit",
            sourceType: attrs.value.noteType === "endnote" ? "endnoteRef" : "footnoteRef",
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
            issue = "Canonical editing requires valid bookmark boundary attributes.";
            return;
          }
          cells.push({
            type: "zeroWidth",
            sources: [bookmarkMarkerFromAttrs(attrs.value)],
            from,
            to,
            node,
          });
        } else if (node.type.name === "moveRangeBoundary") {
          const attrs = readMoveRangeBoundaryAttrs(node);
          if (!attrs.ok) {
            issue = "Canonical editing requires valid move-range boundary attributes.";
            return;
          }
          cells.push({ type: "zeroWidth", sources: [attrs.value], from, to, node });
        } else {
          const attrs = readRangeAnchorAttrs(node);
          if (!attrs.ok) {
            issue = "Canonical editing requires a valid collapsed range anchor.";
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
        if (attrs.value.level === PRESERVED_XML_LEVELS.inline && attrs.value.text === "")
          cells.push({
            type: "zeroWidth",
            sources: [{ type: "preservedInline", xml: attrs.value.xml, text: "" }],
            from,
            to,
            node,
          });
        else cells.push({ type: "unit", from, to, node });
        return;
      }
      case "refused":
        issue = `Canonical editing cannot preserve ${node.type.name} boundary positions.`;
        return;
      case undefined:
        issue = `Canonical editing has no position mapping for inline ${node.type.name}.`;
        return;
      default:
        policy satisfies never;
    }
  });
  return issue === undefined ? Result.ok(cells) : refuse(issue);
};

export type CanonicalInlineGap = { position: number; zeroWidthBefore: number };
type InlineProjection = { text: string; boundaries: CanonicalInlineGap[][] };
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
    case "preservedInline":
      return right.type === "preservedInline" && left.xml === right.xml && left.text === right.text;
    default:
      return false;
  }
};

type CanonicalInlineProjectionArgs = {
  source: Paragraph;
  paragraph: PMNode;
  pairedBookmarkIds: ReadonlySet<number>;
  commentContext?: {
    rangedIds: ReadonlySet<number>;
    pointIds: ReadonlySet<number>;
    openIds: ReadonlySet<number>;
  };
};

type CommentTransition = { starts: Set<number>; ends: Set<number> };

/** Marks carry nonempty ranges; explicit range anchors carry empty ranges. */
const commentMarkTransitions = ({
  source,
  paragraph,
  commentContext,
}: Pick<CanonicalInlineProjectionArgs, "source" | "paragraph" | "commentContext">) => {
  const spans = inlineLeafSpans(source.content);
  const rangedIds = new Set(commentContext?.rangedIds);
  const pointIds = new Set(commentContext?.pointIds);
  const finalIds = new Set(commentContext?.openIds);
  for (const { node } of spans) {
    if (node.type === "commentRangeStart" || node.type === "commentRangeEnd")
      rangedIds.add(node.id);
    if (node.type === "commentReference") pointIds.add(node.id);
    if (node.type === "commentRangeStart") finalIds.add(node.id);
    if (node.type === "commentRangeEnd") finalIds.delete(node.id);
  }
  const transitions = new Map<number, CommentTransition>();
  let previous = new Set(commentContext?.openIds);
  let issue: string | undefined;
  const record = (position: number, current: Set<number>) => {
    const starts = new Set([...current].filter((id) => !previous.has(id)));
    const ends = new Set([...previous].filter((id) => !current.has(id)));
    if (starts.size > 0 || ends.size > 0) transitions.set(position, { starts, ends });
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
      if (rangedIds.has(id)) current.add(id);
      else if (!pointIds.has(id)) {
        issue = "Canonical editing cannot attribute a comment mark to its source.";
        return;
      }
    }
    record(position, current);
  });
  // A range may continue into the next paragraph, whose marks verify this context.
  record(paragraph.content.size, finalIds);
  return issue === undefined ? Result.ok(transitions) : refuse(issue);
};

/** Bind rendered gaps to the operation owner's source leaf ordinals, including collapsed pairs. */
export const projectCanonicalInline = ({
  source,
  paragraph,
  pairedBookmarkIds,
  commentContext,
}: CanonicalInlineProjectionArgs): Result<InlineProjection, InlineProjectionError> => {
  const native = canonicalNativeCells(paragraph);
  if (native.isErr()) return native;
  const comments = commentMarkTransitions({
    source,
    paragraph,
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
        if (context === "paragraph" && !pairedBookmarkIds.has(item.id)) erasedBookmarks.add(item);
        continue;
      }
      if (CANONICAL_SOURCE_INLINE_POLICIES[item.type] !== "container") continue;
      collectErasedBookmarks(
        sourceContainerChildren(item),
        item.type === "run" || item.type === "inlineWrapper" ? context : "nested",
      );
    }
  };
  collectErasedBookmarks(source.content, "paragraph");
  const text = paragraphLogicalText(source);
  const boundaries: CanonicalInlineGap[][] = Array.from({ length: text.length + 1 }, () => []);
  let index = 0;
  let consumed = 0;
  const addGap = (position: number, gap: { offset: number; zeroWidthBefore: number }) => {
    const gaps = boundaries.at(gap.offset);
    if (!gaps) panic(`Source leaf offset ${gap.offset} exceeds paragraph ${source.paraId}.`);
    const existing = gaps.findIndex((entry) => entry.position === position);
    const entry = { position, zeroWidthBefore: gap.zeroWidthBefore };
    // Erased source leaves share a native seam: keep its right-affine source gap.
    if (existing < 0) gaps.push(entry);
    else gaps[existing] = entry;
  };
  const consumeErasedLeaves = (position: number | null) => {
    while (consumed === 0) {
      const span = spans.at(index);
      if (
        position !== null &&
        span &&
        (span.node.type === "commentRangeStart" || span.node.type === "commentRangeEnd")
      ) {
        const transition = comments.value.get(position);
        const ids = span.node.type === "commentRangeStart" ? transition?.starts : transition?.ends;
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
    let span = spans.at(index);
    if (!span)
      return refuse(`Canonical editing cannot map native ${cell.node.type.name} to a source leaf.`);
    addGap(
      cell.from,
      consumed === 0 ? span.before : { offset: span.before.offset + consumed, zeroWidthBefore: 0 },
    );
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
          index += 1;
        }
        if (!span) return refuse("Canonical editing lost a source range boundary.");
        addGap(cell.to, span.after);
        break;
      case "unit":
      case "markedUnit":
        if (
          cell.type === "markedUnit" &&
          ((span.node.type !== "footnoteRef" && span.node.type !== "endnoteRef") ||
            span.node.type !== cell.sourceType ||
            String(span.node.id) !== cell.label)
        )
          return refuse("Canonical editing cannot preserve note-reference identity.");
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
            return refuse("Canonical editing cannot preserve the source text carrier.");
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
  if ([...comments.value.values()].some(({ starts, ends }) => starts.size > 0 || ends.size > 0))
    return refuse("Canonical editing cannot attribute a comment mark transition to its source.");
  const missing = spans.at(index);
  if (missing || consumed !== 0)
    return refuse(
      `Canonical editing cannot display the source positions of ${missing?.node.type ?? "incomplete text"}.`,
    );
  if (boundaries.some((gaps) => gaps.length === 0))
    return refuse("Canonical editing lost a source character boundary.");
  return Result.ok({ text, boundaries });
};
