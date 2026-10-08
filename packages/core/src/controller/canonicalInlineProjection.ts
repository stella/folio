import { panic, Result, TaggedError } from "better-result";
import { inlineLeafSpans, paragraphLogicalText } from "@stll/docx-core/ops";
import { hasIllegalXmlCharacters } from "@stll/docx-core";
import type { Node as PMNode } from "prosemirror-model";
import type { Paragraph, ParagraphContent, RunContent } from "../types/document";
import {
  readBookmarkBoundaryAttrs,
  bookmarkMarkerFromAttrs,
} from "../prosemirror/bookmarkBoundaryAttrs";
import { readMoveRangeBoundaryAttrs } from "../prosemirror/moveRangeBoundaryAttrs";
import { readRangeAnchorAttrs } from "../prosemirror/rangeAnchorAttrs";
import { readNoteMarkerAttrs } from "../internal/noteMarkerAttrs";
import { readPreservedXmlAttrs } from "../prosemirror/attrs/index";
import { PRESERVED_XML_LEVELS } from "../prosemirror/schema/nodes";
import { TRACKED_RUN_INLINE_ATOM_DISPOSITIONS } from "../prosemirror/trackedRunInlineAtoms";

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
      case "text":
        if (node.marks.some(({ type }) => type.name === "footnoteRef"))
          cells.push({ type: "unit", from, to, node });
        else cells.push({ type: "text", text: node.text ?? "", from, to, node });
        return;
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

/** Bind rendered gaps to the operation owner's source leaf ordinals, including collapsed pairs. */
export const projectCanonicalInline = (
  source: Paragraph,
  paragraph: PMNode,
): Result<InlineProjection, InlineProjectionError> => {
  const native = canonicalNativeCells(paragraph);
  if (native.isErr()) return native;
  const spans = inlineLeafSpans(source.content);
  const text = paragraphLogicalText(source);
  const boundaries: CanonicalInlineGap[][] = Array.from({ length: text.length + 1 }, () => []);
  let index = 0;
  let consumed = 0;
  const addGap = (position: number, gap: { offset: number; zeroWidthBefore: number }) => {
    const gaps = boundaries.at(gap.offset);
    if (!gaps) panic(`Source leaf offset ${gap.offset} exceeds paragraph ${source.paraId}.`);
    if (
      !gaps.some(
        (entry) => entry.position === position && entry.zeroWidthBefore === gap.zeroWidthBefore,
      )
    )
      gaps.push({ position, zeroWidthBefore: gap.zeroWidthBefore });
  };
  addGap(0, { offset: 0, zeroWidthBefore: 0 });
  for (const cell of native.value) {
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
          span = spans.at(index);
          if (
            !span ||
            span.node.type !== "text" ||
            span.node.text.charAt(consumed) !== cell.text.charAt(unit)
          )
            return refuse(
              `Canonical editing cannot map ${span?.node.type ?? "missing source text"} to native text positions.`,
            );
          consumed += 1;
          addGap(cell.from + unit + 1, {
            offset: span.before.offset + consumed,
            zeroWidthBefore: 0,
          });
          if (consumed === span.node.text.length) {
            index += 1;
            consumed = 0;
          }
        }
        break;
      default:
        cell satisfies never;
    }
  }
  const missing = spans.at(index);
  if (missing || consumed !== 0)
    return refuse(
      `Canonical editing cannot display the source positions of ${missing?.node.type ?? "incomplete text"}.`,
    );
  if (boundaries.some((gaps) => gaps.length === 0))
    return refuse("Canonical editing lost a source character boundary.");
  return Result.ok({ text, boundaries });
};
