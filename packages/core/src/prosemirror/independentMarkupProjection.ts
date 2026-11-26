import { Fragment, type Node as PMNode } from "prosemirror-model";
import { StepMap } from "prosemirror-transform";

import {
  type ResolvedDocumentView,
  resolvedDocumentView,
  type RevisionDecision,
} from "@stll/docx-core/ops";

import type { Paragraph, ParagraphContent, Run } from "../types/document";
import { expectTrackedChangeMarkAttrs } from "./attrs";
import { INLINE_WRAPPER_MARK_NAME } from "./extensions/marks/InlineWrapperExtension";
import { getFolioNodeRevisionCarriers } from "./revisionCarriers";

type InlineChanges = NonNullable<ResolvedDocumentView["independentInline"]>;
type PositionRange = { from: number; to: number };

export type IndependentMarkupProjection = {
  doc: PMNode;
  positionMap: StepMap;
  changeBarRanges: readonly PositionRange[];
};

const isInlineRevision = (name: string): boolean =>
  name === "insertion" || name === "deletion";

const projectedText = (
  paragraph: PMNode,
  { omitted }: InlineChanges[number],
): Fragment | undefined => {
  const nodes: PMNode[] = [];
  let cutIndex = 0;
  for (let index = 0, start = 0; index < paragraph.childCount; index += 1) {
    const node = paragraph.child(index);
    if (!node.isText) return undefined;
    const end = start + node.nodeSize;
    const marks = node.marks.filter(({ type }) => !isInlineRevision(type.name));
    let at = start;
    for (let cut = omitted[cutIndex]; cut !== undefined && cut.from < end;) {
      if (cut.from > at)
        nodes.push(node.cut(at - start, cut.from - start).mark(marks));
      at = Math.max(at, cut.to);
      if (cut.to > end) break;
      cutIndex += 1;
      cut = omitted[cutIndex];
    }
    if (at < end) nodes.push(node.cut(at - start).mark(marks));
    start = end;
  }
  return Fragment.fromArray(nodes);
};

/**
 * The native resolver supplies exactly which text ranges disappeared. Keep the
 * original formatting and addresses of the rest, without converting standing
 * blocks or reconstructing their run styles.
 */
export const projectIndependentMarkup = ({
  doc,
  changes,
  indices,
}: {
  doc: PMNode;
  changes: InlineChanges;
  /** For a detached source, where the inspected paragraphs originally stood. */
  indices?: readonly number[];
}): IndependentMarkupProjection | undefined => {
  if (changes.length === 0)
    return { doc, positionMap: StepMap.empty, changeBarRanges: [] };
  const children: PMNode[] = [];
  const ranges: number[] = [];
  const changeBarRanges: PositionRange[] = [];
  let previous = 0;
  let changeIndex = 0;
  let removed = 0;
  let last = -1;
  for (let index = 0, position = 0; index < doc.childCount; index += 1) {
    const original = doc.child(index);
    const change = changes[changeIndex];
    const target =
      change === undefined
        ? undefined
        : (indices?.[change.index] ?? change.index);
    if (change === undefined || target !== index) {
      children.push(original);
      position += original.nodeSize;
      continue;
    }
    if (
      original.type.name !== "paragraph" ||
      original.content.size !== change.sourceLength
    )
      return undefined;
    const content = projectedText(original, change);
    if (content === undefined) return undefined;
    const paragraph = original.copy(content);
    const from = position - removed;
    changeBarRanges.push({ from, to: from + paragraph.nodeSize });
    children.push(paragraph);
    for (const { from: first, to } of change.omitted) {
      const start = position + 1 + first;
      const width = to - first;
      if (start === last)
        ranges[previous + 1] = (ranges[previous + 1] ?? 0) + width;
      else {
        previous = ranges.length;
        ranges.push(start, width, 0);
      }
      last = start + width;
      removed += width;
    }
    position += original.nodeSize;
    changeIndex += 1;
  }
  if (changeIndex !== changes.length) return undefined;
  return {
    doc: doc.copy(Fragment.fromArray(children)),
    positionMap: new StepMap(ranges),
    changeBarRanges,
  };
};

/** Project a detached text story after canonical resolution of its text carriers. */
export const projectDetachedText = (
  doc: PMNode,
  decision: RevisionDecision,
): IndependentMarkupProjection | undefined => {
  const changed: Paragraph[] = [];
  const indices: number[] = [];
  for (let index = 0, position = 0; index < doc.childCount; index += 1) {
    const paragraph = doc.child(index);
    if (
      paragraph.type.name !== "paragraph" ||
      paragraph.attrs["_suggestedInsert"] != null ||
      getFolioNodeRevisionCarriers(paragraph, position).length > 0
    )
      return undefined;
    let revised = false;
    const content: ParagraphContent[] = [];
    for (
      let childIndex = 0;
      childIndex < paragraph.childCount;
      childIndex += 1
    ) {
      const child = paragraph.child(childIndex);
      if (!child.isText || child.text === undefined) return undefined;
      let wrapper:
        | Extract<ParagraphContent, { type: "insertion" | "deletion" }>
        | undefined;
      const run: Run = {
        type: "run",
        content: [{ type: "text", text: child.text }],
      };
      for (const mark of child.marks) {
        const type = mark.type.name;
        if (
          type === "comment" ||
          type === "hyperlink" ||
          type === "runPropertyChange" ||
          type === INLINE_WRAPPER_MARK_NAME
        )
          return undefined;
        if (type !== "insertion" && type !== "deletion") continue;
        const attrs = expectTrackedChangeMarkAttrs(mark);
        if (
          wrapper !== undefined ||
          attrs.moveKind != null ||
          attrs._docxResolutionJoins != null ||
          (attrs._docxRevisionAncestors?.length ?? 0) > 0
        )
          return undefined;
        wrapper = {
          type,
          info: { id: attrs.revisionId, author: attrs.author },
          content: [run],
        };
      }
      revised ||= wrapper !== undefined;
      content.push(wrapper ?? run);
    }
    if (revised) {
      changed.push({ type: "paragraph", content });
      indices.push(index);
    }
    position += paragraph.nodeSize;
  }
  if (changed.length === 0)
    return { doc, positionMap: StepMap.empty, changeBarRanges: [] };
  const resolved = resolvedDocumentView(
    { package: { document: { content: changed } } },
    decision,
  );
  return resolved.independentInline === undefined
    ? undefined
    : projectIndependentMarkup({
        doc,
        changes: resolved.independentInline,
        indices,
      });
};
