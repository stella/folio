import { panic } from "better-result";
import { Fragment, Slice, type Mark, type MarkType, type Node as PMNode } from "prosemirror-model";
import type { Transaction } from "prosemirror-state";
import { ReplaceStep, StepMap, type Mappable } from "prosemirror-transform";

import { recreateProseNodeWithParagraphPropertySource } from "../docx/paragraphPropertySource";
import { expectRunPropertyChangeMarkAttrs } from "../prosemirror/attrs";
import { textFormattingToMarks } from "../prosemirror/conversion/toProseDoc";
import { RUN_FORMATTING_MARK_NAMES } from "../prosemirror/runFormattingMarkNames";

export type HeadlessRevisionResolutionMode = "accept" | "reject";

type HeadlessDeleteRange = {
  from: number;
  to: number;
};

type HeadlessInlineContext = {
  mode: HeadlessRevisionResolutionMode;
  keepType: MarkType | undefined;
  removeType: MarkType | undefined;
  deleteRanges: HeadlessDeleteRange[];
  changedParagraphRanges: HeadlessDeleteRange[];
};

export type HeadlessInlineChangeTracking = {
  ranges: readonly HeadlessDeleteRange[];
  mappingFrom: number;
};

/**
 * A synchronous headless replacement with the granular position map of the
 * inline deletions it applies. It must never enter history or collaboration;
 * map and serialization fail loudly if that invariant is broken.
 */
class HeadlessInlineResolutionStep extends ReplaceStep {
  private readonly positionMap: StepMap;

  constructor(from: number, to: number, slice: Slice, positionMap: StepMap) {
    super(from, to, slice);
    this.positionMap = positionMap;
  }

  override getMap(): StepMap {
    return this.positionMap;
  }

  override invert(doc: PMNode): ReplaceStep {
    return new HeadlessInlineResolutionStep(
      this.from,
      this.from + this.slice.size,
      doc.slice(this.from, this.to),
      this.positionMap.invert(),
    );
  }

  override map(_mapping: Mappable): never {
    return panic("A headless revision-resolution step cannot be mapped");
  }

  override toJSON(): never {
    return panic("A headless revision-resolution step cannot be serialized");
  }
}

const marksEqual = (left: readonly Mark[], right: readonly Mark[]): boolean =>
  left.length === right.length &&
  left.every((mark, index) => {
    const candidate = right[index];
    return candidate !== undefined && mark.eq(candidate);
  });

const rebuildNode = (
  source: PMNode,
  attrs: PMNode["attrs"],
  content: Fragment,
  marks: readonly Mark[] = source.marks,
): PMNode =>
  recreateProseNodeWithParagraphPropertySource(source, {
    attrs,
    content,
    marks,
  });

const resolveInlineNode = (
  node: PMNode,
  mode: HeadlessRevisionResolutionMode,
  keepType: MarkType | undefined,
  removeType: MarkType | undefined,
): PMNode | null => {
  let marks: readonly Mark[] = node.marks;
  const runPropertyChangeMark = marks.find((mark) => mark.type.name === "runPropertyChange");
  if (runPropertyChangeMark) {
    const { changes } = expectRunPropertyChangeMarkAttrs(runPropertyChangeMark);
    if (changes.length > 0) {
      marks = marks.filter((mark) => mark !== runPropertyChangeMark);
      if (mode === "reject") {
        marks = marks.filter((mark) => !RUN_FORMATTING_MARK_NAMES.has(mark.type.name));
        const previousFormatting = changes.at(0)?.previousFormatting;
        for (const previousMark of textFormattingToMarks(previousFormatting)) {
          marks = previousMark.addToSet(marks);
        }
        if (previousFormatting?.styleId) {
          const characterStyle = node.type.schema.marks["characterStyle"];
          if (characterStyle) {
            marks = characterStyle
              .create({ styleId: previousFormatting.styleId, _styleRPr: null })
              .addToSet(marks);
          }
        }
      }
    }
  }

  if (removeType && node.marks.some((mark) => mark.type === removeType)) {
    return null;
  }
  if (keepType) {
    marks = marks.filter((mark) => mark.type !== keepType);
  }
  return marksEqual(marks, node.marks) ? node : node.mark(marks);
};

const resolveInlineContent = (
  node: PMNode,
  position: number,
  context: HeadlessInlineContext,
): PMNode | null => {
  let resolvedNode = node;
  if (node.isInline) {
    const resolved = resolveInlineNode(node, context.mode, context.keepType, context.removeType);
    if (resolved === null) {
      context.deleteRanges.push({ from: position, to: position + node.nodeSize });
      return null;
    }
    resolvedNode = resolved;
    if (resolvedNode.isLeaf) {
      return resolvedNode;
    }
  }

  const children: PMNode[] = [];
  const contentStart = resolvedNode.type.name === "doc" ? 0 : position + 1;
  let contentChanged = false;
  resolvedNode.forEach((child, offset) => {
    const resolved = resolveInlineContent(child, contentStart + offset, context);
    if (resolved) {
      children.push(resolved);
      contentChanged ||= resolved !== child;
    } else {
      contentChanged = true;
    }
  });

  if (!contentChanged) {
    return resolvedNode;
  }
  if (resolvedNode.type.name === "paragraph") {
    context.changedParagraphRanges.push({ from: position, to: position + resolvedNode.nodeSize });
  }
  return rebuildNode(node, resolvedNode.attrs, Fragment.fromArray(children), resolvedNode.marks);
};

const coalesceDeleteRanges = (
  deleteRanges: readonly HeadlessDeleteRange[],
): HeadlessDeleteRange[] => {
  const coalesced: HeadlessDeleteRange[] = [];
  for (const range of deleteRanges) {
    const previous = coalesced.at(-1);
    if (previous && range.from <= previous.to) {
      previous.to = Math.max(previous.to, range.to);
    } else {
      coalesced.push({ ...range });
    }
  }
  return coalesced;
};

export const appendHeadlessInlineResolution = (
  tr: Transaction,
  mode: HeadlessRevisionResolutionMode,
  keepType: MarkType | undefined,
  removeType: MarkType | undefined,
): HeadlessInlineChangeTracking | null => {
  const context: HeadlessInlineContext = {
    mode,
    keepType,
    removeType,
    deleteRanges: [],
    changedParagraphRanges: [],
  };
  const resolved = resolveInlineContent(tr.doc, -1, context);
  if (!resolved || resolved.eq(tr.doc)) {
    return null;
  }
  const deleteRanges = coalesceDeleteRanges(context.deleteRanges);
  const positionMap = new StepMap(deleteRanges.flatMap(({ from, to }) => [from, to - from, 0]));
  const mappingFrom = tr.steps.length;
  tr.step(
    new HeadlessInlineResolutionStep(
      0,
      tr.doc.content.size,
      new Slice(resolved.content, 0, 0),
      positionMap,
    ),
  );
  return { ranges: context.changedParagraphRanges, mappingFrom };
};
