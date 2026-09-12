import { panic } from "better-result";
import { Fragment, Slice, type Mark, type MarkType, type Node as PMNode } from "prosemirror-model";
import type { Transaction } from "prosemirror-state";
import { ReplaceStep, StepMap, type Mappable } from "prosemirror-transform";

import { recreateProseNodeWithParagraphPropertySource } from "../docx/paragraphPropertySource";
import { expectRunPropertyChangeMarkAttrs } from "../prosemirror/attrs";
import { reconstructResolvedRunFormattingMarks } from "../prosemirror/runPropertyChangeResolution";
import { RUN_FORMATTING_MARK_NAMES } from "../prosemirror/runFormattingMarkNames";
import {
  paragraphRunStyleContext,
  type ParagraphRunStyleContext,
  type RunStyleResolver,
} from "../prosemirror/runStyleFormatting";

export type HeadlessRevisionResolutionMode = "accept" | "reject";

type HeadlessRange = {
  from: number;
  to: number;
};

type HeadlessReplacementRange = HeadlessRange & {
  newSize: number;
};

type HeadlessInlineContext = {
  mode: HeadlessRevisionResolutionMode;
  keepType: MarkType | undefined;
  removeType: MarkType | undefined;
  replacementRanges: HeadlessReplacementRange[];
  changedParagraphRanges: HeadlessRange[];
  styleResolver: RunStyleResolver | null;
};

export type HeadlessInlineChangeTracking = {
  ranges: readonly HeadlessRange[];
  mappingFrom: number;
};

const EMPTY_PARAGRAPH_RUN_STYLE_CONTEXT: ParagraphRunStyleContext = {
  baseParagraphFormatting: undefined,
  paragraphFormatting: undefined,
  paragraphMarkFormatting: undefined,
  paragraphMarkPrecedesStyle: false,
};

type ParagraphRunStyleScope = {
  paragraph: PMNode;
  resolved?: ParagraphRunStyleContext;
};

const resolveParagraphRunStyleScope = (
  scope: ParagraphRunStyleScope | undefined,
  styleResolver: RunStyleResolver | null,
): ParagraphRunStyleContext => {
  if (!scope) {
    return EMPTY_PARAGRAPH_RUN_STYLE_CONTEXT;
  }
  scope.resolved ??= paragraphRunStyleContext(scope.paragraph, styleResolver);
  return scope.resolved;
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

type ResolveInlineNodeOptions = {
  node: PMNode;
  context: HeadlessInlineContext;
  paragraphScope: ParagraphRunStyleScope | undefined;
};

const resolveInlineNode = ({
  node,
  context,
  paragraphScope,
}: ResolveInlineNodeOptions): PMNode | null => {
  let marks: readonly Mark[] = node.marks;
  const runPropertyChangeMark = marks.find((mark) => mark.type.name === "runPropertyChange");
  if (runPropertyChangeMark) {
    const { changes } = expectRunPropertyChangeMarkAttrs(runPropertyChangeMark);
    if (changes.length > 0) {
      marks = marks.filter((mark) => mark !== runPropertyChangeMark);
      marks = marks.filter((mark) => !RUN_FORMATTING_MARK_NAMES.has(mark.type.name));
      const previousFormatting = changes.at(0)?.previousFormatting;
      for (const resolvedMark of reconstructResolvedRunFormattingMarks({
        node,
        paragraphContext: resolveParagraphRunStyleScope(paragraphScope, context.styleResolver),
        previousFormatting,
        mode: context.mode,
        styleResolver: context.styleResolver,
      })) {
        marks = resolvedMark.addToSet(marks);
      }
    }
  }

  if (context.removeType && node.marks.some((mark) => mark.type === context.removeType)) {
    return null;
  }
  if (context.keepType) {
    marks = marks.filter((mark) => mark.type !== context.keepType);
  }
  return marksEqual(marks, node.marks) ? node : node.mark(marks);
};

type ResolveInlineContentOptions = {
  node: PMNode;
  position: number;
  context: HeadlessInlineContext;
  inheritedParagraphScope?: ParagraphRunStyleScope;
};

const resolveInlineContent = ({
  node,
  position,
  context,
  inheritedParagraphScope,
}: ResolveInlineContentOptions): PMNode | null => {
  const paragraphScope =
    node.type.name === "paragraph" ? { paragraph: node } : inheritedParagraphScope;
  const replacementRangeStart = context.replacementRanges.length;
  let resolvedNode = node;
  if (node.isInline) {
    const resolved = resolveInlineNode({
      node,
      context,
      paragraphScope,
    });
    if (resolved === null) {
      context.replacementRanges.push({
        from: position,
        to: position + node.nodeSize,
        newSize: 0,
      });
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
    const resolved = resolveInlineContent({
      node: child,
      position: contentStart + offset,
      context,
      ...(paragraphScope !== undefined ? { inheritedParagraphScope: paragraphScope } : {}),
    });
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
  let resolvedContent = Fragment.fromArray(children);
  if (!resolvedNode.type.validContent(resolvedContent)) {
    // A required-content parent may need a generated child after resolution
    // removes its last carrier. Let the schema choose that filler instead of
    // constructing an invalid node.
    const fitted = resolvedNode.type.createAndFill(
      resolvedNode.attrs,
      resolvedContent,
      resolvedNode.marks,
    );
    if (!fitted || !resolvedNode.type.validContent(fitted.content)) {
      if (!resolvedNode.isInline) {
        return panic(`Headless inline resolution invalidated ${resolvedNode.type.name} content`);
      }
      context.replacementRanges.splice(
        replacementRangeStart,
        context.replacementRanges.length - replacementRangeStart,
        { from: position, to: position + node.nodeSize, newSize: 0 },
      );
      return null;
    }
    // The fitted content is larger than the deletion-only reconstruction.
    // Replace its nested maps with the actual content-size transition so a
    // selection after the parent moves by the same distance as the slice.
    context.replacementRanges.splice(
      replacementRangeStart,
      context.replacementRanges.length - replacementRangeStart,
      {
        from: contentStart,
        to: contentStart + node.content.size,
        newSize: fitted.content.size,
      },
    );
    resolvedContent = fitted.content;
  }
  if (resolvedNode.type.name === "paragraph") {
    context.changedParagraphRanges.push({ from: position, to: position + resolvedNode.nodeSize });
  }
  // Rebuild through the provenance-aware owner so a filled paragraph keeps
  // its captured property source.
  return rebuildNode(node, resolvedNode.attrs, resolvedContent, resolvedNode.marks);
};

const coalesceReplacementRanges = (
  replacementRanges: readonly HeadlessReplacementRange[],
): HeadlessReplacementRange[] => {
  const coalesced: HeadlessReplacementRange[] = [];
  for (const range of replacementRanges) {
    const previous = coalesced.at(-1);
    if (previous && previous.newSize === 0 && range.newSize === 0 && range.from <= previous.to) {
      previous.to = Math.max(previous.to, range.to);
    } else {
      coalesced.push({ ...range });
    }
  }
  return coalesced;
};

type AppendHeadlessInlineResolutionOptions = {
  tr: Transaction;
  mode: HeadlessRevisionResolutionMode;
  keepType: MarkType | undefined;
  removeType: MarkType | undefined;
  styleResolver: RunStyleResolver | null;
};

export const appendHeadlessInlineResolution = ({
  tr,
  mode,
  keepType,
  removeType,
  styleResolver,
}: AppendHeadlessInlineResolutionOptions): HeadlessInlineChangeTracking | null => {
  const context: HeadlessInlineContext = {
    mode,
    keepType,
    removeType,
    replacementRanges: [],
    changedParagraphRanges: [],
    styleResolver,
  };
  const resolved = resolveInlineContent({ node: tr.doc, position: -1, context });
  if (!resolved || resolved.eq(tr.doc)) {
    return null;
  }
  const replacementRanges = coalesceReplacementRanges(context.replacementRanges);
  const positionMap = new StepMap(
    replacementRanges.flatMap(({ from, to, newSize }) => [from, to - from, newSize]),
  );
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
