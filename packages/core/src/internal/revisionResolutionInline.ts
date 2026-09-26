import { panic } from "better-result";
import { Fragment, Slice, type Mark, type MarkType, type Node as PMNode } from "prosemirror-model";
import { ReplaceStep, StepMap, type Step } from "prosemirror-transform";

import { recordNodeResolution } from "./revisionResolutionEdits";
import { expectTrackedChangeMarkAttrs } from "../prosemirror/attrs";

import { recreateProseNodeWithParagraphPropertySource } from "../docx/paragraphPropertySource";
import { expectRunPropertyChangeMarkAttrs } from "../prosemirror/attrs";
import {
  resolutionRemovesControl,
  withoutResolvedEnclosures,
} from "../prosemirror/contentControlRevisions";
import { INLINE_CONTENT_CONTROL_NODE_NAME } from "../prosemirror/extensions/nodes/SdtExtension";
import { continuedRunMarks } from "../prosemirror/rejoinRunCarriers";
import { reconstructRejectedRunFormattingMarks } from "../prosemirror/runPropertyChangeResolution";
import { RUN_FORMATTING_MARK_NAMES } from "../prosemirror/runFormattingMarkNames";
import {
  paragraphRunStyleContext,
  type ParagraphRunStyleContext,
  type RunStyleResolver,
} from "../prosemirror/runStyleFormatting";

export type RevisionResolutionMode = "accept" | "reject";

type RevisionRange = {
  from: number;
  to: number;
};

type RevisionReplacementRange = RevisionRange & {
  newSize: number;
  slice?: Slice;
};

type RevisionInlineContext = {
  mode: RevisionResolutionMode;
  keepType: MarkType | undefined;
  removeType: MarkType | undefined;
  replacementRanges: RevisionReplacementRange[];
  steps: Step[];
  changedParagraphRanges: RevisionRange[];
  styleResolver: RunStyleResolver | null;
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
  context: RevisionInlineContext;
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
      if (context.mode === "reject") {
        marks = marks.filter((mark) => !RUN_FORMATTING_MARK_NAMES.has(mark.type.name));
        const previousFormatting = changes.at(0)?.previousFormatting;
        for (const previousMark of reconstructRejectedRunFormattingMarks({
          node,
          paragraphContext: resolveParagraphRunStyleScope(paragraphScope, context.styleResolver),
          previousFormatting,
          styleResolver: context.styleResolver,
        })) {
          marks = previousMark.addToSet(marks);
        }
      }
    }
  }

  const nestedRemoval = node.marks.some(
    (mark) =>
      (mark.type.name === "insertion" || mark.type.name === "deletion") &&
      expectTrackedChangeMarkAttrs(mark)._docxRevisionAncestors?.some((layer) =>
        context.mode === "accept"
          ? layer.type === "deletion" || layer.type === "moveFrom"
          : layer.type === "insertion" || layer.type === "moveTo",
      ),
  );
  if (
    nestedRemoval ||
    (context.removeType && node.marks.some((mark) => mark.type === context.removeType))
  ) {
    return null;
  }
  if (context.keepType) {
    marks = marks.filter((mark) => mark.type !== context.keepType);
  }
  return marksEqual(marks, node.marks) ? node : node.mark(marks);
};

type RejoinResolvedRunsOptions = {
  children: PMNode[];
  positions: number[];
  resolvedBoundaries: ReadonlySet<number>;
  paragraphScope: ParagraphRunStyleScope | undefined;
  context: RevisionInlineContext;
};

/**
 * The pieces a revision split off its run, one run again: the tree-walk counterpart of
 * `rejoinRunsAt`, over the children of one inline container.
 */
const rejoinResolvedRuns = ({
  children,
  positions,
  resolvedBoundaries,
  paragraphScope,
  context,
}: RejoinResolvedRunsOptions): void => {
  for (const index of resolvedBoundaries) {
    const left = children[index - 1];
    const right = children[index];
    if (!left || !right) {
      continue;
    }
    const marks = continuedRunMarks({
      left,
      right,
      context: () => resolveParagraphRunStyleScope(paragraphScope, context.styleResolver),
      styleResolver: context.styleResolver,
    });
    if (marks !== null) {
      const joined = right.mark(marks);
      recordNodeResolution({
        before: right,
        after: joined,
        position: positions[index] ?? 0,
        steps: context.steps,
      });
      children[index] = joined;
    }
  }
};

type ResolveInlineContentOptions = {
  node: PMNode;
  position: number;
  context: RevisionInlineContext;
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
    recordNodeResolution({ before: node, after: resolvedNode, position, steps: context.steps });
    if (resolvedNode.isLeaf) {
      return resolvedNode;
    }
  }

  // A revision that encloses a control takes the control with it; one over
  // its content leaves it standing, emptied. The editor-command path decides
  // the same thing through the same predicate.
  if (
    resolutionRemovesControl({
      control: resolvedNode,
      removeType: context.removeType,
      resolves: () => true,
    })
  ) {
    context.replacementRanges.push({ from: position, to: position + node.nodeSize, newSize: 0 });
    return null;
  }

  const children: PMNode[] = [];
  const positions: number[] = [];
  /** Indices into `children` where resolved content began or ended. */
  const resolvedBoundaries = new Set<number>();
  const contentStart = resolvedNode.type.name === "doc" ? 0 : position + 1;
  let contentChanged = false;
  resolvedNode.forEach((child, offset) => {
    const resolved = resolveInlineContent({
      node: child,
      position: contentStart + offset,
      context,
      ...(paragraphScope !== undefined ? { inheritedParagraphScope: paragraphScope } : {}),
    });
    if (resolved !== child) {
      contentChanged = true;
      resolvedBoundaries.add(children.length);
    }
    if (resolved) {
      children.push(resolved);
      positions.push(contentStart + offset);
      if (resolved !== child) {
        resolvedBoundaries.add(children.length);
      }
    }
  });

  // Every revision resolves here, so none a control names encloses it any more.
  const resolvedAttrs =
    resolvedNode.type.name === INLINE_CONTENT_CONTROL_NODE_NAME
      ? (withoutResolvedEnclosures(resolvedNode, () => true) ?? resolvedNode.attrs)
      : resolvedNode.attrs;
  if (resolvedAttrs !== resolvedNode.attrs)
    recordNodeResolution({
      before: resolvedNode,
      after: rebuildNode(resolvedNode, resolvedAttrs, resolvedNode.content),
      position,
      steps: context.steps,
    });
  if (!contentChanged && resolvedAttrs === resolvedNode.attrs) {
    return resolvedNode;
  }
  if (contentChanged && resolvedNode.inlineContent) {
    rejoinResolvedRuns({ children, positions, resolvedBoundaries, paragraphScope, context });
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
        return panic(`Revision inline resolution invalidated ${resolvedNode.type.name} content`);
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
        slice: new Slice(fitted.content, 0, 0),
      },
    );
    resolvedContent = fitted.content;
  }
  if (resolvedNode.type.name === "paragraph") {
    context.changedParagraphRanges.push({ from: position, to: position + resolvedNode.nodeSize });
  }
  // Rebuild through the provenance-aware owner so a filled paragraph keeps
  // its captured property source.
  return rebuildNode(node, resolvedAttrs, resolvedContent, resolvedNode.marks);
};

const coalesceReplacementRanges = (
  replacementRanges: readonly RevisionReplacementRange[],
): RevisionReplacementRange[] => {
  const coalesced: RevisionReplacementRange[] = [];
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

type ResolveInlineRevisionsOptions = {
  doc: PMNode;
  mode: RevisionResolutionMode;
  keepType: MarkType | undefined;
  removeType: MarkType | undefined;
  styleResolver: RunStyleResolver | null;
};

export const resolveInlineRevisions = ({
  doc,
  mode,
  keepType,
  removeType,
  styleResolver,
}: ResolveInlineRevisionsOptions) => {
  const context: RevisionInlineContext = {
    mode,
    keepType,
    removeType,
    replacementRanges: [],
    steps: [],
    changedParagraphRanges: [],
    styleResolver,
  };
  const resolved = resolveInlineContent({ node: doc, position: -1, context });
  if (!resolved || resolved.eq(doc)) {
    return null;
  }
  const replacementRanges = coalesceReplacementRanges(context.replacementRanges);
  const positionMap = new StepMap(
    replacementRanges.flatMap(({ from, to, newSize }) => [from, to - from, newSize]),
  );
  for (const { from, to, slice } of replacementRanges.toReversed()) {
    context.steps.push(new ReplaceStep(from, to, slice ?? Slice.empty));
  }
  return { resolved, positionMap, ranges: context.changedParagraphRanges, steps: context.steps };
};
