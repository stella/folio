/**
 * Splits a paragraph at its page-break runs into paragraph fragments and
 * page-break blocks.
 */

import type { Node as PMNode } from "prosemirror-model";
import { panic } from "better-result";
import { partitionRunsAtPageBreaks } from "../../internal/pageBreakRunPartition";
import type {
  ParagraphBlock,
  PageBreakBlock,
  Run,
  ParagraphAttrs,
} from "../../layout-engine/types";
import { nextBlockId } from "./flowConversionShared";
import type { PageBreakRunProjection } from "./paragraphRuns";
import { hasVisibleParagraphPayload } from "./paragraphConversion";

function paragraphFragmentAttrs(
  source: ParagraphAttrs | undefined,
  index: number,
  count: number,
): ParagraphAttrs | undefined {
  if (!source) {
    return undefined;
  }
  const attrs: ParagraphAttrs = {
    ...source,
    ...(source.indent ? { indent: { ...source.indent } } : {}),
    ...(source.spacing ? { spacing: { ...source.spacing } } : {}),
    ...(source.automaticSpacing ? { automaticSpacing: { ...source.automaticSpacing } } : {}),
    ...(source.spacingExplicit ? { spacingExplicit: { ...source.spacingExplicit } } : {}),
  };
  if (index > 0) {
    delete attrs.listMarker;
    delete attrs.listIsBullet;
    delete attrs.listMarkerFormatting;
    delete attrs.listMarkerHidden;
    delete attrs.listMarkerAlignment;
    delete attrs.listMarkerSuffix;
    delete attrs.listMarkerRevision;
    delete attrs.listMarkerSecondSlotOffsetTwips;
    delete attrs.reserveEmptyOutlineHeight;
    delete attrs.pageBreakBefore;
    delete attrs.renderedPageBreakBefore;
    if (attrs.indent) {
      delete attrs.indent.firstLine;
      delete attrs.indent.hanging;
    }
    if (attrs.spacing) {
      delete attrs.spacing.before;
    }
    if (attrs.automaticSpacing) {
      delete attrs.automaticSpacing.before;
    }
    if (attrs.spacingExplicit) {
      delete attrs.spacingExplicit.before;
    }
  }
  if (index < count - 1) {
    delete attrs.keepNext;
    delete attrs.runInWithNext;
    delete attrs.listParagraphMarkFontSize;
    if (attrs.spacing) {
      delete attrs.spacing.after;
    }
    if (attrs.automaticSpacing) {
      delete attrs.automaticSpacing.after;
    }
    if (attrs.spacingExplicit) {
      delete attrs.spacingExplicit.after;
    }
  }
  return attrs;
}

type SplitParagraphAtPageBreaksOptions = {
  pageBreaks: readonly PageBreakRunProjection[];
  paragraph: ParagraphBlock;
  splitPageBreakAndParagraphMark: boolean;
};

const runIsZeroWidthBoundaryMarker = (run: Run): boolean => {
  switch (run.kind) {
    case "text":
      return run.text.length === 0;
    case "field":
      return (run.fallback ?? "").length === 0;
    case "renderedPageBreak":
      return true;
    case "image":
    case "lineBreak":
    case "math":
    case "tab":
      return false;
    default: {
      const unsupported: never = run;
      return unsupported;
    }
  }
};

/** Every page-break run under a node, nested tables and text boxes included. */
export const countPageBreakRuns = (node: PMNode): number => {
  let count = node.type.name === "pageBreakRun" ? 1 : 0;
  node.descendants((descendant) => {
    if (descendant.type.name === "pageBreakRun") {
      count += 1;
    }
    return true;
  });
  return count;
};

/** Decide leading-break eligibility from the exact projected runs consumed by layout. */
export const hasSingleLeadingProjectedPageBreak = (
  runs: readonly Run[],
  pageBreaks: readonly PageBreakRunProjection[],
): boolean => {
  if (pageBreaks.length !== 1) {
    return false;
  }
  const partitioned = partitionRunsAtPageBreaks(runs, pageBreaks);
  if (partitioned.type === "overlap") {
    return false;
  }
  return partitioned.partitions[0]?.before.every(runIsZeroWidthBoundaryMarker) === true;
};

export function splitParagraphAtPageBreaks({
  pageBreaks,
  paragraph,
  splitPageBreakAndParagraphMark,
}: SplitParagraphAtPageBreaksOptions): (ParagraphBlock | PageBreakBlock)[] {
  if (pageBreaks.length === 0) {
    return [paragraph];
  }

  const result: (ParagraphBlock | PageBreakBlock)[] = [];
  let fragmentStart = paragraph.pmStart ?? pageBreaks[0]!.pmStart;
  let emittedParagraph = false;

  const appendParagraph = (runs: Run[], pmStart: number, pmEnd: number): void => {
    const fragment: ParagraphBlock = {
      ...paragraph,
      id: emittedParagraph ? nextBlockId() : paragraph.id,
      runs,
      pmStart,
      pmEnd,
    };
    if (emittedParagraph) {
      delete fragment.bookmarks;
    }
    result.push(fragment);
    emittedParagraph = true;
  };

  const partitioned = partitionRunsAtPageBreaks(paragraph.runs, pageBreaks);
  if (partitioned.type === "overlap") {
    panic("An inline layout run overlaps an explicit page-break carrier");
  }
  for (const { before, pageBreak } of partitioned.partitions) {
    if (before.length > 0) {
      appendParagraph(before, fragmentStart, pageBreak.pmStart);
    }
    result.push({
      kind: "pageBreak",
      id: nextBlockId(),
      pmStart: pageBreak.pmStart,
      pmEnd: pageBreak.pmEnd,
      ...pageBreak.trackedChange,
    });
    fragmentStart = pageBreak.pmEnd;
  }

  const remainingRuns = partitioned.remaining;
  let paragraphMarkCarrier: ParagraphBlock | undefined;
  if (remainingRuns.length > 0) {
    appendParagraph(remainingRuns, fragmentStart, paragraph.pmEnd ?? fragmentStart);
  } else if (
    paragraph.runs.length === 0 ||
    (splitPageBreakAndParagraphMark &&
      pageBreaks.at(-1)?.pmEnd === (paragraph.pmEnd ?? fragmentStart) - 1)
  ) {
    appendParagraph([], fragmentStart, paragraph.pmEnd ?? fragmentStart);
    const carrier = result.at(-1);
    if (
      carrier?.kind === "paragraph" &&
      !splitPageBreakAndParagraphMark &&
      !hasVisibleParagraphPayload(carrier.attrs ?? {})
    ) {
      carrier.attrs = { ...carrier.attrs, suppressEmptyParagraphHeight: true };
      paragraphMarkCarrier = carrier;
    }
  }

  const fragments = result.filter((block): block is ParagraphBlock => block.kind === "paragraph");
  for (let index = 0; index < fragments.length; index += 1) {
    const fragment = fragments[index];
    if (!fragment) {
      continue;
    }
    const attrs = paragraphFragmentAttrs(fragment.attrs, index, fragments.length);
    if (attrs) {
      fragment.attrs = attrs;
    } else {
      delete fragment.attrs;
    }
  }
  if (paragraphMarkCarrier?.attrs) {
    // Without the w:splitPgBreakAndParaMark compatibility setting, the
    // paragraph mark stays on the page that ends with the break. The carrier
    // only maps that mark after the break, so the paragraph's before/after
    // spacing belongs to the previous page and must not open the next one:
    // otherwise a following nextPage section cannot reuse that still-blank
    // page and strands an empty sheet.
    paragraphMarkCarrier.attrs = withoutParagraphSpacing(paragraphMarkCarrier.attrs);
  }
  return result;
}

function withoutParagraphSpacing(source: ParagraphAttrs): ParagraphAttrs {
  const attrs: ParagraphAttrs = { ...source };
  if (attrs.spacing) {
    const { before: _before, after: _after, ...lineSpacing } = attrs.spacing;
    if (Object.keys(lineSpacing).length > 0) {
      attrs.spacing = lineSpacing;
    } else {
      delete attrs.spacing;
    }
  }
  delete attrs.automaticSpacing;
  delete attrs.spacingExplicit;
  return attrs;
}
