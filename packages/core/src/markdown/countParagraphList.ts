/**
 * Advance the Markdown walk's list counter past one paragraph, as the page
 * counts it. Every paragraph the page counts, Markdown counts, whatever it
 * renders as and whether or not it writes it: the items after a numbered
 * heading continue from its number, and the paragraphs of a text box count
 * right after the paragraph the box is drawn in (where the page and the
 * content reader count them), although Markdown writes no text box.
 */

import { getCachedNumberingMap } from "../docx/numberingParser";
import {
  BLOCK_TREE_DESCENT,
  visitBlockTreeRecords,
} from "../docx/paragraphTraversal";
import { listLabelAttrsFromRendering } from "../prosemirror/listLabels";
import {
  effectiveParagraphNumberingReference,
  paragraphNumberingAttr,
} from "../prosemirror/numberingAttr";
import { resolveParagraphList } from "../prosemirror/paragraphList";
import type {
  DocxPackage,
  ListRendering,
  Paragraph,
  Shape,
} from "../types/document";
import type { RenderContext } from "./types";

/** Explicit projection carried by canonical models that are not attached to a package. */
const standaloneListRendering = (para: Paragraph): ListRendering | undefined =>
  para.listRendering;

export type CountedParagraphList = {
  /**
   * The list the paragraph shows: its own numbering's, or, while a pending
   * property change removed it, the one that numbering showed.
   */
  visibleList: ListRendering | undefined;
  /** The label the counter gave the paragraph; absent when it shows none. */
  label: string | undefined;
};

export function countParagraphList(
  ctx: RenderContext,
  pkg: DocxPackage | undefined,
  para: Paragraph,
): CountedParagraphList {
  const inherited = ctx.paragraphNumbering(para.formatting?.styleId);
  const inheritedReference = effectiveParagraphNumberingReference({
    numPr: undefined,
    numPrFromStyle: inherited,
  });
  const sources = {
    numPr: para.formatting?.numPr,
    numPrFromStyle: para.formatting?.numPrFromStyle,
  };
  const numbering = pkg?.numbering
    ? getCachedNumberingMap(pkg.numbering)
    : null;
  const resolvedList = resolveParagraphList(
    { ...sources, inheritedReference: inheritedReference ?? null },
    numbering,
    standaloneListRendering(para),
  );
  const resolvedRendering =
    resolvedList.kind === "marker" ? resolvedList.rendering : undefined;
  const counterAttrs = listLabelAttrsFromRendering(resolvedRendering, {
    ...sources,
    numPrFromStyle:
      resolvedList.kind === "none"
        ? undefined
        : (inherited ?? sources.numPrFromStyle),
  });
  if (
    sources.numPr === undefined &&
    sources.numPrFromStyle !== undefined &&
    inheritedReference === undefined &&
    para.propertyChanges !== undefined &&
    para.propertyChanges.length > 0
  ) {
    counterAttrs.numPr = paragraphNumberingAttr({ kind: "none" });
  }
  if (para.pPrMark) counterAttrs.pPrMark = para.pPrMark;
  let originalList: ListRendering | undefined;
  // Clean output is the document with every change accepted: a pending
  // property change's earlier numbering is what rejecting restores, and it
  // neither numbers nor counts this paragraph there.
  if (para.propertyChanges && ctx.opts.trackedChanges !== "clean") {
    counterAttrs._propertyChanges = para.propertyChanges.map(
      ({ previousFormatting, ...change }) => {
        if (!previousFormatting) return change;
        const previousInherited = ctx.paragraphNumbering(
          previousFormatting.styleId,
        );
        const previousInheritedReference = effectiveParagraphNumberingReference(
          {
            numPr: undefined,
            numPrFromStyle: previousInherited,
          },
        );
        const previousSources = {
          numPr: previousFormatting.numPr,
          numPrFromStyle: previousFormatting.numPrFromStyle,
        };
        const previousList = resolveParagraphList(
          {
            ...previousSources,
            inheritedReference: previousInheritedReference ?? null,
          },
          numbering,
        );
        const previousRendering =
          previousList.kind === "marker" ? previousList.rendering : undefined;
        originalList ??= previousRendering;
        const previousAttrs = listLabelAttrsFromRendering(previousRendering, {
          ...previousSources,
          numPrFromStyle:
            previousList.kind === "none"
              ? undefined
              : (previousInherited ?? previousSources.numPrFromStyle),
        });
        return {
          ...change,
          previousFormatting: {
            ...previousFormatting,
            ...previousAttrs,
            ...(previousAttrs.numPr === undefined
              ? { numPr: null }
              : { numPr: previousAttrs.numPr }),
            ...(previousAttrs.numPrFromStyle === undefined
              ? { numPrFromStyle: null }
              : { numPrFromStyle: previousAttrs.numPrFromStyle }),
          },
        };
      },
    );
  }
  const countedList = ctx.nextListLabel(counterAttrs, resolvedRendering);
  return {
    visibleList:
      resolvedRendering ??
      (sources.numPr?.kind === "levelOnly" ? undefined : originalList),
    label: countedList.kind === "marker" ? countedList.label : undefined,
  };
}

/**
 * Count the paragraphs of a text box Markdown does not write, cells
 * included, in document order. A text box holds no text box, so the walk
 * stops at each paragraph.
 */
export function countTextBoxLists(
  ctx: RenderContext,
  pkg: DocxPackage | undefined,
  shape: Shape,
): void {
  if (shape.textBody === undefined) return;
  visitBlockTreeRecords(shape.textBody.content, (record) => {
    if (record.type !== "paragraph") return BLOCK_TREE_DESCENT.descend;
    countParagraphList(ctx, pkg, record);
    return BLOCK_TREE_DESCENT.prune;
  });
}
