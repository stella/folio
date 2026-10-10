import type { Node as PMNode } from "prosemirror-model";
import type { SectionProperties } from "../../types/document";
import type { NumberingMap } from "../../docx/numberingParser";
import { effectiveParagraphNumbering, paragraphNumberingAttr } from "../numberingAttr";
import { resolveParagraphNumbering } from "../../docx/numberingReference";
import { expectParagraphAttrs } from "../attrs";
import { resolveParagraphDefaultTextFormatting } from "../styles/paragraphStyleCascade";
import { rejectedListRenderingPatch } from "../listRendering";
import { listIndentationProvenancePatch } from "../styles/resolvedStyleAttrs";
import { paragraphIndentationFromFormatting } from "../paragraphIndentation";
import type { RunStyleResolver } from "../runStyleFormatting";
import type { ParagraphPropertyChangeAttrs } from "../schema/nodes";
import {
  removeParagraphPropertyChanges,
  paragraphRejectAttrPatch,
  paragraphRejectOriginalFormatting,
  sectionRejectProperties,
} from "./propertyChangeScope";

type ResolveParagraphPropertiesOptions = {
  node: PMNode;
  boundaryCovered: boolean;
  mode: "accept" | "reject";
  revisionSet: ReadonlySet<number> | null;
  styleResolver: RunStyleResolver | null;
  numbering: NumberingMap | null;
};

/** Resolve paragraph and section property records before resolving their inline content. */
export const resolveParagraphChangeAttrs = ({
  node,
  boundaryCovered,
  mode,
  revisionSet,
  styleResolver,
  numbering,
}: ResolveParagraphPropertiesOptions): Record<string, unknown> | null => {
  let nextAttrs: Record<string, unknown> | null = null;

  // Process paragraph property changes (w:pPrChange)
  const propertyChanges = expectParagraphAttrs(node)._propertyChanges;

  if (Array.isArray(propertyChanges) && propertyChanges.length > 0 && boundaryCovered) {
    const matchesPropertyChange = (change: ParagraphPropertyChangeAttrs) =>
      revisionSet === null || revisionSet.has(change.info.id);
    if (propertyChanges.some(matchesPropertyChange)) {
      const rejection =
        mode === "reject"
          ? removeParagraphPropertyChanges(propertyChanges, matchesPropertyChange)
          : null;
      let remaining: ParagraphPropertyChangeAttrs[];
      if (rejection === null) {
        remaining = propertyChanges.filter((change) => !matchesPropertyChange(change));
      } else if (rejection.type === "unchanged") {
        remaining = propertyChanges;
      } else {
        remaining = rejection.remaining;
      }
      nextAttrs = {
        ...node.attrs,
        _propertyChanges: remaining.length > 0 ? remaining : null,
      };
      if (rejection?.type === "restore-previous") {
        // Word stores the complete old pPr in the pPrChange, so a
        // reject restores it WHOLESALE within CT_PPrBase scope: a
        // property the change ADDED resets too. Out-of-scope attrs
        // (inline sectPr, paragraph-mark rPr, identity) survive; see
        // propertyChangeScope.ts. Earlier removed runs were folded
        // into the next retained entry, so only a removed trailing
        // run changes the live properties now.
        const currentAttrs = expectParagraphAttrs(node);
        const inheritedAlignment = currentAttrs.alignmentFromStyle;
        const restoredStyleId = rejection.previousFormatting?.styleId ?? undefined;
        const sameStyle = restoredStyleId === (currentAttrs.styleId ?? undefined);
        let previousFormattingFromStyle = sameStyle
          ? currentAttrs._styleResolvedFormatting
          : undefined;
        if (styleResolver) {
          previousFormattingFromStyle =
            styleResolver.resolveParagraphStyle(restoredStyleId).paragraphFormatting;
        } else if (inheritedAlignment !== undefined) {
          previousFormattingFromStyle = {
            ...previousFormattingFromStyle,
            alignment: inheritedAlignment,
          };
        }
        let styleNumbering = previousFormattingFromStyle?.numPr;
        if (!styleResolver && sameStyle) {
          styleNumbering = currentAttrs.numPrFromStyle ?? undefined;
        }
        const inheritedNumbering =
          styleNumbering == null
            ? null
            : paragraphNumberingAttr(resolveParagraphNumbering(styleNumbering));
        const recordedNumbering = rejection.previousFormatting?.numPr;
        const restoredNumPrFromStyle =
          recordedNumbering?.kind === "reference" || recordedNumbering?.kind === "none"
            ? null
            : inheritedNumbering;
        const effectiveNumbering = effectiveParagraphNumbering({
          numPr: recordedNumbering,
          numPrFromStyle: restoredNumPrFromStyle,
        });
        const restoredNumbering =
          effectiveNumbering === undefined ? null : paragraphNumberingAttr(effectiveNumbering);
        const restoredFormatting = paragraphRejectOriginalFormatting(
          rejection.previousFormatting,
          node.attrs["_originalFormatting"],
        );
        const indentation = listIndentationProvenancePatch({
          direct: paragraphIndentationFromFormatting(restoredFormatting ?? undefined),
          styleFormatting: previousFormattingFromStyle,
          numberingSource: recordedNumbering?.kind === "reference" ? "paragraph" : "style",
          numPr:
            restoredNumbering?.kind === "reference"
              ? { numId: restoredNumbering.numId, ilvl: restoredNumbering.ilvl ?? 0 }
              : undefined,
          numbering,
        });
        const inheritedFormatting = {
          ...previousFormattingFromStyle,
          ...indentation._resolvedFormatting,
        };
        Object.assign(
          nextAttrs,
          paragraphRejectAttrPatch(rejection.previousFormatting, inheritedFormatting),
          rejectedListRenderingPatch({
            current: expectParagraphAttrs(node),
            previousFormatting: rejection.previousFormatting,
            numbering,
            restoredNumbering,
          }),
          {
            // The stated and style tiers stay separate; effective membership
            // above is for rendering and indentation provenance only.
            numPr: recordedNumbering ?? null,
            numPrFromStyle: restoredNumPrFromStyle,
          },
        );
        Object.assign(nextAttrs, indentation);
        nextAttrs["_originalFormatting"] = restoredFormatting;
        if (styleResolver) {
          nextAttrs["defaultTextFormatting"] =
            resolveParagraphDefaultTextFormatting(
              rejection.previousFormatting?.styleId,
              restoredFormatting ?? undefined,
              styleResolver,
            ) ?? null;
        }
      }
    }
  }

  // Process inline section-property changes (w:sectPrChange) carried
  // on the paragraph's `_sectionProperties` attr.
  const sectionProperties = expectParagraphAttrs(node)._sectionProperties;
  const sectionChanges = sectionProperties?.propertyChanges;
  if (
    sectionProperties &&
    Array.isArray(sectionChanges) &&
    sectionChanges.length > 0 &&
    boundaryCovered
  ) {
    const matches = sectionChanges.filter(
      (c) => revisionSet === null || (c.info && revisionSet.has(c.info.id)),
    );
    if (matches.length > 0) {
      const remaining = sectionChanges.filter(
        (c) => revisionSet !== null && (!c.info || !revisionSet.has(c.info.id)),
      );
      let restored: SectionProperties = { ...sectionProperties };
      if (mode === "reject") {
        for (const change of matches.toReversed()) {
          restored = sectionRejectProperties({
            live: restored,
            previousProperties: change.previousProperties,
            previousReferences: change.previousReferences,
          });
        }
      }
      delete restored.propertyChanges;
      if (remaining.length > 0) {
        restored.propertyChanges = remaining;
      }
      nextAttrs = nextAttrs ?? { ...node.attrs };
      nextAttrs["_sectionProperties"] = restored;
    }
  }
  return nextAttrs;
};
