import type { Node as PMNode } from "prosemirror-model";
import type { ParagraphFormatting, SectionProperties } from "../../types/document";
import { expectParagraphAttrs } from "../attrs";
import { resolveParagraphDefaultTextFormatting } from "../conversion/toProseDoc";
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
};

/** Resolve paragraph and section property records before resolving their inline content. */
export const resolveParagraphChangeAttrs = ({
  node,
  boundaryCovered,
  mode,
  revisionSet,
  styleResolver,
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
        const inheritedAlignment = expectParagraphAttrs(node).alignmentFromStyle;
        let previousFormattingFromStyle: ParagraphFormatting | undefined;
        if (styleResolver) {
          previousFormattingFromStyle = styleResolver.resolveParagraphStyle(
            rejection.previousFormatting?.styleId,
          ).paragraphFormatting;
        } else if (inheritedAlignment !== undefined) {
          previousFormattingFromStyle = { alignment: inheritedAlignment };
        }
        Object.assign(
          nextAttrs,
          paragraphRejectAttrPatch(rejection.previousFormatting, previousFormattingFromStyle),
        );
        const restoredFormatting = paragraphRejectOriginalFormatting(
          rejection.previousFormatting,
          node.attrs["_originalFormatting"],
        );
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
