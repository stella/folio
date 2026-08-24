import type { ParagraphAttrs, ParagraphBlock } from "../layout-engine/types";
import { isRtlParagraph } from "./paragraphBaseDirection";

type ParagraphAlignment = NonNullable<ParagraphAttrs["alignment"]>;

export type PhysicalParagraphInlineLayout = {
  alignment: ParagraphAlignment;
  explicitAlignment?: ParagraphAlignment;
  indentLeft: number;
  indentRight: number;
  isRtl: boolean;
};

const mirrorHorizontalAlignment = (
  alignment: ParagraphAlignment | undefined,
): ParagraphAlignment | undefined => {
  if (alignment === "left") {
    return "right";
  }
  if (alignment === "right") {
    return "left";
  }
  return alignment;
};

/** Resolve logical OOXML paragraph sides into physical inline geometry. */
export const resolvePhysicalParagraphInlineLayout = (
  block: ParagraphBlock,
): PhysicalParagraphInlineLayout => {
  const attrs = block.attrs;
  const isRtl = isRtlParagraph(block);
  const mirrorsHorizontalSides = attrs?.bidi === true;
  const explicitAlignment = mirrorsHorizontalSides
    ? mirrorHorizontalAlignment(attrs?.alignment)
    : attrs?.alignment;
  const indentLeft = attrs?.indent?.left ?? 0;
  const indentRight = attrs?.indent?.right ?? 0;

  return {
    alignment: explicitAlignment ?? (isRtl ? "right" : "left"),
    ...(explicitAlignment === undefined ? {} : { explicitAlignment }),
    indentLeft: mirrorsHorizontalSides ? indentRight : indentLeft,
    indentRight: mirrorsHorizontalSides ? indentLeft : indentRight,
    isRtl,
  };
};
