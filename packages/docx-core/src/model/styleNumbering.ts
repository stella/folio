import type {
  NonEmptyNumberingLevelIndentGeometry,
  NumberingLevelIndentGeometry,
  ParagraphFormatting,
} from "./formatting";
import type { StyleDefinitions } from "./styles";

export type NumberingIdentity = NonNullable<ParagraphFormatting["numPr"]>;

/** Compare numbering identities, treating an omitted level as level zero. */
export const numberingIdentityEqual = (
  left: NumberingIdentity | null | undefined,
  right: NumberingIdentity | null | undefined,
): boolean => {
  if (left == null || right == null) {
    return left == null && right == null;
  }
  return left.numId === right.numId && (left.ilvl ?? 0) === (right.ilvl ?? 0);
};

/** Narrow numbering geometry to the structurally non-empty provenance payload. */
export const isNonEmptyNumberingLevelIndentGeometry = (
  geometry: NumberingLevelIndentGeometry | null | undefined,
): geometry is NonEmptyNumberingLevelIndentGeometry =>
  geometry !== undefined &&
  geometry !== null &&
  (geometry.indentLeft !== undefined ||
    geometry.indentRight !== undefined ||
    geometry.indentFirstLine !== undefined);

/** Return numbering-owned indentation only when its identity and baseline agree. */
export const verifiedNumberingLevelIndent = (
  formatting: ParagraphFormatting | null | undefined,
): NonEmptyNumberingLevelIndentGeometry | undefined => {
  const active = formatting?.numPr;
  const provenance = formatting?.numberingLevelIndent;
  if (
    active?.numId === undefined ||
    active.numId === 0 ||
    provenance?.type !== "owned" ||
    !isNonEmptyNumberingLevelIndentGeometry(provenance.owned) ||
    !isNonEmptyNumberingLevelIndentGeometry(provenance.baseline) ||
    !numberingIdentityEqual(active, provenance)
  ) {
    return undefined;
  }
  for (const key of [
    "indentLeft",
    "indentRight",
    "indentFirstLine",
    "hangingIndent",
  ] as const) {
    if (
      provenance.owned[key] !== undefined &&
      provenance.owned[key] !== provenance.baseline[key]
    ) {
      return undefined;
    }
  }
  return provenance.owned;
};

export type ParagraphStyleNumberingResolver = (
  formatting: ParagraphFormatting | undefined,
) => boolean;

/** Build a bounded resolver for numbering supplied by paragraph-style inheritance. */
export const createParagraphStyleNumberingResolver = (
  styles: StyleDefinitions | undefined,
): ParagraphStyleNumberingResolver => {
  if (!styles) {
    return () => false;
  }

  const styleById = new Map(styles.styles.map((style) => [style.styleId, style]));
  const cache = new Map<string, NumberingIdentity | undefined>();
  const resolveStyleNumbering = (styleId: string): NumberingIdentity | undefined => {
    if (cache.has(styleId)) {
      return cache.get(styleId);
    }

    const chain: (typeof styles.styles)[number][] = [];
    const seen = new Set<string>();
    let style = styleById.get(styleId);
    while (style?.type === "paragraph" && !seen.has(style.styleId)) {
      seen.add(style.styleId);
      chain.push(style);
      style = style.basedOn ? styleById.get(style.basedOn) : undefined;
    }

    let numbering: NumberingIdentity | undefined;
    for (const current of chain.toReversed()) {
      if (current.pPr?.numPr !== undefined) {
        numbering = { ...numbering, ...current.pPr.numPr };
      }
    }
    cache.set(styleId, numbering);
    return numbering;
  };

  return (formatting) => {
    const provenance = formatting?.numPrFromStyle;
    return (
      formatting?.styleId !== undefined &&
      provenance?.numId !== undefined &&
      numberingIdentityEqual(resolveStyleNumbering(formatting.styleId), provenance)
    );
  };
};
