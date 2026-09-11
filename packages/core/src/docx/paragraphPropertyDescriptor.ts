import type { ParagraphFormatting } from "../types/document";

export const PARAGRAPH_PROPERTY_OWNER = {
  pPrBase: "pPr-base",
  derived: "derived",
  paragraphMark: "paragraph-mark",
} as const;

export const PARAGRAPH_PROPERTY_PROJECTION = {
  direct: "direct",
  boolean: "boolean",
  alignment: "alignment",
  direction: "direction",
  spacing: "spacing",
  indentation: "indentation",
  opaque: "opaque",
  numberingProvenance: "numbering-provenance",
  preserveLive: "preserve-live",
} as const;

export const PARAGRAPH_STYLE_TRANSITION = {
  replace: "replace",
  numbering: "numbering",
  identity: "identity",
  preserve: "preserve",
} as const;

export type ParagraphPropertyDescriptor = {
  owner: (typeof PARAGRAPH_PROPERTY_OWNER)[keyof typeof PARAGRAPH_PROPERTY_OWNER];
  projection: (typeof PARAGRAPH_PROPERTY_PROJECTION)[keyof typeof PARAGRAPH_PROPERTY_PROJECTION];
  styleTransition: (typeof PARAGRAPH_STYLE_TRANSITION)[keyof typeof PARAGRAPH_STYLE_TRANSITION];
};

/**
 * Total ownership/restoration contract for `ParagraphFormatting`.
 *
 * `w:pPrChange` stores CT_PPrBase, while paragraph-mark rPr and derived
 * style/numbering provenance live outside that payload. Keeping every model
 * property in one exhaustive descriptor prevents snapshots, rejection,
 * serialization, and style transitions from silently growing different
 * interpretations of that boundary when `ParagraphFormatting` gains a field.
 */
export const PARAGRAPH_FORMATTING_PROPERTY_DESCRIPTOR = {
  alignment: {
    owner: "pPr-base",
    projection: "alignment",
    styleTransition: "replace",
  },
  bidi: {
    owner: "pPr-base",
    projection: "direction",
    styleTransition: "replace",
  },
  kinsoku: {
    owner: "pPr-base",
    projection: "boolean",
    styleTransition: "replace",
  },
  overflowPunctuation: {
    owner: "pPr-base",
    projection: "boolean",
    styleTransition: "replace",
  },
  spaceBefore: {
    owner: "pPr-base",
    projection: "spacing",
    styleTransition: "replace",
  },
  spaceAfter: {
    owner: "pPr-base",
    projection: "spacing",
    styleTransition: "replace",
  },
  lineSpacing: {
    owner: "pPr-base",
    projection: "spacing",
    styleTransition: "replace",
  },
  lineSpacingRule: {
    owner: "pPr-base",
    projection: "spacing",
    styleTransition: "replace",
  },
  snapToGrid: {
    owner: "pPr-base",
    projection: "boolean",
    styleTransition: "replace",
  },
  beforeAutospacing: {
    owner: "pPr-base",
    projection: "spacing",
    styleTransition: "replace",
  },
  afterAutospacing: {
    owner: "pPr-base",
    projection: "spacing",
    styleTransition: "replace",
  },
  spacingExplicit: {
    owner: "pPr-base",
    projection: "spacing",
    styleTransition: "replace",
  },
  indentLeft: {
    owner: "pPr-base",
    projection: "indentation",
    styleTransition: "replace",
  },
  indentRight: {
    owner: "pPr-base",
    projection: "indentation",
    styleTransition: "replace",
  },
  indentFirstLine: {
    owner: "pPr-base",
    projection: "indentation",
    styleTransition: "replace",
  },
  hangingIndent: {
    owner: "pPr-base",
    projection: "indentation",
    styleTransition: "replace",
  },
  numberingLevelIndent: {
    owner: "derived",
    projection: "numbering-provenance",
    styleTransition: "replace",
  },
  borders: {
    owner: "pPr-base",
    projection: "direct",
    styleTransition: "replace",
  },
  shading: {
    owner: "pPr-base",
    projection: "direct",
    styleTransition: "replace",
  },
  tabs: {
    owner: "pPr-base",
    projection: "direct",
    styleTransition: "replace",
  },
  keepNext: {
    owner: "pPr-base",
    projection: "boolean",
    styleTransition: "replace",
  },
  keepLines: {
    owner: "pPr-base",
    projection: "boolean",
    styleTransition: "replace",
  },
  widowControl: {
    owner: "pPr-base",
    projection: "boolean",
    styleTransition: "replace",
  },
  pageBreakBefore: {
    owner: "pPr-base",
    projection: "boolean",
    styleTransition: "replace",
  },
  contextualSpacing: {
    owner: "pPr-base",
    projection: "boolean",
    styleTransition: "replace",
  },
  numPr: {
    owner: "pPr-base",
    projection: "direct",
    styleTransition: "numbering",
  },
  numPrFromStyle: {
    owner: "derived",
    projection: "numbering-provenance",
    styleTransition: "numbering",
  },
  outlineLevel: {
    owner: "pPr-base",
    projection: "direct",
    styleTransition: "replace",
  },
  styleId: {
    owner: "pPr-base",
    projection: "direct",
    styleTransition: "identity",
  },
  frame: {
    owner: "pPr-base",
    projection: "opaque",
    styleTransition: "replace",
  },
  suppressLineNumbers: {
    owner: "pPr-base",
    projection: "boolean",
    styleTransition: "replace",
  },
  suppressAutoHyphens: {
    owner: "pPr-base",
    projection: "boolean",
    styleTransition: "replace",
  },
  runProperties: {
    owner: "paragraph-mark",
    projection: "preserve-live",
    styleTransition: "preserve",
  },
  runInWithNext: {
    owner: "paragraph-mark",
    projection: "preserve-live",
    styleTransition: "preserve",
  },
} as const satisfies Record<keyof ParagraphFormatting, ParagraphPropertyDescriptor>;

// SAFETY: `Object.keys` erases keys that the total descriptor establishes.
export const PARAGRAPH_FORMATTING_PROPERTY_KEY_LIST = Object.freeze(
  Object.keys(PARAGRAPH_FORMATTING_PROPERTY_DESCRIPTOR) as (
    keyof typeof PARAGRAPH_FORMATTING_PROPERTY_DESCRIPTOR
  )[],
);

type FormattingKeysOwnedBy<Owner extends ParagraphPropertyDescriptor["owner"]> = {
  [Key in keyof typeof PARAGRAPH_FORMATTING_PROPERTY_DESCRIPTOR]: (typeof PARAGRAPH_FORMATTING_PROPERTY_DESCRIPTOR)[Key]["owner"] extends Owner
    ? Key
    : never;
}[keyof typeof PARAGRAPH_FORMATTING_PROPERTY_DESCRIPTOR];

/** Exact authored CT_PPrBase payload, before style or numbering resolution. */
export type AuthoredParagraphPropertyKey = FormattingKeysOwnedBy<"pPr-base">;
export type AuthoredParagraphProperties = Pick<
  ParagraphFormatting,
  AuthoredParagraphPropertyKey
>;

/** Paragraph-mark properties intentionally kept outside CT_PPrBase provenance. */
export type ParagraphMarkPropertyKey = FormattingKeysOwnedBy<"paragraph-mark">;
export type ParagraphMarkProperties = Pick<ParagraphFormatting, ParagraphMarkPropertyKey>;

/** Effective-only or derived properties intentionally kept outside authored pPr. */
export type DerivedParagraphPropertyKey = FormattingKeysOwnedBy<"derived">;
export type DerivedParagraphProperties = Pick<ParagraphFormatting, DerivedParagraphPropertyKey>;

/** Exhaustive modeled contents of one raw `w:pPr` source capture. */
export type ParagraphPropertySourceFingerprint = Readonly<{
  pPrBase: AuthoredParagraphProperties;
  paragraphMark: ParagraphMarkProperties;
}>;

const clonePropertyValue = (value: unknown): unknown => {
  if (Array.isArray(value)) {
    return value.map(clonePropertyValue);
  }
  if (typeof value === "object" && value !== null) {
    const clone: Record<string, unknown> = {};
    for (const [key, entry] of Object.entries(value)) {
      clone[key] = clonePropertyValue(entry);
    }
    return clone;
  }
  return value;
};

const selectParagraphPropertiesOwnedBy = <
  Owner extends ParagraphPropertyDescriptor["owner"],
>(
  formatting: ParagraphFormatting | null | undefined,
  owner: Owner,
): Pick<ParagraphFormatting, FormattingKeysOwnedBy<Owner>> => {
  const selected: Partial<ParagraphFormatting> = {};
  if (formatting) {
    for (const rawKey of PARAGRAPH_FORMATTING_PROPERTY_KEY_LIST) {
      const descriptor = PARAGRAPH_FORMATTING_PROPERTY_DESCRIPTOR[rawKey];
      if (descriptor.owner !== owner) {
        continue;
      }
      const value = Reflect.get(formatting, rawKey);
      if (value !== undefined) {
        Reflect.set(selected, rawKey, clonePropertyValue(value));
      }
    }
  }
  // SAFETY: the total descriptor limits selected keys to the requested owner.
  return selected as Pick<ParagraphFormatting, FormattingKeysOwnedBy<Owner>>;
};

/** Select exact CT_PPrBase-owned fields from a broader formatting object. */
export const selectAuthoredParagraphProperties = (
  formatting: ParagraphFormatting | null | undefined,
): AuthoredParagraphProperties =>
  selectParagraphPropertiesOwnedBy(formatting, PARAGRAPH_PROPERTY_OWNER.pPrBase);

/** Select exact paragraph-mark `w:rPr` fields from broader formatting. */
export const selectParagraphMarkProperties = (
  formatting: ParagraphFormatting | null | undefined,
): ParagraphMarkProperties =>
  selectParagraphPropertiesOwnedBy(formatting, PARAGRAPH_PROPERTY_OWNER.paragraphMark);

export const paragraphPropertySourceFingerprintFromFormatting = (
  formatting: ParagraphFormatting | null | undefined,
): ParagraphPropertySourceFingerprint => ({
  pPrBase: selectAuthoredParagraphProperties(formatting),
  paragraphMark: selectParagraphMarkProperties(formatting),
});

export const paragraphPropertySourceFingerprintFromParts = (
  pPrBase: AuthoredParagraphProperties,
  paragraphMark: ParagraphMarkProperties,
): ParagraphPropertySourceFingerprint => ({
  pPrBase: selectAuthoredParagraphProperties(pPrBase),
  paragraphMark: selectParagraphMarkProperties(paragraphMark),
});

const canonicalJsonValue = (value: unknown): unknown => {
  if (Array.isArray(value)) {
    return value.map(canonicalJsonValue);
  }
  if (typeof value === "object" && value !== null) {
    const result: Record<string, unknown> = {};
    for (const key of Object.keys(value).toSorted()) {
      result[key] = canonicalJsonValue(Reflect.get(value, key));
    }
    return result;
  }
  return value;
};

/** Canonical raw-replay identity for the exact authored CT_PPrBase payload. */
export const canonicalAuthoredParagraphPropertiesJson = (
  formatting: ParagraphFormatting | null | undefined,
): string => JSON.stringify(canonicalJsonValue(selectAuthoredParagraphProperties(formatting)));

/** Canonical replay identity for every modeled property inside raw `w:pPr`. */
export const canonicalParagraphPropertySourceFingerprintJson = (
  fingerprint: ParagraphPropertySourceFingerprint,
): string => JSON.stringify(canonicalJsonValue(fingerprint));

const formattingKeys = (
  predicate: (descriptor: ParagraphPropertyDescriptor) => boolean,
): readonly string[] =>
  Object.entries(PARAGRAPH_FORMATTING_PROPERTY_DESCRIPTOR)
    .filter(([, descriptor]) => predicate(descriptor))
    .map(([key]) => key);

export const PPR_OPAQUE_FORMATTING_KEYS = formattingKeys(
  ({ owner, projection }) =>
    owner === PARAGRAPH_PROPERTY_OWNER.pPrBase &&
    projection === PARAGRAPH_PROPERTY_PROJECTION.opaque,
);

export const PPR_CHANGE_SCOPED_FORMATTING_KEYS = formattingKeys(
  ({ owner }) => owner === PARAGRAPH_PROPERTY_OWNER.pPrBase,
);

export const PPR_PARAGRAPH_MARK_FORMATTING_KEYS = formattingKeys(
  ({ owner }) => owner === PARAGRAPH_PROPERTY_OWNER.paragraphMark,
);

export const PPR_STYLE_REPLACED_FORMATTING_KEYS = formattingKeys(
  ({ styleTransition }) => styleTransition === PARAGRAPH_STYLE_TRANSITION.replace,
);

export const PARAGRAPH_FORMATTING_PROPERTY_KEYS: ReadonlySet<string> = new Set(
  PARAGRAPH_FORMATTING_PROPERTY_KEY_LIST,
);
