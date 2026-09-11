import type { ParagraphFormatting } from "./formatting";

const freezeDescriptor = <const Value extends Record<string, object>>(value: Value): Value => {
  for (const entry of Object.values(value)) {
    Object.freeze(entry);
  }
  return Object.freeze(value);
};

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

export const PARAGRAPH_PROPERTY_CASCADE = {
  fieldwise: "fieldwise",
  firstLine: "first-line",
  preserveDerived: "preserve-derived",
  replace: "replace",
  runProperties: "run-properties",
  tabStops: "tab-stops",
} as const;

export type ParagraphPropertyDescriptor = {
  owner: (typeof PARAGRAPH_PROPERTY_OWNER)[keyof typeof PARAGRAPH_PROPERTY_OWNER];
  projection: (typeof PARAGRAPH_PROPERTY_PROJECTION)[keyof typeof PARAGRAPH_PROPERTY_PROJECTION];
  styleTransition: (typeof PARAGRAPH_STYLE_TRANSITION)[keyof typeof PARAGRAPH_STYLE_TRANSITION];
  cascade: (typeof PARAGRAPH_PROPERTY_CASCADE)[keyof typeof PARAGRAPH_PROPERTY_CASCADE];
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
export const PARAGRAPH_FORMATTING_PROPERTY_DESCRIPTOR = freezeDescriptor({
  alignment: {
    owner: "pPr-base",
    projection: "alignment",
    styleTransition: "replace",
    cascade: "replace",
  },
  bidi: {
    owner: "pPr-base",
    projection: "direction",
    styleTransition: "replace",
    cascade: "replace",
  },
  kinsoku: {
    owner: "pPr-base",
    projection: "boolean",
    styleTransition: "replace",
    cascade: "replace",
  },
  overflowPunctuation: {
    owner: "pPr-base",
    projection: "boolean",
    styleTransition: "replace",
    cascade: "replace",
  },
  spaceBefore: {
    owner: "pPr-base",
    projection: "spacing",
    styleTransition: "replace",
    cascade: "replace",
  },
  spaceAfter: {
    owner: "pPr-base",
    projection: "spacing",
    styleTransition: "replace",
    cascade: "replace",
  },
  lineSpacing: {
    owner: "pPr-base",
    projection: "spacing",
    styleTransition: "replace",
    cascade: "replace",
  },
  lineSpacingRule: {
    owner: "pPr-base",
    projection: "spacing",
    styleTransition: "replace",
    cascade: "replace",
  },
  snapToGrid: {
    owner: "pPr-base",
    projection: "boolean",
    styleTransition: "replace",
    cascade: "replace",
  },
  beforeAutospacing: {
    owner: "pPr-base",
    projection: "spacing",
    styleTransition: "replace",
    cascade: "replace",
  },
  afterAutospacing: {
    owner: "pPr-base",
    projection: "spacing",
    styleTransition: "replace",
    cascade: "replace",
  },
  spacingExplicit: {
    owner: "pPr-base",
    projection: "spacing",
    styleTransition: "replace",
    cascade: "replace",
  },
  indentLeft: {
    owner: "pPr-base",
    projection: "indentation",
    styleTransition: "replace",
    cascade: "replace",
  },
  indentRight: {
    owner: "pPr-base",
    projection: "indentation",
    styleTransition: "replace",
    cascade: "replace",
  },
  indentFirstLine: {
    owner: "pPr-base",
    projection: "indentation",
    styleTransition: "replace",
    cascade: "first-line",
  },
  hangingIndent: {
    owner: "pPr-base",
    projection: "indentation",
    styleTransition: "replace",
    cascade: "replace",
  },
  numberingLevelIndent: {
    owner: "derived",
    projection: "numbering-provenance",
    styleTransition: "replace",
    cascade: "preserve-derived",
  },
  borders: {
    owner: "pPr-base",
    projection: "direct",
    styleTransition: "replace",
    cascade: "fieldwise",
  },
  shading: {
    owner: "pPr-base",
    projection: "direct",
    styleTransition: "replace",
    cascade: "replace",
  },
  tabs: {
    owner: "pPr-base",
    projection: "direct",
    styleTransition: "replace",
    cascade: "tab-stops",
  },
  keepNext: {
    owner: "pPr-base",
    projection: "boolean",
    styleTransition: "replace",
    cascade: "replace",
  },
  keepLines: {
    owner: "pPr-base",
    projection: "boolean",
    styleTransition: "replace",
    cascade: "replace",
  },
  widowControl: {
    owner: "pPr-base",
    projection: "boolean",
    styleTransition: "replace",
    cascade: "replace",
  },
  pageBreakBefore: {
    owner: "pPr-base",
    projection: "boolean",
    styleTransition: "replace",
    cascade: "replace",
  },
  contextualSpacing: {
    owner: "pPr-base",
    projection: "boolean",
    styleTransition: "replace",
    cascade: "replace",
  },
  numPr: {
    owner: "pPr-base",
    projection: "direct",
    styleTransition: "numbering",
    cascade: "fieldwise",
  },
  numPrFromStyle: {
    owner: "derived",
    projection: "numbering-provenance",
    styleTransition: "numbering",
    cascade: "preserve-derived",
  },
  outlineLevel: {
    owner: "pPr-base",
    projection: "direct",
    styleTransition: "replace",
    cascade: "replace",
  },
  styleId: {
    owner: "pPr-base",
    projection: "direct",
    styleTransition: "identity",
    cascade: "replace",
  },
  frame: {
    owner: "pPr-base",
    projection: "opaque",
    styleTransition: "replace",
    cascade: "fieldwise",
  },
  suppressLineNumbers: {
    owner: "pPr-base",
    projection: "boolean",
    styleTransition: "replace",
    cascade: "replace",
  },
  suppressAutoHyphens: {
    owner: "pPr-base",
    projection: "boolean",
    styleTransition: "replace",
    cascade: "replace",
  },
  runProperties: {
    owner: "paragraph-mark",
    projection: "preserve-live",
    styleTransition: "preserve",
    cascade: "run-properties",
  },
  runInWithNext: {
    owner: "paragraph-mark",
    projection: "preserve-live",
    styleTransition: "preserve",
    cascade: "replace",
  },
} as const satisfies Record<keyof ParagraphFormatting, ParagraphPropertyDescriptor>);

export const PARAGRAPH_PROPERTY_SOURCE_XML_GROUP = Object.freeze({
  alignment: "jc",
  bidi: "bidi",
  kinsoku: "kinsoku",
  overflowPunctuation: "overflowPunct",
  spaceBefore: "spacing",
  spaceAfter: "spacing",
  lineSpacing: "spacing",
  lineSpacingRule: "spacing",
  snapToGrid: "snapToGrid",
  beforeAutospacing: "spacing",
  afterAutospacing: "spacing",
  spacingExplicit: "spacing",
  indentLeft: "ind",
  indentRight: "ind",
  indentFirstLine: "ind",
  hangingIndent: "ind",
  numberingLevelIndent: null,
  borders: "pBdr",
  shading: "shd",
  tabs: "tabs",
  keepNext: "keepNext",
  keepLines: "keepLines",
  widowControl: "widowControl",
  pageBreakBefore: "pageBreakBefore",
  contextualSpacing: "contextualSpacing",
  numPr: "numPr",
  numPrFromStyle: null,
  outlineLevel: "outlineLvl",
  styleId: "pStyle",
  frame: "framePr",
  suppressLineNumbers: "suppressLineNumbers",
  suppressAutoHyphens: "suppressAutoHyphens",
  runProperties: "rPr",
  runInWithNext: "rPr",
} as const satisfies Record<keyof ParagraphFormatting, string | null>);

export type ParagraphPropertySourceXmlGroup = Exclude<
  (typeof PARAGRAPH_PROPERTY_SOURCE_XML_GROUP)[keyof typeof PARAGRAPH_PROPERTY_SOURCE_XML_GROUP],
  null
>;

// SAFETY: `Object.keys` erases keys that the total descriptor establishes.
export const PARAGRAPH_FORMATTING_PROPERTY_KEY_LIST = Object.freeze(
  Object.keys(
    PARAGRAPH_FORMATTING_PROPERTY_DESCRIPTOR,
  ) as (keyof typeof PARAGRAPH_FORMATTING_PROPERTY_DESCRIPTOR)[],
);

type FormattingKeysOwnedBy<Owner extends ParagraphPropertyDescriptor["owner"]> = {
  [Key in keyof typeof PARAGRAPH_FORMATTING_PROPERTY_DESCRIPTOR]: (typeof PARAGRAPH_FORMATTING_PROPERTY_DESCRIPTOR)[Key]["owner"] extends Owner
    ? Key
    : never;
}[keyof typeof PARAGRAPH_FORMATTING_PROPERTY_DESCRIPTOR];

/** Exact authored CT_PPrBase payload, before style or numbering resolution. */
export type AuthoredParagraphPropertyKey = FormattingKeysOwnedBy<"pPr-base">;
export type AuthoredParagraphProperties = Pick<ParagraphFormatting, AuthoredParagraphPropertyKey>;

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

const selectParagraphPropertiesOwnedBy = <Owner extends ParagraphPropertyDescriptor["owner"]>(
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

/**
 * Replace every raw-source-owned field while retaining derived formatting.
 * PM-to-model conversion uses this as the single inverse of the authored and
 * paragraph-mark projections; effective values can never leak into save truth.
 */
export const paragraphFormattingWithPropertySourceFingerprint = (
  formatting: ParagraphFormatting | null | undefined,
  fingerprint: ParagraphPropertySourceFingerprint,
): ParagraphFormatting | undefined => {
  const result: ParagraphFormatting = { ...formatting };
  for (const key of PARAGRAPH_FORMATTING_PROPERTY_KEY_LIST) {
    if (PARAGRAPH_FORMATTING_PROPERTY_DESCRIPTOR[key].owner !== PARAGRAPH_PROPERTY_OWNER.derived) {
      Reflect.deleteProperty(result, key);
    }
  }
  Object.assign(
    result,
    selectAuthoredParagraphProperties(fingerprint.pPrBase),
    selectParagraphMarkProperties(fingerprint.paragraphMark),
  );
  return Object.keys(result).length === 0 ? undefined : result;
};

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

/** Canonical replay identity for every modeled property inside raw `w:pPr`. */
export const canonicalParagraphPropertySourceFingerprintJson = (
  fingerprint: ParagraphPropertySourceFingerprint,
): string => JSON.stringify(canonicalJsonValue(fingerprint));

const canonicalPropertyValueEqual = (left: unknown, right: unknown): boolean =>
  JSON.stringify(canonicalJsonValue(left)) === JSON.stringify(canonicalJsonValue(right));

/**
 * Immutable proof of the modeled raw-property groups changed since capture.
 * Serializers use the groups to replace only owned OOXML children and retain
 * every unknown child from the imported `w:pPr`.
 */
export class ParagraphPropertySourceDelta {
  readonly #changedParagraphMarkKeys: readonly ParagraphMarkPropertyKey[];
  readonly #changedPPrBaseKeys: readonly AuthoredParagraphPropertyKey[];
  readonly #changedXmlGroups: readonly ParagraphPropertySourceXmlGroup[];

  private constructor({
    changedParagraphMarkKeys,
    changedPPrBaseKeys,
    changedXmlGroups,
  }: {
    changedParagraphMarkKeys: ParagraphMarkPropertyKey[];
    changedPPrBaseKeys: AuthoredParagraphPropertyKey[];
    changedXmlGroups: ParagraphPropertySourceXmlGroup[];
  }) {
    this.#changedParagraphMarkKeys = Object.freeze(changedParagraphMarkKeys);
    this.#changedPPrBaseKeys = Object.freeze(changedPPrBaseKeys);
    this.#changedXmlGroups = Object.freeze(changedXmlGroups);
    Object.freeze(this);
  }

  static between(
    before: ParagraphPropertySourceFingerprint,
    after: ParagraphPropertySourceFingerprint,
  ): ParagraphPropertySourceDelta {
    const changedParagraphMarkKeys: ParagraphMarkPropertyKey[] = [];
    const changedPPrBaseKeys: AuthoredParagraphPropertyKey[] = [];
    const changedXmlGroups: ParagraphPropertySourceXmlGroup[] = [];
    const seenGroups = new Set<ParagraphPropertySourceXmlGroup>();
    for (const key of PARAGRAPH_FORMATTING_PROPERTY_KEY_LIST) {
      const descriptor = PARAGRAPH_FORMATTING_PROPERTY_DESCRIPTOR[key];
      if (descriptor.owner === PARAGRAPH_PROPERTY_OWNER.derived) {
        continue;
      }
      const beforeOwner =
        descriptor.owner === PARAGRAPH_PROPERTY_OWNER.pPrBase
          ? before.pPrBase
          : before.paragraphMark;
      const afterOwner =
        descriptor.owner === PARAGRAPH_PROPERTY_OWNER.pPrBase ? after.pPrBase : after.paragraphMark;
      if (
        canonicalPropertyValueEqual(Reflect.get(beforeOwner, key), Reflect.get(afterOwner, key))
      ) {
        continue;
      }
      if (descriptor.owner === PARAGRAPH_PROPERTY_OWNER.pPrBase) {
        // SAFETY: the descriptor owner narrows the key to CT_PPrBase ownership.
        changedPPrBaseKeys.push(key as AuthoredParagraphPropertyKey);
      } else {
        // SAFETY: the only remaining descriptor owner is the paragraph mark.
        changedParagraphMarkKeys.push(key as ParagraphMarkPropertyKey);
      }
      const group = PARAGRAPH_PROPERTY_SOURCE_XML_GROUP[key];
      if (group !== null && !seenGroups.has(group)) {
        seenGroups.add(group);
        changedXmlGroups.push(group);
      }
    }
    return new ParagraphPropertySourceDelta({
      changedParagraphMarkKeys,
      changedPPrBaseKeys,
      changedXmlGroups,
    });
  }

  get changedParagraphMarkKeys(): readonly ParagraphMarkPropertyKey[] {
    return this.#changedParagraphMarkKeys;
  }

  get changedPPrBaseKeys(): readonly AuthoredParagraphPropertyKey[] {
    return this.#changedPPrBaseKeys;
  }

  get changedXmlGroups(): readonly ParagraphPropertySourceXmlGroup[] {
    return this.#changedXmlGroups;
  }

  get isEmpty(): boolean {
    return this.#changedXmlGroups.length === 0;
  }
}

export const paragraphPropertySourceDelta = (
  before: ParagraphPropertySourceFingerprint,
  after: ParagraphPropertySourceFingerprint,
): ParagraphPropertySourceDelta => ParagraphPropertySourceDelta.between(before, after);

const formattingKeys = (
  predicate: (descriptor: ParagraphPropertyDescriptor) => boolean,
): readonly string[] =>
  Object.freeze(
    Object.entries(PARAGRAPH_FORMATTING_PROPERTY_DESCRIPTOR)
      .filter(([, descriptor]) => predicate(descriptor))
      .map(([key]) => key),
  );

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

const PARAGRAPH_FORMATTING_PROPERTY_KEY_SET = new Set<string>(
  PARAGRAPH_FORMATTING_PROPERTY_KEY_LIST,
);

export const isParagraphFormattingPropertyKey = (
  key: string,
): key is keyof typeof PARAGRAPH_FORMATTING_PROPERTY_DESCRIPTOR =>
  PARAGRAPH_FORMATTING_PROPERTY_KEY_SET.has(key);
