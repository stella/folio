import type { ParagraphFormatting } from "../../types/document";
import type { ParagraphAttrs } from "./nodes";

type ParagraphAttrDefaults = Partial<{
  readonly [Key in keyof ParagraphAttrs]: ParagraphAttrs[Key] | null | undefined;
}>;

const ownKeys = <const Value extends Readonly<Record<string, unknown>>>(
  value: Value,
): readonly (keyof Value)[] => {
  // SAFETY: Value is constrained to string-keyed records, and Object.keys
  // returns exactly the record's own enumerable string keys.
  return Object.keys(value) as (keyof Value)[];
};

/** Canonical schema defaults for attrs replaced wholesale by a pPr change. */
export const PPR_CHANGE_SCOPED_ATTR_DEFAULTS = Object.freeze({
  styleId: null,
  numPr: null,
  alignment: null,
  kinsoku: null,
  overflowPunctuation: null,
  suppressAutoHyphens: null,
  spaceBefore: null,
  spaceAfter: null,
  lineSpacing: null,
  lineSpacingRule: null,
  lineSpacingExplicit: null,
  snapToGrid: null,
  spacingExplicit: null,
  indentLeft: null,
  indentRight: null,
  indentFirstLine: null,
  hangingIndent: false,
  borders: null,
  shading: null,
  tabs: null,
  pageBreakBefore: null,
  keepNext: null,
  keepLines: null,
  widowControl: null,
  contextualSpacing: null,
  outlineLevel: null,
  direction: null,
  _autospacingBase: null,
} as const satisfies ParagraphAttrDefaults);

export const PPR_CHANGE_SCOPED_ATTR_KEYS = Object.freeze(ownKeys(PPR_CHANGE_SCOPED_ATTR_DEFAULTS));

/** Canonical schema defaults for derived list-rendering state. */
export const LIST_RENDERING_ATTR_DEFAULTS = Object.freeze({
  listIsBullet: null,
  listIsLegal: null,
  listNumFmt: null,
  listMarker: null,
  listMarkerTemplate: undefined,
  listMarkerHidden: null,
  listMarkerFormatting: null,
  listMarkerAlignment: null,
  listMarkerSuffix: null,
  listMarkerAllCaps: null,
  listImplicitChildLevelAdvances: null,
  listMarkerSecondSlotOffsetTwips: null,
  listLevelNumFmts: null,
  listLevelStarts: null,
  listAbstractNumId: null,
  listStartOverride: null,
} as const satisfies ParagraphAttrDefaults);

export const LIST_RENDERING_ATTR_KEYS = Object.freeze(ownKeys(LIST_RENDERING_ATTR_DEFAULTS));

type ParagraphFormattingFieldDisposition =
  | { readonly type: "attr"; readonly attr: keyof typeof PPR_CHANGE_SCOPED_ATTR_DEFAULTS }
  | {
      readonly type: "mapped";
      readonly attrs: readonly (keyof typeof PPR_CHANGE_SCOPED_ATTR_DEFAULTS)[];
    }
  | { readonly type: "original-only" }
  | { readonly type: "preserved-live-original" }
  | { readonly type: "outside-change-scope" };

/**
 * Total ownership map for every modeled paragraph-formatting field. Adding a
 * field requires deciding how a pPr change transports it before TypeScript
 * accepts the model extension.
 */
export const PPR_FORMATTING_FIELD_DISPOSITIONS = Object.freeze({
  alignment: { type: "attr", attr: "alignment" },
  bidi: { type: "mapped", attrs: ["direction"] },
  kinsoku: { type: "attr", attr: "kinsoku" },
  overflowPunctuation: { type: "attr", attr: "overflowPunctuation" },
  spaceBefore: { type: "attr", attr: "spaceBefore" },
  spaceAfter: { type: "attr", attr: "spaceAfter" },
  lineSpacing: { type: "mapped", attrs: ["lineSpacing", "lineSpacingExplicit"] },
  lineSpacingRule: { type: "mapped", attrs: ["lineSpacingRule", "lineSpacingExplicit"] },
  snapToGrid: { type: "attr", attr: "snapToGrid" },
  beforeAutospacing: { type: "mapped", attrs: ["_autospacingBase"] },
  afterAutospacing: { type: "mapped", attrs: ["_autospacingBase"] },
  spacingExplicit: { type: "attr", attr: "spacingExplicit" },
  indentLeft: { type: "attr", attr: "indentLeft" },
  indentRight: { type: "attr", attr: "indentRight" },
  indentFirstLine: { type: "attr", attr: "indentFirstLine" },
  hangingIndent: { type: "attr", attr: "hangingIndent" },
  borders: { type: "attr", attr: "borders" },
  shading: { type: "attr", attr: "shading" },
  tabs: { type: "attr", attr: "tabs" },
  keepNext: { type: "attr", attr: "keepNext" },
  keepLines: { type: "attr", attr: "keepLines" },
  widowControl: { type: "attr", attr: "widowControl" },
  pageBreakBefore: { type: "attr", attr: "pageBreakBefore" },
  contextualSpacing: { type: "attr", attr: "contextualSpacing" },
  numPr: { type: "attr", attr: "numPr" },
  numPrFromStyle: { type: "outside-change-scope" },
  outlineLevel: { type: "attr", attr: "outlineLevel" },
  styleId: { type: "attr", attr: "styleId" },
  frame: { type: "original-only" },
  suppressLineNumbers: { type: "original-only" },
  suppressAutoHyphens: { type: "attr", attr: "suppressAutoHyphens" },
  runProperties: { type: "preserved-live-original" },
  runInWithNext: { type: "preserved-live-original" },
} as const satisfies Record<keyof ParagraphFormatting, ParagraphFormattingFieldDisposition>);

export const PPR_CHANGE_SCOPED_FORMATTING_KEYS = Object.freeze(
  ownKeys(PPR_FORMATTING_FIELD_DISPOSITIONS).filter((key) => {
    const type = PPR_FORMATTING_FIELD_DISPOSITIONS[key].type;
    return type !== "outside-change-scope" && type !== "preserved-live-original";
  }),
);

export const PPR_PARSER_ONLY_FORMATTING_KEYS = Object.freeze(
  ownKeys(PPR_FORMATTING_FIELD_DISPOSITIONS).filter(
    (key) => PPR_FORMATTING_FIELD_DISPOSITIONS[key].type !== "attr",
  ),
);

export const PPR_ORIGINAL_ONLY_FORMATTING_KEYS = Object.freeze(
  ownKeys(PPR_FORMATTING_FIELD_DISPOSITIONS).filter(
    (key) => PPR_FORMATTING_FIELD_DISPOSITIONS[key].type === "original-only",
  ),
);

export const PPR_PRESERVED_LIVE_FORMATTING_KEYS = Object.freeze(
  ownKeys(PPR_FORMATTING_FIELD_DISPOSITIONS).filter(
    (key) => PPR_FORMATTING_FIELD_DISPOSITIONS[key].type === "preserved-live-original",
  ),
);
