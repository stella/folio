/**
 * Which schema simple type each `narrowEnum` site's value set stands for.
 *
 * `narrowEnum(value, XSchema)` returns `undefined` for a token the picklist
 * omits, and every caller then drops the attribute — so a picklist that has
 * drifted from the enumeration it mirrors is a silent survival loss, one token
 * at a time. `ThemeColorSlotSchema` was such a mirror: it spelled six members
 * the DrawingML way and omitted seven, and nothing compared it to
 * `ST_ThemeColor`.
 *
 * This table binds each site to the type it mirrors so
 * `scripts/narrowed-enum-schema-types.test.ts` can compare them member for
 * member. It is total over the picklists `parserEnums.ts` exports: a new one
 * does not compile until it has a decision, and a `narrowEnum` call naming a
 * schema the table does not carry fails the wiring test.
 *
 * A `diverges` entry records the exact difference, so the difference can shrink
 * but not grow: a schema refresh, or a picklist edit, that widens it fails.
 */

import { THEME_COLORS } from "../../packages/docx-core/src/model/themeColor.gen";
import * as parserEnums from "../../packages/core/src/docx/parserEnums";

/**
 * Every picklist `parserEnums.ts` exports, derived from the module rather than
 * listed: a new one does not compile until the table below decides about it.
 */
export type NarrowedEnumSite = Extract<keyof typeof parserEnums, `${string}Schema`>;

/** Namespace prefixes a simple-type key may use. */
export const NARROWED_ENUM_NAMESPACE_URIS = {
  a: "http://schemas.openxmlformats.org/drawingml/2006/main",
  s: "http://schemas.openxmlformats.org/officeDocument/2006/sharedTypes",
  w: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
  wp: "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
} as const;

/** `"w:ST_Jc"` — prefixed because two vocabularies declare a `ST_Jc`. */
type SchemaSimpleType = string;

export type NarrowedEnumBinding =
  | {
      /** The picklist is the enumeration, member for member. */
      readonly kind: "matches";
      readonly simpleType: SchemaSimpleType;
    }
  | {
      /** The picklist and the enumeration differ, by exactly this much. */
      readonly kind: "diverges";
      readonly simpleType: SchemaSimpleType;
      /** Enumeration members the picklist omits. */
      readonly missing: readonly string[];
      /** Picklist members the enumeration does not declare. */
      readonly extra: readonly string[];
      /** What folio does with the difference, and what would close it. */
      readonly reason: string;
    }
  | {
      /**
       * The picklist names what folio paints; the reader keeps a member outside
       * it verbatim, so nothing is lost by narrowing. The picklist must still be
       * a subset of the enumeration: an invented member would be a real defect.
       */
      readonly kind: "keeps-raw";
      readonly simpleType: SchemaSimpleType;
      /** The reader that keeps the raw token, as `<module>#<name>`. */
      readonly reader: string;
    }
  | {
      /** No schema simple type enumerates this value set. */
      readonly kind: "unschematised";
      readonly reason: string;
    };

const matches = (simpleType: SchemaSimpleType): NarrowedEnumBinding => ({
  kind: "matches",
  simpleType,
});

/**
 * The 21 `a:prstGeom/@prst` presets folio has no geometry for.
 * `hasUnsupportedGeometry` detects exactly this case and the shape is replayed
 * from its captured `rawXml`, so nothing is written back as a rectangle.
 */
const SHAPE_TYPE_MISSING = [
  "chartPlus",
  "chartStar",
  "chartX",
  "cornerTabs",
  "diamond",
  "doubleWave",
  "ellipseRibbon",
  "ellipseRibbon2",
  "flowChartOfflineStorage",
  "horizontalScroll",
  "leftRightRibbon",
  "lineInv",
  "nonIsoscelesTrapezoid",
  "plaqueTabs",
  "plus",
  "ribbon",
  "ribbon2",
  "squareTabs",
  "upDownArrowCallout",
  "verticalScroll",
  "wave",
] as const;

export const NARROWED_ENUM_SCHEMA_TYPES = {
  // 193 members, 166 of them page-border art glyphs. `BorderSpec.style` is a
  // `string` and the reader writes `narrowEnum(...) ?? rawStyle`, so a member
  // outside `KnownBorderStyle` survives a save unpainted.
  BorderStyleSchema: {
    kind: "keeps-raw",
    simpleType: "w:ST_Border",
    reader: "packages/core/src/docx/borderParser.ts#parseBorderSpec",
  },
  ConditionalStyleTypeSchema: matches("w:ST_TblStyleOverrideType"),
  EmphasisMarkSchema: matches("w:ST_Em"),
  FieldTypeSchema: {
    kind: "unschematised",
    reason:
      "A field's kind is the first token of `w:instrText`, which the schema types as free text. The union is folio's own vocabulary of the field types it evaluates; an unrecognised one reads as UNKNOWN and the instruction text round-trips verbatim.",
  },
  FloatingTableXSpecSchema: matches("s:ST_XAlign"),
  FloatingTableYSpecSchema: matches("s:ST_YAlign"),
  FontHintSchema: {
    kind: "diverges",
    simpleType: "w:ST_Hint",
    missing: [],
    extra: ["cs"],
    reason:
      "ECMA-376 §17.18.42 lists `cs` and the Transitional XSD does not. Word writes it, so folio accepts it: the divergence widens what parses rather than narrowing it, and removing `cs` would drop the attribute on a real document.",
  },
  FontThemeSchema: matches("w:ST_Theme"),
  FrameWrapSchema: matches("w:ST_Wrap"),
  FrameXAlignSchema: matches("s:ST_XAlign"),
  FrameYAlignSchema: matches("s:ST_YAlign"),
  HighlightColorSchema: matches("w:ST_HighlightColor"),
  ImageHorizontalAlignmentSchema: matches("wp:ST_AlignH"),
  ImageHorizontalRelativeToSchema: matches("wp:ST_RelFromH"),
  ImageVerticalAlignmentSchema: matches("wp:ST_AlignV"),
  ImageVerticalRelativeToSchema: matches("wp:ST_RelFromV"),
  ImageWrapTextSchema: matches("wp:ST_WrapText"),
  LevelSuffixSchema: matches("w:ST_LevelSuffix"),
  LineSpacingRuleSchema: matches("w:ST_LineSpacingRule"),
  NumberFormatSchema: {
    kind: "diverges",
    simpleType: "w:ST_NumberFormat",
    missing: ["bahtText", "custom", "dollarText"],
    extra: ["decimalZero3", "decimalZero4", "decimalZero5"],
    reason:
      "`narrowEnum` drops the attribute for a member the union omits, so a `w:numFmt` of `bahtText` or `dollarText` loses its format on save; `custom` defers to `@w:format`, which folio does not read either. The three `decimalZeroN` members are not in the enumeration at all. Both halves are a survival loss waiting to be measured, not a decision.",
  },
  ParagraphAlignmentSchema: matches("w:ST_Jc"),
  PositionalTabAlignmentSchema: matches("w:ST_PTabAlignment"),
  PositionalTabLeaderSchema: matches("w:ST_PTabLeader"),
  PositionalTabRelativeToSchema: matches("w:ST_PTabRelativeTo"),
  SdtLockSchema: matches("w:ST_Lock"),
  ShadingPatternSchema: matches("w:ST_Shd"),
  ShapeOutlineStyleSchema: matches("a:ST_PresetLineDashVal"),
  ShapeTypeSchema: {
    kind: "diverges",
    simpleType: "a:ST_ShapeType",
    missing: SHAPE_TYPE_MISSING,
    extra: ["textBox"],
    reason:
      "21 presets folio has no geometry for. `hasUnsupportedGeometry` reads the same picklist and marks the shape for verbatim replay, so the markup survives; the union names the geometries folio can draw. `textBox` is folio's own marker for a `wps:txbx` with no `prstGeom`, which the schema has no token for.",
  },
  StyleTypeSchema: matches("w:ST_StyleType"),
  TabLeaderSchema: matches("w:ST_TabTlc"),
  TabStopAlignmentSchema: matches("w:ST_TabJc"),
  TableCellTextDirectionSchema: matches("w:ST_TextDirection"),
  TableRowHeightRuleSchema: matches("w:ST_HeightRule"),
  TableWidthTypeSchema: matches("w:ST_TblWidth"),
  TextEffectSchema: matches("w:ST_TextEffect"),
  UnderlineStyleSchema: matches("w:ST_Underline"),
} as const satisfies Record<NarrowedEnumSite, NarrowedEnumBinding>;

/**
 * A schema enumeration a registered reserved-value slot reaches, and what the
 * model does with it, where no `narrowEnum` picklist mirrors it.
 *
 * `check-reserved-value-coverage.ts` derives the enumerating types every
 * registered slot can carry and fails on one this table and the picklist
 * bindings above both leave out. Without it the registry could keep a claim
 * about a slot whose model enumeration no longer resembles the schema's — which
 * is exactly what happened to `ST_ThemeColor`, for sixteen members and four
 * releases.
 */
export type SlotEnumerationDecision =
  | {
      /**
       * The model's token set, from a list the schema itself generated. It must
       * equal the enumeration exactly.
       */
      readonly kind: "generated";
      readonly tokens: readonly string[];
      /** Where the list comes from, as `<module>#<name>`. */
      readonly source: string;
    }
  | {
      /** The model does not hold this type's tokens as an enumeration. */
      readonly kind: "excluded";
      readonly reason: string;
    };

/**
 * A lexical space the model decodes rather than carries: the tokens are
 * spellings of one value, not members a consumer chooses between.
 */
const LEXICAL_SPACE =
  "A lexical space, not a vocabulary: `ST_OnOff` is six spellings of a boolean and `ST_HexColor` is six hex digits or `auto`. The model stores the decoded value, so there is no token set to compare; `parseOnOffValue` and `resolveColor` own the decoding.";

/**
 * A union the model declares inline, with no exported token list to compare.
 * Binding one means exporting its members the way `parserEnums.ts` does.
 */
const INLINE_UNION =
  "Read by a literal comparison into a union the model declares inline, so there is no token list to compare against the enumeration. Exporting the members as a picklist, the way `parserEnums.ts` does for the narrowed sites, is what would bind it.";

/** Markup folio preserves rather than models. */
const NOT_MODELLED =
  "Not modelled. The element reaches the model only inside verbatim-captured markup, so no reader compares its value and no token can be dropped.";

export const RESERVED_VALUE_SLOT_ENUMERATIONS = {
  "w:ST_ThemeColor": {
    kind: "generated",
    tokens: THEME_COLORS,
    source: "packages/docx-core/src/model/themeColor.gen.ts#THEME_COLORS",
  },
  "s:ST_OnOff": { kind: "excluded", reason: LEXICAL_SPACE },
  "w:ST_HexColor": { kind: "excluded", reason: LEXICAL_SPACE },
  "w:ST_BrClear": { kind: "excluded", reason: INLINE_UNION },
  "w:ST_DocGrid": { kind: "excluded", reason: INLINE_UNION },
  "w:ST_DropCap": { kind: "excluded", reason: INLINE_UNION },
  "w:ST_HdrFtr": { kind: "excluded", reason: INLINE_UNION },
  "w:ST_Merge": { kind: "excluded", reason: INLINE_UNION },
  "w:ST_PageBorderZOrder": { kind: "excluded", reason: INLINE_UNION },
  "w:ST_SectionMark": { kind: "excluded", reason: INLINE_UNION },
  "w:ST_TblLayoutType": { kind: "excluded", reason: INLINE_UNION },
  "w:ST_TblOverlap": { kind: "excluded", reason: INLINE_UNION },
  "s:ST_VerticalAlignRun": { kind: "excluded", reason: INLINE_UNION },
  "a:ST_LineEndType": { kind: "excluded", reason: NOT_MODELLED },
  "a:ST_TextWrappingType": { kind: "excluded", reason: NOT_MODELLED },
  "w:ST_FFTextType": { kind: "excluded", reason: NOT_MODELLED },
} as const satisfies Readonly<Record<string, SlotEnumerationDecision>>;

/** `"w:ST_Jc"` -> `"{<wml uri>}ST_Jc"`. */
export const expandSimpleType = (qualified: string): string => {
  const colon = qualified.indexOf(":");
  const prefix = qualified.slice(0, colon);
  const uri = (NARROWED_ENUM_NAMESPACE_URIS as Readonly<Record<string, string>>)[prefix];
  if (uri === undefined) {
    throw new Error(`Simple type "${qualified}" uses an undeclared namespace prefix.`);
  }
  return `{${uri}}${qualified.slice(colon + 1)}`;
};

/** Each picklist binding, keyed by the schema type it mirrors. */
export const narrowedEnumSiteByType = (): ReadonlyMap<
  string,
  { site: NarrowedEnumSite; binding: NarrowedEnumBinding }
> => {
  const byType = new Map<string, { site: NarrowedEnumSite; binding: NarrowedEnumBinding }>();
  for (const [site, binding] of Object.entries(NARROWED_ENUM_SCHEMA_TYPES) as [
    NarrowedEnumSite,
    NarrowedEnumBinding,
  ][]) {
    if (binding.kind === "unschematised") {
      continue;
    }
    byType.set(expandSimpleType(binding.simpleType), { site, binding });
  }
  return byType;
};

/** The tokens one picklist accepts, read from the picklist itself. */
export const narrowedEnumOptions = (site: NarrowedEnumSite): readonly string[] => {
  const schema = (parserEnums as Readonly<Record<string, unknown>>)[site];
  const options = (schema as { options?: readonly string[] } | undefined)?.options;
  if (options === undefined) {
    throw new Error(`parserEnums.${site} is not a picklist.`);
  }
  return options;
};
