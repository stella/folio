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

import type * as parserEnums from "../../packages/core/src/docx/parserEnums";

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
  ParagraphAlignmentSchema: {
    kind: "diverges",
    simpleType: "w:ST_Jc",
    missing: ["end", "numTab", "start"],
    extra: [],
    reason:
      '`start` and `end` are the Strict logical-direction spellings of `left` and `right`; a `<w:jc w:val="start"/>` fails to narrow and, with no `w:pPr` child dispatcher, takes the whole property set with it. Closed by the paragraph-property dispatcher, which keeps a refused value in the set\'s sink.',
  },
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
  TabStopAlignmentSchema: {
    kind: "diverges",
    simpleType: "w:ST_TabJc",
    missing: ["end", "start"],
    extra: [],
    reason:
      'The same Strict logical-direction spellings as `w:jc`. A `<w:tab w:val="start"/>` loses its alignment; the survival census records both as value-level replay-only losses.',
  },
  TableCellTextDirectionSchema: {
    kind: "diverges",
    simpleType: "w:ST_TextDirection",
    missing: ["lrTb", "lrTbV", "tbLrV"],
    extra: [],
    reason:
      "`lrTb` is the default flow and the other two are vertical variants Word writes for East Asian layout. `narrowEnum` drops the attribute, so a cell written with one reads as the container's direction and saves without it.",
  },
  TableRowHeightRuleSchema: matches("w:ST_HeightRule"),
  TableWidthTypeSchema: matches("w:ST_TblWidth"),
  TextEffectSchema: matches("w:ST_TextEffect"),
  UnderlineStyleSchema: matches("w:ST_Underline"),
} as const satisfies Record<NarrowedEnumSite, NarrowedEnumBinding>;
