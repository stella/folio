/**
 * Every disposition map, keyed by the model type it is total over.
 *
 * This is the single place the reserved-value lint and
 * `scripts/check-reserved-value-coverage.ts` read, so neither mirrors the
 * decisions by hand. A model type covered here is gated at compile time by the
 * `satisfies Record<keyof T, ReservedValueDisposition>` on its own map; a model
 * type not listed here is not covered yet, and `docs/reserved-values.md` tracks
 * what is left.
 */

import type { ReservedValueDisposition, ReservedValueMap } from "./disposition";
import {
  COLUMN_RESERVED,
  ENDNOTE_PROPERTIES_RESERVED,
  ENDNOTE_RESERVED,
  FOOTNOTE_PROPERTIES_RESERVED,
  FOOTNOTE_RESERVED,
  SECTION_BACKGROUND_RESERVED,
  SECTION_DOC_GRID_RESERVED,
  SECTION_LINE_NUMBERS_RESERVED,
  SECTION_PAGE_BORDERS_RESERVED,
  SECTION_PAGE_NUMBERING_RESERVED,
  SECTION_PROPERTIES_RESERVED,
} from "./content";
import { BORDER_SPEC_RESERVED, COLOR_VALUE_RESERVED, SHADING_PROPERTIES_RESERVED } from "./colors";
import {
  CELL_MARGINS_RESERVED,
  CONDITIONAL_FORMAT_STYLE_RESERVED,
  FLOATING_TABLE_PROPERTIES_RESERVED,
  FONT_FAMILY_RESERVED,
  PARAGRAPH_FORMATTING_RESERVED,
  PARAGRAPH_FRAME_RESERVED,
  PARAGRAPH_NUMBERING_RESERVED,
  RUN_LANGUAGE_RESERVED,
  SPACING_EXPLICIT_RESERVED,
  TAB_STOP_RESERVED,
  TABLE_BORDERS_RESERVED,
  TABLE_CELL_BORDERS_RESERVED,
  TABLE_CELL_FORMATTING_RESERVED,
  TABLE_FORMATTING_RESERVED,
  TABLE_LOOK_RESERVED,
  TABLE_MEASUREMENT_RESERVED,
  TABLE_ROW_FORMATTING_RESERVED,
  TEXT_FORMATTING_RESERVED,
  UNDERLINE_RESERVED,
} from "./formatting";
import {
  ABSTRACT_NUMBERING_RESERVED,
  LEVEL_OVERRIDE_RESERVED,
  LIST_LEVEL_RESERVED,
  LIST_RENDERING_RESERVED,
  NUMBERING_DEFINITIONS_RESERVED,
  NUMBERING_INSTANCE_RESERVED,
} from "./lists";

export const RESERVED_VALUE_REGISTRY: Readonly<Record<string, ReservedValueMap>> = {
  AbstractNumbering: ABSTRACT_NUMBERING_RESERVED,
  BorderSpec: BORDER_SPEC_RESERVED,
  CellMargins: CELL_MARGINS_RESERVED,
  ColorValue: COLOR_VALUE_RESERVED,
  Column: COLUMN_RESERVED,
  ConditionalFormatStyle: CONDITIONAL_FORMAT_STYLE_RESERVED,
  Endnote: ENDNOTE_RESERVED,
  EndnoteProperties: ENDNOTE_PROPERTIES_RESERVED,
  FloatingTableProperties: FLOATING_TABLE_PROPERTIES_RESERVED,
  Footnote: FOOTNOTE_RESERVED,
  FootnoteProperties: FOOTNOTE_PROPERTIES_RESERVED,
  ListLevel: LIST_LEVEL_RESERVED,
  ListRendering: LIST_RENDERING_RESERVED,
  NumberingDefinitions: NUMBERING_DEFINITIONS_RESERVED,
  NumberingInstance: NUMBERING_INSTANCE_RESERVED,
  "NumberingInstance.levelOverrides": LEVEL_OVERRIDE_RESERVED,
  ParagraphFormatting: PARAGRAPH_FORMATTING_RESERVED,
  "ParagraphFormatting.frame": PARAGRAPH_FRAME_RESERVED,
  "ParagraphFormatting.numPr": PARAGRAPH_NUMBERING_RESERVED,
  SectionProperties: SECTION_PROPERTIES_RESERVED,
  "SectionProperties.background": SECTION_BACKGROUND_RESERVED,
  "SectionProperties.docGrid": SECTION_DOC_GRID_RESERVED,
  "SectionProperties.lineNumbers": SECTION_LINE_NUMBERS_RESERVED,
  "SectionProperties.pageBorders": SECTION_PAGE_BORDERS_RESERVED,
  "SectionProperties.pageNumbering": SECTION_PAGE_NUMBERING_RESERVED,
  ShadingProperties: SHADING_PROPERTIES_RESERVED,
  SpacingExplicit: SPACING_EXPLICIT_RESERVED,
  TabStop: TAB_STOP_RESERVED,
  TableBorders: TABLE_BORDERS_RESERVED,
  TableCellBorders: TABLE_CELL_BORDERS_RESERVED,
  TableCellFormatting: TABLE_CELL_FORMATTING_RESERVED,
  TableFormatting: TABLE_FORMATTING_RESERVED,
  TableLook: TABLE_LOOK_RESERVED,
  TableMeasurement: TABLE_MEASUREMENT_RESERVED,
  TableRowFormatting: TABLE_ROW_FORMATTING_RESERVED,
  TextFormatting: TEXT_FORMATTING_RESERVED,
  "TextFormatting.fontFamily": FONT_FAMILY_RESERVED,
  "TextFormatting.language": RUN_LANGUAGE_RESERVED,
  "TextFormatting.underline": UNDERLINE_RESERVED,
};

/** One field's decision, with the model type and field it was recorded against. */
export type ReservedValueEntry = {
  modelType: string;
  field: string;
  disposition: ReservedValueDisposition;
};

/** Every recorded decision, in registry order. */
export const reservedValueEntries = (): ReservedValueEntry[] => {
  const entries: ReservedValueEntry[] = [];
  for (const [modelType, map] of Object.entries(RESERVED_VALUE_REGISTRY)) {
    for (const [field, disposition] of Object.entries(map)) {
      entries.push({ modelType, field, disposition });
    }
  }
  return entries;
};

/**
 * Slot keys the registry names, split on `|` and de-duplicated.
 *
 * `scripts/check-reserved-value-coverage.ts` resolves these against the
 * committed schema graph; the reserved-value lint never sees them.
 */
export const reservedValueSlotKeys = (): string[] => {
  const slots = new Set<string>();
  for (const { disposition } of reservedValueEntries()) {
    if (disposition === "no-reserved-value") {
      continue;
    }
    for (const slot of disposition.slot.split("|")) {
      slots.add(slot);
    }
  }
  return [...slots].sort();
};
