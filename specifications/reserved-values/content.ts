/** Reserved-value decisions for section properties and note properties. */

import type {
  Column,
  Endnote,
  EndnoteProperties,
  Footnote,
  FootnoteProperties,
  SectionProperties,
} from "../../packages/docx-core/src/model/content";
import type { ExhaustiveFields } from "../../packages/docx-core/src/model/exhaustiveFields";
import {
  NO_RESERVED_VALUE,
  notModelled,
  readerOwned,
  type ReservedValueDisposition,
  toggle,
} from "./disposition.ts";
import { RESERVED_VALUE_READERS } from "./readers.ts";

export const SECTION_PROPERTIES_RESERVED = {
  pageWidth: NO_RESERVED_VALUE,
  pageHeight: NO_RESERVED_VALUE,
  orientation: NO_RESERVED_VALUE,
  marginTop: NO_RESERVED_VALUE,
  marginBottom: NO_RESERVED_VALUE,
  marginLeft: NO_RESERVED_VALUE,
  marginRight: NO_RESERVED_VALUE,
  headerDistance: NO_RESERVED_VALUE,
  footerDistance: NO_RESERVED_VALUE,
  gutter: NO_RESERVED_VALUE,
  columnCount: NO_RESERVED_VALUE,
  columnSpace: NO_RESERVED_VALUE,
  equalWidth: toggle("w:cols@equalWidth"),
  separator: toggle("w:cols@sep"),
  columns: NO_RESERVED_VALUE,
  sectionStart: readerOwned({
    slot: "w:type@val",
    sentinel: "absent",
    reader: RESERVED_VALUE_READERS.sectionBreak,
    evidence: "sectpr-type-absent-means-nextpage",
  }),
  verticalAlign: NO_RESERVED_VALUE,
  textDirection: NO_RESERVED_VALUE,
  bidi: NO_RESERVED_VALUE,
  headerReferences: NO_RESERVED_VALUE,
  footerReferences: NO_RESERVED_VALUE,
  titlePg: NO_RESERVED_VALUE,
  evenAndOddHeaders: NO_RESERVED_VALUE,
  lineNumbers: NO_RESERVED_VALUE,
  pageNumbering: NO_RESERVED_VALUE,
  pageBorders: NO_RESERVED_VALUE,
  background: NO_RESERVED_VALUE,
  footnotePr: NO_RESERVED_VALUE,
  footnoteColumns: NO_RESERVED_VALUE,
  endnotePr: NO_RESERVED_VALUE,
  docGrid: NO_RESERVED_VALUE,
  paperSrcFirst: NO_RESERVED_VALUE,
  paperSrcOther: NO_RESERVED_VALUE,
  formProtection: NO_RESERVED_VALUE,
  noEndnote: NO_RESERVED_VALUE,
  rtlGutter: NO_RESERVED_VALUE,
  printerSettingsRelationshipId: NO_RESERVED_VALUE,
  propertyChanges: NO_RESERVED_VALUE,
  // Children and attributes replayed as the source wrote them. A reserved
  // value is a spelling the model interprets; these slots interpret nothing.
  preserved: NO_RESERVED_VALUE,
  preservedAttributes: NO_RESERVED_VALUE,
} satisfies Record<keyof SectionProperties, ReservedValueDisposition>;

export type ExhaustiveSectionPropertiesReserved = ExhaustiveFields<
  SectionProperties,
  keyof typeof SECTION_PROPERTIES_RESERVED
>;

export const COLUMN_RESERVED = {
  width: NO_RESERVED_VALUE,
  space: NO_RESERVED_VALUE,
} satisfies Record<keyof Column, ReservedValueDisposition>;

export type ExhaustiveColumnReserved = ExhaustiveFields<Column, keyof typeof COLUMN_RESERVED>;

type SectionLineNumbers = NonNullable<SectionProperties["lineNumbers"]>;

export const SECTION_LINE_NUMBERS_RESERVED = {
  start: NO_RESERVED_VALUE,
  countBy: NO_RESERVED_VALUE,
  distance: NO_RESERVED_VALUE,
  restart: NO_RESERVED_VALUE,
} satisfies Record<keyof SectionLineNumbers, ReservedValueDisposition>;

export type ExhaustiveSectionLineNumbersReserved = ExhaustiveFields<
  SectionLineNumbers,
  keyof typeof SECTION_LINE_NUMBERS_RESERVED
>;

type SectionPageNumbering = NonNullable<SectionProperties["pageNumbering"]>;

export const SECTION_PAGE_NUMBERING_RESERVED = {
  format: NO_RESERVED_VALUE,
  start: NO_RESERVED_VALUE,
  chapterStyle: NO_RESERVED_VALUE,
  chapterSeparator: NO_RESERVED_VALUE,
} satisfies Record<keyof SectionPageNumbering, ReservedValueDisposition>;

export type ExhaustiveSectionPageNumberingReserved = ExhaustiveFields<
  SectionPageNumbering,
  keyof typeof SECTION_PAGE_NUMBERING_RESERVED
>;

type SectionPageBorders = NonNullable<SectionProperties["pageBorders"]>;

export const SECTION_PAGE_BORDERS_RESERVED = {
  top: NO_RESERVED_VALUE,
  bottom: NO_RESERVED_VALUE,
  left: NO_RESERVED_VALUE,
  right: NO_RESERVED_VALUE,
  display: NO_RESERVED_VALUE,
  offsetFrom: NO_RESERVED_VALUE,
  zOrder: notModelled({
    slot: "w:pgBorders@zOrder",
    sentinel: "front",
    reason:
      "`front` is the attribute's XSD default, so an absent `@w:zOrder` and an explicit `front` mean the same thing. folio paints page borders above the page background either way and never consults the field, so `back` renders as `front`.",
  }),
} satisfies Record<keyof SectionPageBorders, ReservedValueDisposition>;

export type ExhaustiveSectionPageBordersReserved = ExhaustiveFields<
  SectionPageBorders,
  keyof typeof SECTION_PAGE_BORDERS_RESERVED
>;

type SectionBackground = NonNullable<SectionProperties["background"]>;

export const SECTION_BACKGROUND_RESERVED = {
  color: NO_RESERVED_VALUE,
  themeColor: NO_RESERVED_VALUE,
  themeTint: NO_RESERVED_VALUE,
  themeShade: NO_RESERVED_VALUE,
} satisfies Record<keyof SectionBackground, ReservedValueDisposition>;

export type ExhaustiveSectionBackgroundReserved = ExhaustiveFields<
  SectionBackground,
  keyof typeof SECTION_BACKGROUND_RESERVED
>;

type SectionDocGrid = NonNullable<SectionProperties["docGrid"]>;

export const SECTION_DOC_GRID_RESERVED = {
  type: notModelled({
    slot: "w:docGrid@type",
    sentinel: "default",
    reason:
      "`default` is a named ST_DocGrid member (no grid), not an unset marker. folio applies the grid only when `linePitch` is positive, so `default` and an absent `w:docGrid` behave alike and `snapToChars` behaves as `lines`.",
  }),
  linePitch: NO_RESERVED_VALUE,
  charSpace: NO_RESERVED_VALUE,
} satisfies Record<keyof SectionDocGrid, ReservedValueDisposition>;

export type ExhaustiveSectionDocGridReserved = ExhaustiveFields<
  SectionDocGrid,
  keyof typeof SECTION_DOC_GRID_RESERVED
>;

/**
 * A note's reserved ids are carried by `noteType`, which folio derives from
 * `@w:type` and never from the id.
 */
const NOTE_RESERVED = {
  type: NO_RESERVED_VALUE,
  id: readerOwned({
    slot: "w:footnote@id|w:endnote@id",
    sentinel: "-1|0",
    reader: RESERVED_VALUE_READERS.noteType,
    evidence: "note-ids-minus-one-and-zero-are-reserved",
  }),
  noteType: readerOwned({
    slot: "w:footnote@type|w:endnote@type",
    sentinel: "separator|continuationSeparator|continuationNotice",
    reader: RESERVED_VALUE_READERS.noteType,
    evidence: "note-ids-minus-one-and-zero-are-reserved",
  }),
  content: NO_RESERVED_VALUE,
} satisfies Record<keyof Footnote, ReservedValueDisposition>;

export const FOOTNOTE_RESERVED = NOTE_RESERVED satisfies Record<
  keyof Footnote,
  ReservedValueDisposition
>;
export const ENDNOTE_RESERVED = NOTE_RESERVED satisfies Record<
  keyof Endnote,
  ReservedValueDisposition
>;

export type ExhaustiveNoteReserved = ExhaustiveFields<Footnote, keyof typeof NOTE_RESERVED> &
  ExhaustiveFields<Endnote, keyof typeof NOTE_RESERVED>;

export const FOOTNOTE_PROPERTIES_RESERVED = {
  position: NO_RESERVED_VALUE,
  numFmt: NO_RESERVED_VALUE,
  numFmtFormat: NO_RESERVED_VALUE,
  numStart: NO_RESERVED_VALUE,
  numRestart: NO_RESERVED_VALUE,
} satisfies Record<keyof FootnoteProperties, ReservedValueDisposition>;

export type ExhaustiveFootnotePropertiesReserved = ExhaustiveFields<
  FootnoteProperties,
  keyof typeof FOOTNOTE_PROPERTIES_RESERVED
>;

export const ENDNOTE_PROPERTIES_RESERVED = {
  position: NO_RESERVED_VALUE,
  numFmt: NO_RESERVED_VALUE,
  numFmtFormat: NO_RESERVED_VALUE,
  numStart: NO_RESERVED_VALUE,
  numRestart: NO_RESERVED_VALUE,
} satisfies Record<keyof EndnoteProperties, ReservedValueDisposition>;

export type ExhaustiveEndnotePropertiesReserved = ExhaustiveFields<
  EndnoteProperties,
  keyof typeof ENDNOTE_PROPERTIES_RESERVED
>;
