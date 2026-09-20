/** Reserved-value decisions for numbering definitions and computed list rendering. */

import type { ExhaustiveFields } from "../../packages/docx-core/src/model/exhaustiveFields";
import type {
  AbstractNumbering,
  LevelOverride,
  ListLevel,
  ListRendering,
  NumberingDefinitions,
  NumberingInstance,
} from "../../packages/docx-core/src/model/lists";
import {
  NO_RESERVED_VALUE,
  notModelled,
  readerOwned,
  type ReservedValueDisposition,
} from "./disposition";
import { RESERVED_VALUE_READERS } from "./readers";

export const LIST_LEVEL_RESERVED = {
  ilvl: readerOwned({
    slot: "w:lvl@ilvl",
    sentinel: "9",
    reader: RESERVED_VALUE_READERS.numbering,
    evidence: "ilvl-outside-zero-to-eight-names-no-level",
  }),
  start: readerOwned({
    slot: "w:start@val",
    sentinel: "0",
    reader: RESERVED_VALUE_READERS.numbering,
    evidence: "lvl-start-zero-is-a-legal-first-number",
  }),
  numFmt: readerOwned({
    slot: "w:numFmt@val",
    sentinel: "none|bullet|custom",
    reader: RESERVED_VALUE_READERS.numberingLevelMarker,
  }),
  lvlText: NO_RESERVED_VALUE,
  lvlJc: NO_RESERVED_VALUE,
  suffix: readerOwned({
    slot: "w:suff@val",
    sentinel: "nothing",
    reader: RESERVED_VALUE_READERS.numbering,
  }),
  pStyle: readerOwned({
    slot: "w:pStyle@val",
    sentinel: "unresolvable-styleid",
    reader: RESERVED_VALUE_READERS.numbering,
    evidence: "unknown-styleid-falls-back-to-the-default-style",
  }),
  lvlPicBulletId: NO_RESERVED_VALUE,
  lvlTextNull: NO_RESERVED_VALUE,
  tplc: NO_RESERVED_VALUE,
  tentative: NO_RESERVED_VALUE,
  preserved: NO_RESERVED_VALUE,
  preservedAttributes: NO_RESERVED_VALUE,
  pPr: NO_RESERVED_VALUE,
  rPr: NO_RESERVED_VALUE,
  lvlRestart: notModelled({
    slot: "w:lvlRestart@val",
    sentinel: "0",
    reason:
      "`0` means this level never restarts, whatever happens at a higher one. folio parses and re-serializes the value but the counter machinery restarts on the nearest higher level unconditionally, so a `0` level renumbers where Word would not.",
    evidence: "lvlrestart-zero-never-restarts",
  }),
  isLgl: NO_RESERVED_VALUE,
  legacy: NO_RESERVED_VALUE,
} satisfies Record<keyof ListLevel, ReservedValueDisposition>;

export type ExhaustiveListLevelReserved = ExhaustiveFields<
  ListLevel,
  keyof typeof LIST_LEVEL_RESERVED
>;

export const ABSTRACT_NUMBERING_RESERVED = {
  abstractNumId: NO_RESERVED_VALUE,
  multiLevelType: NO_RESERVED_VALUE,
  numStyleLink: readerOwned({
    slot: "w:numStyleLink@val",
    sentinel: "unresolvable-styleid",
    reader: RESERVED_VALUE_READERS.numbering,
    evidence: "unknown-styleid-falls-back-to-the-default-style",
  }),
  styleLink: readerOwned({
    slot: "w:styleLink@val",
    sentinel: "unresolvable-styleid",
    reader: RESERVED_VALUE_READERS.numbering,
    evidence: "unknown-styleid-falls-back-to-the-default-style",
  }),
  levels: NO_RESERVED_VALUE,
  name: NO_RESERVED_VALUE,
  nsid: NO_RESERVED_VALUE,
  tmpl: NO_RESERVED_VALUE,
  preserved: NO_RESERVED_VALUE,
  preservedAttributes: NO_RESERVED_VALUE,
} satisfies Record<keyof AbstractNumbering, ReservedValueDisposition>;

export type ExhaustiveAbstractNumberingReserved = ExhaustiveFields<
  AbstractNumbering,
  keyof typeof ABSTRACT_NUMBERING_RESERVED
>;

export const NUMBERING_INSTANCE_RESERVED = {
  numId: readerOwned({
    slot: "w:num@numId",
    sentinel: "0",
    reader: RESERVED_VALUE_READERS.numberingReference,
    evidence: "numid-zero-is-no-numbering",
  }),
  abstractNumId: NO_RESERVED_VALUE,
  levelOverrides: NO_RESERVED_VALUE,
  preserved: NO_RESERVED_VALUE,
  preservedAttributes: NO_RESERVED_VALUE,
} satisfies Record<keyof NumberingInstance, ReservedValueDisposition>;

export type ExhaustiveNumberingInstanceReserved = ExhaustiveFields<
  NumberingInstance,
  keyof typeof NUMBERING_INSTANCE_RESERVED
>;

export const LEVEL_OVERRIDE_RESERVED = {
  ilvl: readerOwned({
    slot: "w:lvlOverride@ilvl",
    sentinel: "9",
    reader: RESERVED_VALUE_READERS.numbering,
    evidence: "ilvl-outside-zero-to-eight-names-no-level",
  }),
  startOverride: readerOwned({
    slot: "w:startOverride@val",
    sentinel: "0",
    reader: RESERVED_VALUE_READERS.numbering,
    evidence: "lvl-start-zero-is-a-legal-first-number",
  }),
  lvl: NO_RESERVED_VALUE,
  preserved: NO_RESERVED_VALUE,
  preservedAttributes: NO_RESERVED_VALUE,
} satisfies Record<keyof LevelOverride, ReservedValueDisposition>;

export type ExhaustiveLevelOverrideReserved = ExhaustiveFields<
  LevelOverride,
  keyof typeof LEVEL_OVERRIDE_RESERVED
>;

export const NUMBERING_DEFINITIONS_RESERVED = {
  abstractNums: NO_RESERVED_VALUE,
  nums: NO_RESERVED_VALUE,
  preserved: NO_RESERVED_VALUE,
  preservedAttributes: NO_RESERVED_VALUE,
} satisfies Record<keyof NumberingDefinitions, ReservedValueDisposition>;

export type ExhaustiveNumberingDefinitionsReserved = ExhaustiveFields<
  NumberingDefinitions,
  keyof typeof NUMBERING_DEFINITIONS_RESERVED
>;

/**
 * `ListRendering` is computed, not parsed: every field is already the resolved
 * answer, so the reserved values its inputs carry were spent upstream. The two
 * exceptions carry an OOXML value through verbatim.
 */
export const LIST_RENDERING_RESERVED = {
  marker: NO_RESERVED_VALUE,
  markerTemplate: NO_RESERVED_VALUE,
  level: NO_RESERVED_VALUE,
  numId: readerOwned({
    slot: "w:numId@val",
    sentinel: "0",
    reader: RESERVED_VALUE_READERS.numberingReference,
    evidence: "numid-zero-is-no-numbering",
  }),
  isBullet: NO_RESERVED_VALUE,
  isLegal: NO_RESERVED_VALUE,
  numFmt: readerOwned({
    slot: "w:numFmt@val",
    sentinel: "none|bullet|custom",
    reader: RESERVED_VALUE_READERS.numberingLevelMarker,
  }),
  markerHidden: NO_RESERVED_VALUE,
  markerFormatting: NO_RESERVED_VALUE,
  markerAlignment: NO_RESERVED_VALUE,
  markerAllCaps: NO_RESERVED_VALUE,
  markerSuffix: readerOwned({
    slot: "w:suff@val",
    sentinel: "nothing",
    reader: RESERVED_VALUE_READERS.numbering,
  }),
  levelNumFmts: NO_RESERVED_VALUE,
  levelStarts: NO_RESERVED_VALUE,
  abstractNumId: NO_RESERVED_VALUE,
  startOverride: NO_RESERVED_VALUE,
  implicitChildLevelAdvances: NO_RESERVED_VALUE,
  markerSecondSlotOffsetTwips: NO_RESERVED_VALUE,
} satisfies Record<keyof ListRendering, ReservedValueDisposition>;

export type ExhaustiveListRenderingReserved = ExhaustiveFields<
  ListRendering,
  keyof typeof LIST_RENDERING_RESERVED
>;
