/**
 * One carrier for a section break.
 *
 * ECMA-376 Part 1 §17.6.18 defines a `w:sectPr` inside a `w:pPr` as the
 * properties of the section that *ends* at that paragraph's mark, and §17.6.17
 * defines the body's trailing `w:sectPr` as the properties of the last section.
 * There is exactly one record per section, so a paragraph either holds that
 * record or it does not: `_sectionProperties` is the whole state.
 *
 * The break type (`w:type`, §17.6.22) is a field of that record, not a second
 * carrier beside it. It is derived here for every reader, so a reader cannot
 * see a type the record does not state, and a writer cannot state a type no
 * record carries.
 *
 * The record's reference identity is load-bearing: the from-leg tells a copy
 * (two halves of a split paragraph over one object) from two authored breaks
 * (two objects, however equal their contents) by it. Every writer below
 * therefore mints once and shares, and never clones.
 *
 * @packageDocumentation
 */

import type { Node as PMNode } from "prosemirror-model";

import type { SectionProperties } from "../types/document";

/**
 * The section-start values the editor can author and the layout understands.
 *
 * `ST_SectionMark` also admits `nextColumn`, which Folio round-trips but
 * neither offers nor lays out; deriving through this list keeps such a record
 * typeless to a reader rather than handing it a value it cannot render.
 */
export const SECTION_BREAK_TYPES = ["nextPage", "continuous", "oddPage", "evenPage"] as const;

/** A section start the editor can author. */
export type SectionBreakType = (typeof SECTION_BREAK_TYPES)[number];

const isSectionBreakType = (value: unknown): value is SectionBreakType =>
  SECTION_BREAK_TYPES.some((type) => type === value);

/** A break type read off an untrusted boundary (a `data-` attribute, a snapshot). */
export const parseSectionBreakType = (value: unknown): SectionBreakType | null =>
  isSectionBreakType(value) ? value : null;

/** The section record a paragraph holds, or `null` when it ends no section. */
export const sectionPropertiesOf = (paragraph: PMNode): SectionProperties | null =>
  (paragraph.attrs["_sectionProperties"] as SectionProperties | null | undefined) ?? null;

/**
 * The break type a section record states, or `null` when it states none the
 * editor authors. The single derivation every reader goes through.
 */
export const sectionBreakTypeOf = (
  properties: SectionProperties | null | undefined,
): SectionBreakType | null => {
  const sectionStart = properties?.sectionStart;
  return isSectionBreakType(sectionStart) ? sectionStart : null;
};

/**
 * Mint the one record a newly inserted break carries.
 *
 * A minted record states only its type: every other setting of the new section
 * is inherited from the section that follows it, which is what a reader of the
 * package does with a `w:sectPr` that omits them.
 */
export const mintSectionProperties = (breakType: SectionBreakType): SectionProperties => ({
  sectionStart: breakType,
});
