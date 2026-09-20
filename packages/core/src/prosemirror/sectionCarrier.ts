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

import type { SectionProperties, SectionStart } from "../types/document";

/**
 * `ST_SectionMark` in full (§17.18.77). The list is total over the model union:
 * `sectionBreakTypeOf` returns a record's `SectionStart` as a
 * `SectionBreakType`, so a member missing here fails to compile rather than
 * going typeless to every reader. Which members the *insert* commands offer is
 * a separate, narrower question (`InsertableSectionBreak`).
 */
export const SECTION_BREAK_TYPES = [
  "nextPage",
  "nextColumn",
  "continuous",
  "evenPage",
  "oddPage",
] as const satisfies readonly SectionStart[];

/** A section start a `w:sectPr` can state. */
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
 * The break type a section record states, or `null` when it states none.
 *
 * The single derivation every reader goes through. Every producer of a record
 * already validated `w:type` against the enumeration (the parser, the insert
 * commands, the Yjs migration's `parseSectionBreakType`), so the field needs no
 * second check here; declaring the narrower return type is what proves
 * `SECTION_BREAK_TYPES` still covers the model union.
 */
export const sectionBreakTypeOf = (
  properties: SectionProperties | null | undefined,
): SectionBreakType | null => properties?.sectionStart ?? null;

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
