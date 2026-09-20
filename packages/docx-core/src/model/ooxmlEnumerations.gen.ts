/**
 * GENERATED FILE — do not edit.
 *
 * The model enumerations that mirror an OOXML simple type, member for
 * member, so a token the format declares cannot go unspelled.
 *
 * Derived from `specifications/generated/docx-transitional-schema.gen.json`.
 * Regenerate with:
 *
 *   bun run generate:ooxml-enumerations
 */

/**
 * `ST_Jc`: every token a `w:jc/@w:val` may carry.
 *
 * `start` and `end` are the direction-aware alignments, distinct from
 * `left` and `right`: in a right-to-left paragraph `start` sits at the
 * right margin. They are members in their own right, not spellings of
 * `left` and `right`, and the layout resolves them against the
 * paragraph's direction. `numTab` aligns to the list number's tab stop.
 */
export const PARAGRAPH_ALIGNMENTS = [
  "start",
  "center",
  "end",
  "both",
  "mediumKashida",
  "distribute",
  "numTab",
  "highKashida",
  "lowKashida",
  "thaiDistribute",
  "left",
  "right",
] as const;

export type ParagraphAlignment =
  | "start"
  | "center"
  | "end"
  | "both"
  | "mediumKashida"
  | "distribute"
  | "numTab"
  | "highKashida"
  | "lowKashida"
  | "thaiDistribute"
  | "left"
  | "right";

/**
 * `ST_TabJc`: every token a `w:tab/@w:val` may carry.
 *
 * `start` and `end` are the direction-aware members, the same distinction
 * `ST_Jc` draws; `clear` removes an inherited stop rather than declaring
 * one, and `num` is the stop a numbered paragraph's text hangs from.
 */
export const TAB_STOP_ALIGNMENTS = [
  "clear",
  "start",
  "center",
  "end",
  "decimal",
  "bar",
  "num",
  "left",
  "right",
] as const;

export type TabStopAlignment =
  | "clear"
  | "start"
  | "center"
  | "end"
  | "decimal"
  | "bar"
  | "num"
  | "left"
  | "right";
