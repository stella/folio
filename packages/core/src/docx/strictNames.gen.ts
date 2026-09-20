/**
 * GENERATED FILE — do not edit.
 *
 * The WordprocessingML names ECMA-376 Part 1 spells by writing direction and
 * Part 4 spells by physical side, derived from the cited list in
 * `specifications/strict-names/renames.ts` and the committed schema graph by
 * `scripts/generate-strict-names.ts`. Regenerate with:
 *
 *   bun run generate:strict-names
 *
 * A key is the complex type the name belongs to — the element's own type, or
 * the attribute's owner — then the local name, marked `@` for an attribute.
 * The comment above each entry names the types that declare the pair.
 */

/** What folio writes for a name a Strict producer spelled by writing direction. */
export const TRANSITIONAL_NAME_BY_STRICT_NAME = {
  // declared by CT_TblBorders, CT_TcBorders
  "CT_Border end": "right",
  // declared by CT_TblBorders, CT_TcBorders
  "CT_Border start": "left",
  // declared by CT_Ind
  "CT_Ind @end": "right",
  // declared by CT_Ind
  "CT_Ind @endChars": "rightChars",
  // declared by CT_Ind
  "CT_Ind @start": "left",
  // declared by CT_Ind
  "CT_Ind @startChars": "leftChars",
  // declared by CT_TblCellMar, CT_TcMar
  "CT_TblWidth end": "right",
  // declared by CT_TblCellMar, CT_TcMar
  "CT_TblWidth start": "left",
} as const;

/** A name Part 1 renamed, keyed by the complex type it belongs to. */
export type StrictName = keyof typeof TRANSITIONAL_NAME_BY_STRICT_NAME;

/** The same keys as a list, so a consumer can walk them without widening them. */
export const STRICT_NAMES = [
  "CT_Border end",
  "CT_Border start",
  "CT_Ind @end",
  "CT_Ind @endChars",
  "CT_Ind @start",
  "CT_Ind @startChars",
  "CT_TblWidth end",
  "CT_TblWidth start",
] as const satisfies readonly StrictName[];

/** What a Strict producer may have written where folio writes the Transitional name. */
export const STRICT_NAMES_BY_TRANSITIONAL_NAME = {
  "CT_Border left": ["start"],
  "CT_Border right": ["end"],
  "CT_Ind @left": ["start"],
  "CT_Ind @leftChars": ["startChars"],
  "CT_Ind @right": ["end"],
  "CT_Ind @rightChars": ["endChars"],
  "CT_TblWidth left": ["start"],
  "CT_TblWidth right": ["end"],
} as const;

/** A slot a reader must take in either spelling, keyed the same way. */
export type RenamedSlot = keyof typeof STRICT_NAMES_BY_TRANSITIONAL_NAME;
