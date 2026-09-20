/**
 * The WordprocessingML names ECMA-376 Part 1 spells differently from Part 4.
 *
 * Part 1 names a horizontal edge by the direction writing runs in, so a
 * right-to-left document does not have to call its leading edge `left`. Part 4
 * keeps the physical name and declares both spellings, so a Transitional
 * consumer reads either. folio rebuilds every package as Transitional and its
 * model holds one direction, so every save writes the physical name and every
 * reader has to take the logical one.
 *
 * **Why this list is hand-written and the rest is not.** The repository vendors
 * no Strict schema: `specifications/generated/docx-transitional-schema.gen.json`
 * is built from the Part 2 and Part 4 archives alone
 * (`scripts/generate-ooxml-schema-graph.ts`), and Part 1's archive is a 43 MB
 * cache-only source CI does not fetch. So the axis below is the human half —
 * which of two spellings Part 1 declares — and it is the only human half.
 * `scripts/generate-strict-names.ts` finds every slot a rename applies to in
 * the committed Transitional graph and fails when a rename names a spelling the
 * graph does not declare, so an entry cannot be wrong about where it applies.
 * It can only be wrong about the direction, which is the one thing a citation
 * settles.
 */

/** One name ECMA-376 Part 1 declares, and the Part 4 name for the same slot. */
export type StrictRename = {
  /** The local name ECMA-376 Part 1 declares. */
  strict: string;
  /** The local name ECMA-376 Part 4 declares for the same slot. */
  transitional: string;
  /** Why the two are one slot under two names. */
  reason: string;
  /** The pinned source that settles the direction, by its id in `specifications/sources.json`. */
  source: string;
  /** Where in that source the pair is declared. */
  locator: string;
};

const LOGICAL_EDGE = (edge: "leading" | "trailing", unit: string): string =>
  `The ${edge} edge, in ${unit}. ECMA-376 Part 1 names it by writing direction so a ` +
  "right-to-left document does not call its " +
  `${edge} edge ${edge === "leading" ? "left" : "right"}; Part 4 keeps the physical name and ` +
  "declares both, and folio's model holds one direction, so a save writes the physical name.";

export const STRICT_NAME_RENAMES = [
  {
    strict: "start",
    transitional: "left",
    reason: LOGICAL_EDGE("leading", "twips"),
    source: "ecma-376-part-1",
    locator: "wml.xsd, CT_Ind/@start",
  },
  {
    strict: "end",
    transitional: "right",
    reason: LOGICAL_EDGE("trailing", "twips"),
    source: "ecma-376-part-1",
    locator: "wml.xsd, CT_Ind/@end",
  },
  {
    strict: "startChars",
    transitional: "leftChars",
    reason: LOGICAL_EDGE("leading", "hundredths of a character"),
    source: "ecma-376-part-1",
    locator: "wml.xsd, CT_Ind/@startChars",
  },
  {
    strict: "endChars",
    transitional: "rightChars",
    reason: LOGICAL_EDGE("trailing", "hundredths of a character"),
    source: "ecma-376-part-1",
    locator: "wml.xsd, CT_Ind/@endChars",
  },
] as const satisfies readonly StrictRename[];

/** The rename that renames `strict`, for the tables and prose derived from it. */
export const renameOf = (strict: string): StrictRename => {
  const rename = STRICT_NAME_RENAMES.find((entry) => entry.strict === strict);
  if (rename === undefined) {
    throw new Error(`${strict} is not a name STRICT_NAME_RENAMES declares`);
  }
  return rename;
};
