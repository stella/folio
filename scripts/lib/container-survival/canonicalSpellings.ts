/**
 * The spellings folio canonicalises, and what each one comes back as.
 *
 * The survival law matches a subject by spelling: the occurrence walk looks for
 * `<w:start`, `attributeIn` for ` w:start="`. folio rebuilds every package as
 * Transitional and writes one spelling of each equivalence below, so a fixture
 * that authored the other reads as lost although nothing was lost — the law's
 * equality is narrower than folio's own output form. This table is that list of
 * equivalences, and the probe accepts a canonical spelling **only** for a
 * subject an entry names. It is not leniency: a subject with no entry is still
 * matched by its authored spelling alone.
 *
 * **It is total against nothing, and cannot be.** There is no schema-declared
 * class of "names folio canonicalises" to be total over: the Transitional
 * schema declares `w:start` and `w:left` as two ordinary optional siblings of
 * `CT_TblBorders`, and the three defaults below are prose the XSD cannot
 * express — `CT_OnOff/@w:val`, `CT_VMerge/@w:val` and `CT_TabStop/@w:leader`
 * all carry `default: undefined` in the committed graph. So the guard runs the
 * other way: every entry must be exercised by a pair the census generates
 * (`container-survival-canonical.test.ts`), which makes the table shrink-only,
 * and each entry carries the rule it rests on and the line that performs it,
 * because the one failure mode that matters is a wrong entry reading a real
 * loss as a survival.
 */

import { spell } from "./fixture";
import { type QualifiedName, qualify, WML_NAMESPACE } from "./schemaSpace";

/**
 * Where folio performs a canonicalisation.
 *
 * The text is part of the citation: a line number alone rots the moment
 * something above it moves, and a rotted citation for a table whose entries
 * cannot be derived is how a wrong equivalence survives review.
 */
type Citation = {
  /** Repository-relative path of the serializer. */
  file: string;
  line: number;
  /** What that line reads, checked by the test that reads the file. */
  writes: string;
};

type Canonicalisation = {
  /** The rule the equivalence rests on, with its citation. */
  reason: string;
  writtenBy: readonly [Citation, ...Citation[]];
};

/**
 * One documented equivalence between what a document may say and what folio writes.
 *
 * `element` and `attribute` are renames: the same thing under two names, and
 * folio writes the Transitional one. `value` is an omission: the attribute is
 * absent from folio's output and the format says an absent one means exactly
 * what the fixture wrote. The three kinds reach different halves of the probe,
 * so they are separated here rather than sharing a field that means a different
 * thing in each.
 */
export type CanonicalSpelling =
  | (Canonicalisation & {
      kind: "element";
      /**
       * The complex type the element is declared with, local name, in
       * WordprocessingML. A name alone is ambiguous: `w:lvl/w:start` is a
       * `CT_DecimalNumber` and means a list's first number.
       */
      type: string;
      /** What a document may write. */
      authored: string;
      /** What folio writes instead. */
      canonical: string;
    })
  | (Canonicalisation & {
      kind: "attribute";
      /** The complex type of the element carrying the attribute. */
      type: string;
      authored: string;
      canonical: string;
    })
  | (Canonicalisation & {
      kind: "value";
      /** The complex type of the element carrying the attribute. */
      type: string;
      /** The attribute whose value is omitted. */
      attribute: string;
      /** The values folio writes by leaving the attribute out. */
      authored: readonly [string, ...string[]];
      canonical: "absent";
    });

const TABLE_SERIALIZER = "packages/core/src/docx/serializer/tableSerializer.ts";
const PARAGRAPH_FORMATTING = "packages/core/src/internal/paragraphFormattingSerialization.ts";

/**
 * `w:start`/`w:end` are the logical-direction names for `w:left`/`w:right`.
 *
 * ECMA-376 Part 4's Transitional schema declares both pairs as optional
 * siblings of the same type — `CT_TblBorders`, `CT_TcBorders` for `CT_Border`,
 * `CT_TblCellMar`, `CT_TcMar` for `CT_TblWidth`, verified in
 * `specifications/generated/docx-transitional-schema.gen.json` — because Part 1
 * gives a bidirectional document the logical spelling and the Transitional
 * profile keeps the physical one. folio's readers take either
 * (`tableParser.ts` and `styleParser.ts` both fall back from `left` to `start`)
 * and its model holds one direction, so every save writes the physical name.
 */
const DIRECTION = (element: "start" | "end"): string =>
  `w:${element} is the logical-direction spelling of w:${element === "start" ? "left" : "right"}: ` +
  "the Transitional schema declares both as optional siblings of the same type, folio's readers " +
  "take either, and its model holds one direction, so a save writes the physical name.";

export const CANONICAL_SPELLINGS = [
  {
    kind: "element",
    type: "CT_Border",
    authored: "w:start",
    canonical: "w:left",
    reason: DIRECTION("start"),
    writtenBy: [
      { file: TABLE_SERIALIZER, line: 196, writes: 'appendBorder(borders.left, "left")' },
    ],
  },
  {
    kind: "element",
    type: "CT_Border",
    authored: "w:end",
    canonical: "w:right",
    reason: DIRECTION("end"),
    writtenBy: [
      { file: TABLE_SERIALIZER, line: 198, writes: 'appendBorder(borders.right, "right")' },
    ],
  },
  {
    kind: "element",
    type: "CT_TblWidth",
    authored: "w:start",
    canonical: "w:left",
    reason: DIRECTION("start"),
    writtenBy: [
      { file: TABLE_SERIALIZER, line: 256, writes: 'serializeMeasurement(margins.left, "left")' },
    ],
  },
  {
    kind: "element",
    type: "CT_TblWidth",
    authored: "w:end",
    canonical: "w:right",
    reason: DIRECTION("end"),
    writtenBy: [
      { file: TABLE_SERIALIZER, line: 264, writes: 'serializeMeasurement(margins.right, "right")' },
    ],
  },
  {
    kind: "attribute",
    type: "CT_Ind",
    authored: "w:start",
    canonical: "w:left",
    reason: DIRECTION("start"),
    writtenBy: [{ file: PARAGRAPH_FORMATTING, line: 257, writes: 'w:left="${intAttr(' }],
  },
  {
    kind: "attribute",
    type: "CT_Ind",
    authored: "w:end",
    canonical: "w:right",
    reason: DIRECTION("end"),
    writtenBy: [{ file: PARAGRAPH_FORMATTING, line: 260, writes: 'w:right="${intAttr(' }],
  },
  {
    kind: "value",
    type: "CT_OnOff",
    attribute: "w:val",
    authored: ["true", "1", "on"],
    canonical: "absent",
    reason:
      "A CT_OnOff element with no w:val is on: the attribute is optional and the committed graph " +
      "records no default for it, so the bare element is the on state, which is the form folio " +
      'writes. The off state keeps the attribute (w:val="0"), which is why only the three on ' +
      "spellings are listed here; the distinction between an explicit off and an absent element " +
      "is the one the evidence record `toggle-property-xor` rests on.",
    writtenBy: [
      { file: PARAGRAPH_FORMATTING, line: 169, writes: "return `<w:${name}/>`;" },
      {
        file: "packages/core/src/docx/serializer/textFormattingSerializer.ts",
        line: 285,
        writes: 'parts.push("<w:b/>")',
      },
    ],
  },
  {
    kind: "value",
    type: "CT_VMerge",
    attribute: "w:val",
    authored: ["continue"],
    canonical: "absent",
    reason:
      "A w:vMerge with no w:val is a continuation cell: the attribute is optional, the committed " +
      "graph records no default, and the omitted value means continue. Recorded as the evidence " +
      "record `vmerge-absent-means-continue`, and as the reserved value `absent|continue` in " +
      "`specifications/reserved-values/formatting.ts`.",
    writtenBy: [{ file: TABLE_SERIALIZER, line: 745, writes: 'parts.push("<w:vMerge/>")' }],
  },
  {
    kind: "value",
    type: "CT_TabStop",
    attribute: "w:leader",
    authored: ["none"],
    canonical: "absent",
    reason:
      "A tab stop with no w:leader draws no leader: the attribute is optional, the committed " +
      "graph records no default, and `none` is the member of ST_TabTlc that says so. folio " +
      "reads the token and writes the attribute only for a leader that draws something.",
    writtenBy: [{ file: PARAGRAPH_FORMATTING, line: 204, writes: 'leader !== "none"' }],
  },
] as const satisfies readonly CanonicalSpelling[];

/** A subject as the probe holds it: the element it is or sits on, and its type. */
export type CanonicalSubject = {
  /** The child element under test, or the element an attribute sits on. */
  element: QualifiedName;
  /** That element's complex type, qualified. */
  type: string | undefined;
  /** The attribute under test, for an attribute subject. */
  attribute: QualifiedName | undefined;
  /** The value the fixture wrote, for an attribute subject. */
  value: string | undefined;
};

/**
 * What folio may spell this subject as, beyond what the fixture authored.
 *
 * Every field is empty for a subject no entry names, which is every subject but
 * the handful the table lists.
 */
export type CanonicalForms = {
  /** The element spelling folio writes instead of the authored one. */
  element: string | undefined;
  /** The attribute spelling folio writes instead of the authored one. */
  attribute: string | undefined;
  /** Whether an entry says folio writes the authored value by omitting the attribute. */
  absentAttribute: boolean;
};

const NO_CANONICAL_FORMS: CanonicalForms = {
  element: undefined,
  attribute: undefined,
  absentAttribute: false,
};

/** Every entry the table declares over the WordprocessingML type the subject sits on. */
const onType = (subject: CanonicalSubject): readonly CanonicalSpelling[] => {
  if (subject.type === undefined) {
    return [];
  }
  const type = subject.type;
  return CANONICAL_SPELLINGS.filter(
    (entry) => qualify({ namespace: WML_NAMESPACE, name: entry.type }) === type,
  );
};

/**
 * The entries that apply to one subject.
 *
 * Exported for the test that holds the table shrink-only: an entry no pair in
 * the census's space matches is one nobody can check, and an unchecked
 * equivalence is exactly the thing that turns a real loss into a reported
 * survival.
 */
export const canonicalSpellingsFor = (subject: CanonicalSubject): readonly CanonicalSpelling[] => {
  const element = spell(subject.element);
  const attribute = subject.attribute === undefined ? undefined : spell(subject.attribute);
  return onType(subject).filter((entry) => {
    switch (entry.kind) {
      // An element entry applies to the element itself and to a subject that
      // sits on it: the probe finds an attribute by walking to its element
      // first, and that step is spelled the same either way.
      case "element":
        return entry.authored === element;
      case "attribute":
        return entry.authored === attribute;
      case "value":
        return (
          entry.attribute === attribute &&
          subject.value !== undefined &&
          entry.authored.includes(subject.value)
        );
    }
  });
};

export const canonicalFormsOf = (subject: CanonicalSubject): CanonicalForms => {
  const forms = { ...NO_CANONICAL_FORMS };
  for (const entry of canonicalSpellingsFor(subject)) {
    switch (entry.kind) {
      case "element":
        forms.element = entry.canonical;
        break;
      case "attribute":
        forms.attribute = entry.canonical;
        break;
      case "value":
        forms.absentAttribute = true;
        break;
    }
  }
  return forms;
};
