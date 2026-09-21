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
 * **The renames are generated; the omissions are not.** A rename is a schema
 * fact: `packages/core/src/docx/strictNames.gen.ts` carries every name
 * ECMA-376 Part 1 spells by writing direction, derived from the cited list in
 * `specifications/strict-names/renames.ts` and the committed graph, and the
 * entries below are built from it — so the law, the parsers and the serializers
 * cannot disagree about which two names are one slot. What stays hand-written
 * is what no schema states: the three defaults are prose the XSD cannot express
 * (`CT_OnOff/@w:val`, `CT_VMerge/@w:val` and `CT_TabStop/@w:leader` all carry
 * `default: undefined` in the committed graph), and so is the line where folio
 * performs each canonicalisation, which {@link RENAME_SITES} holds total over
 * the generated table: a rename nobody has decided about cannot land.
 *
 * The guard runs the other way too: every entry must be exercised by a pair the
 * census generates (`container-survival-canonical.test.ts`), which makes the
 * table shrink-only, and each entry carries the rule it rests on and the line
 * that performs it, because the one failure mode that matters is a wrong entry
 * reading a real loss as a survival.
 */

import {
  STRICT_NAMES,
  type StrictName,
  TRANSITIONAL_NAME_BY_STRICT_NAME,
} from "../../../packages/core/src/docx/strictNames.gen";
import { renameOf } from "../../../specifications/strict-names/renames";
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
const ON_OFF_SERIALIZER = "packages/docx-core/src/serialize/xml.ts";

/**
 * Where folio writes the Transitional spelling of each renamed name.
 *
 * Total over the generated table, so a rename a schema refresh adds cannot land
 * without somebody saying whether folio writes it. `null` is that decision, not
 * an omission: folio writes neither spelling of the slot, the authored name is
 * a real loss, and an entry here would turn it into a reported survival. Each
 * citation reads the line that binds the writer's name to this very table key,
 * so the two cannot drift apart without the citation going stale.
 */
const RENAME_SITES = {
  "CT_Border start": [
    {
      file: TABLE_SERIALIZER,
      line: 223,
      writes: 'TRANSITIONAL_NAME_BY_STRICT_NAME["CT_Border start"]',
    },
  ],
  "CT_Border end": [
    {
      file: TABLE_SERIALIZER,
      line: 224,
      writes: 'TRANSITIONAL_NAME_BY_STRICT_NAME["CT_Border end"]',
    },
  ],
  "CT_TblWidth start": [
    {
      file: TABLE_SERIALIZER,
      line: 286,
      writes: 'TRANSITIONAL_NAME_BY_STRICT_NAME["CT_TblWidth start"]',
    },
  ],
  "CT_TblWidth end": [
    {
      file: TABLE_SERIALIZER,
      line: 287,
      writes: 'TRANSITIONAL_NAME_BY_STRICT_NAME["CT_TblWidth end"]',
    },
  ],
  "CT_Ind @start": [
    {
      file: PARAGRAPH_FORMATTING,
      line: 216,
      writes: 'TRANSITIONAL_NAME_BY_STRICT_NAME["CT_Ind @start"]',
    },
  ],
  "CT_Ind @end": [
    {
      file: PARAGRAPH_FORMATTING,
      line: 217,
      writes: 'TRANSITIONAL_NAME_BY_STRICT_NAME["CT_Ind @end"]',
    },
  ],
  // An indent in character units has no typed field. The property-element
  // remainder preserves its authored spelling, so no rename equivalence is
  // needed to observe its survival.
  "CT_Ind @startChars": null,
  "CT_Ind @endChars": null,
} as const satisfies Record<StrictName, readonly [Citation, ...Citation[]] | null>;

/** One generated rename, as the probe reads it. */
const renameSpelling = (key: StrictName): CanonicalSpelling | undefined => {
  const writtenBy = RENAME_SITES[key];
  if (writtenBy === null) {
    return undefined;
  }
  const separator = key.indexOf(" ");
  const type = key.slice(0, separator);
  const name = key.slice(separator + 1);
  const strict = name.startsWith("@") ? name.slice(1) : name;
  const canonical = `w:${TRANSITIONAL_NAME_BY_STRICT_NAME[key]}`;
  const { reason } = renameOf(strict);
  return name.startsWith("@")
    ? { kind: "attribute", type, authored: `w:${strict}`, canonical, reason, writtenBy }
    : { kind: "element", type, authored: `w:${strict}`, canonical, reason, writtenBy };
};

const RENAME_SPELLINGS: readonly CanonicalSpelling[] = STRICT_NAMES.map(renameSpelling).filter(
  (entry): entry is CanonicalSpelling => entry !== undefined,
);

/**
 * The attributes folio writes by leaving out: prose the XSD cannot express.
 *
 * Hand-written, one spec rule and one citation each, because no schema states
 * them. They are omissions rather than renames — nothing is respelled, the
 * attribute is absent from folio's output and the format says an absent one
 * means exactly what the fixture wrote — so they arrive by their own route.
 */
const OMITTED_ATTRIBUTE_SPELLINGS = [
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
      {
        file: ON_OFF_SERIALIZER,
        line: 32,
        writes: "return value ? `<w:${name}/>`",
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
    writtenBy: [{ file: TABLE_SERIALIZER, line: 783, writes: '"<w:vMerge/>"' }],
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
    writtenBy: [{ file: PARAGRAPH_FORMATTING, line: 156, writes: 'leader !== "none"' }],
  },
] as const satisfies readonly CanonicalSpelling[];

/**
 * Every equivalence the probe may use: the generated renames, then the
 * documented omissions. A rename is a schema fact and an omission is prose, so
 * they arrive by different routes and meet here.
 */
export const CANONICAL_SPELLINGS: readonly CanonicalSpelling[] = [
  ...RENAME_SPELLINGS,
  ...OMITTED_ATTRIBUTE_SPELLINGS,
];

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
