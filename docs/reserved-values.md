# Reserved values

A _reserved value_ is a value an OOXML slot accepts whose meaning is not the
value itself.

- `<w:numId w:val="0"/>` names no numbering definition. It switches numbering
  off, and on a style it cancels the numbering the style would otherwise
  inherit through `w:basedOn`.
- `<w:outlineLvl w:val="9"/>` means body text, not a tenth heading level. The
  model no longer has a representation for it: see _Designing a sentinel out_.
- `<w:tcW w:w="9360" w:type="auto"/>` carries no width: under `auto` the number
  is meaningless and layout decides.
- `<w:vMerge/>` with no `w:val` is a continuation cell.
- `<w:u w:val="none"/>` cancels an inherited underline; it is not the absence of
  an underline.

Most of them are invisible to the schema. `w:numId`, `w:ilvl`, `w:outlineLvl`
and `w:gridSpan` are all `CT_DecimalNumber` over an unfacetted `xs:integer`, so
no amount of narrowing catches a consumer that reads the number as an ordinary
one.

Every leak folio has had here came from a duplicated reader whose copies
drifted: four shading parsers where one keeps `auto`, two `textFormattingToMarks`
where one checks `"none"`, three table-width resolvers where one reads `auto` as
`dxa`. None came from a single reader that simply forgot. So the contract is not
"check the sentinel"; it is "there is one reader, and the absence of a recorded
decision fails the build".

## The three parts

| Part     | Where                                            | What it guarantees                                                                                                         |
| -------- | ------------------------------------------------ | -------------------------------------------------------------------------------------------------------------------------- |
| Registry | `specifications/reserved-values`                 | Every field of a covered model type has a recorded decision. Adding a field without one fails `bun run typecheck`.         |
| Lint     | `folio-reserved-values/no-bare-reserved-compare` | No new bare comparison against a recorded sentinel outside the module that owns the read. Held to a shrink-only baseline.  |
| Coverage | `bun run check:reserved-value-coverage`          | Every slot the schema graph says can carry a reserved value is named by a registry entry or by an exclusion with a reason. |

The registry is total over the _model_; the coverage check is total over the
_schema_. Neither alone is enough: the model cannot see a slot it does not
reach, and the schema cannot see a sentinel it cannot express.

### Why the registry is not in the package

The maps decide something about `@stll/docx-core`'s types, but nothing that
imports the package needs them: they are read by the compiler, by the lint rule
and by the coverage check. A published package should not pay their inference
cost, and a consumer should not be offered an export it has no use for, so they
live in `specifications/reserved-values` with a tsconfig project of their own,
wired into `bun run typecheck` and budgeted in `scripts/typecheck-budget.json`
like any other project. They import the model types type-only; the dependency
runs one way, tooling to package, and never back.

`ExhaustiveFields` stays in `packages/docx-core/src/model`, because the
paragraph, text and border serializers use it in shipped code. The registry
imports it from there.

## Adding a model field

Add it, and the compiler tells you:

```
Property 'myNewField' is missing in type '{ ... }' but required in type
'Record<keyof ParagraphFormatting, ReservedValueDisposition>'.
```

Open the map for that file (`specifications/reserved-values/formatting.ts` for
`ParagraphFormatting`) and record one of four decisions:

```ts
// The slot has no reserved value: every value it accepts means itself.
myNewField: NO_RESERVED_VALUE,

// One function owns the read; everyone else goes through it.
myNewField: readerOwned({
  slot: "w:myElement@val",
  sentinel: "none",
  reader: RESERVED_VALUE_READERS.runProperties,
  evidence: "my-new-field-none-means-x",
}),

// folio does not model the reserved value, and here is why that is safe.
myNewField: notModelled({
  slot: "w:myElement@val",
  sentinel: "none",
  reason: "…what the value means, and what folio does instead.",
}),

// The model has no representation for the sentinel; the parse boundary maps
// it into a named arm of `carrier` instead. See below.
myNewField: unrepresentable({
  slot: "w:myElement@val",
  sentinel: "none",
  carrier: "MyNewFieldValue",
}),
```

`sentinel` holds the literal token(s) as the markup spells them, separated by
`|`. The lint reads those literals, so `"nil|none"` is data, not prose. Four
spellings name a rule with no literal — `absent`, `both-present`,
`meaningless-under-auto`, `unresolvable-styleid` — and the lint skips them.

## Adding a reader

`reader` is `"<repo-relative module>#<function>"`, taken from
`RESERVED_VALUE_READERS` in `specifications/reserved-values/readers.ts`. The path is part
of the key because a bare name is not unique here: `runParser.ts` and
`styleParser.ts` both declare `parseRunProperties`, and `toProseDoc.ts` and
`markUtils.ts` both declare `textFormattingToMarks`.

The lint exempts the whole module that declares the named function, not the
function alone, because a reader routinely spans a private helper and its
exported entry point. A test asserts that every named reader is a function that
actually exists at that path, so a rename fails `bun test scripts`.

Some slots have no single owner yet: `w:tcW` is read three different ways. Those
entries name the reader that implements the rule correctly today, and every
other site lands in the lint baseline — which is what the baseline is for.

## Designing a sentinel out

A reader is the second-best answer. The best one is a model type the sentinel
cannot inhabit, and then the entry becomes `unrepresentable` and the lint has
nothing left to police: there is no number for a comparison to be bare against.

`w:outlineLvl` is the worked example. It was `outlineLevel?: number`, owned by a
reader the registry named as `headingCollector.ts#collectHeadings` — which only
delegates, so the lint was exempting the wrong module while the rule lived in
`docx/builtInStyles.ts`. It is now

```ts
type OutlineLevel =
  { kind: "bodyText" } | { kind: "heading"; level: 0 | 1 | 2 | 3 | 4 | 5 | 6 | 7 | 8 };
```

so `9` parses to the body-text arm, the nine literal types make a tenth heading
level a compile error, and the registry records the carrier rather than a
reader.

Two consequences are deliberate and worth stating, because both are visible in
the gates:

- **A value outside 0..9 is dropped at the parse boundary.** The Rust projection
  kernel already refuses one outright (`MAXIMUM_OUTLINE_LEVEL` in
  `crates/docx-kernel/src/projection/styles.rs`, `ProjectionError::InvalidDocumentXml`),
  so TypeScript accepting what Rust refuses was itself a divergence. TypeScript
  drops and continues rather than throwing, because a document Word opens must
  still open; making it unrepresentable is what turns "dropped" from a
  convention into a fact. `w:outlineLvl@val=-1` is recorded in the
  container-survival baseline as the value that stops surviving a rebuild.
- **The drop is not yet reported as a `ParseWarning`.** The paragraph body path
  (`parseParagraph` → `parseParagraphProperties`) carries no `ParseContext` at
  all, unlike the border, note and comment readers. Threading one through it is
  its own change; until then the census, not a warning, is the record.

## How the baseline shrinks

`bun run lint` does not run the rule. The repository predates the registry and
still carries a hundred bare comparisons; turning them into errors at once would
only produce a hundred suppressions. Instead:

```sh
bun scripts/reserved-value-baseline.ts              # report, per file
bun run check:reserved-values                       # CI gate
bun scripts/reserved-value-baseline.ts --write-baseline
```

The gate fails when a file's count rises or a clean file gains its first
comparison. It also fails when the count _drops_ without the baseline being
regenerated, so a fix has to lock its own win in. `scripts/reserved-value-baseline.json`
therefore only ever goes down, and a file that loses its last comparison leaves
it for good.

The rule is syntactic. It has no type information, so it matches on the field
name a comparison reads through plus the literal it tests against — both derived
from the registry. That means it cannot see:

- a sentinel held in a local it cannot name (`const v = attrs.vMerge; v === "continue"`);
- a raw attribute string inside a parser (`val === "clear"`), because the
  attribute name is not a model field name — those sites are inside the owning
  module anyway;
- a comparison in a module exempted for one slot but reading another (exempting
  `toProseDoc.ts` for `w:u` also exempts its `w:vMerge` reads);
- which of several fields sharing a name is meant. `type` is `TableMeasurement`,
  `SectionProperties.docGrid`, `HeaderReference` and `ShapeFill` at once, so the
  report names every reader recorded for the name. A handful of receivers that
  never hold a model value (`target`, `source`, `dataset`) are skipped outright,
  because `target.type === "none"` is the undo stack, not a shape with no fill.

## The coverage check

`scripts/check-reserved-value-coverage.ts` walks the committed schema graph
(`specifications/generated/docx-transitional-schema.gen.json`, no network) and
derives candidate slots three ways:

1. an attribute whose simple type enumerates a reserved token (`auto`, `nil`,
   `none`, `clear`, `nothing`, `continue`, `baseline`, `default`, `custom`,
   `off`);
2. an attribute the schema gives a default, so an absent attribute and one
   written with its default mean the same thing;
3. a curated list of numeric sentinels the schema cannot express.

It restricts the walk to the parts folio reads, the way
`generate-strict-value-encodings.ts` does, and fails when a candidate is named by
neither the registry nor `scripts/lib/reserved-value-exclusions.ts`. The
exclusion list is long — the schema declares far more than the editable model
covers — so it is grouped by reason, and an exclusion for a slot that stopped
being a candidate is reported as dead.

Moving a slot out of the exclusion list and into the registry is how it gets
modelled.

## Evidence

A prose-only rule is recorded in `specifications/evidence/records` with graded
confidence, a pinned source, and the fixture that would settle it. `confirmed`
is reserved for what the committed schema graph settles on its own;
`provisional` marks a clause the repository cites but nobody has checked against
the pinned ECMA-376 archive; `reported` marks a rule taken from a secondhand
survey with the clause unverified. A registry entry's `evidence` field names the
record, and a test fails if it names one that does not exist.
