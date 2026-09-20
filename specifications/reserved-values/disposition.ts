/**
 * What a model field decides about the reserved value of the OOXML slot it carries.
 *
 * A *reserved value* is a value an OOXML slot accepts whose meaning is not the
 * value itself: `w:numId` `0` names no numbering definition, `w:outlineLvl` `9`
 * means body text rather than a tenth heading level, a `w:tcW` number is
 * meaningless under `w:type="auto"`. Most of them are invisible to the schema —
 * `w:numId`, `w:ilvl`, `w:outlineLvl` and `w:gridSpan` are all unfacetted
 * `CT_DecimalNumber` — so no amount of narrowing catches a consumer that reads
 * the number as an ordinary one.
 *
 * Every leak folio has had here came from a duplicated reader whose copies
 * drifted, never from one reader that forgot. So each field records its
 * decision once, next to the type that declares it, and every disposition map
 * is total over its type by construction: pair it with an `ExhaustiveFields`
 * alias and a field added to the model without a decision fails
 * `bun run typecheck`.
 *
 * The maps are also what the reserved-value lint and
 * `scripts/check-reserved-value-coverage.ts` read, so nothing mirrors them by
 * hand.
 */

import { RESERVED_VALUE_READERS } from "./readers";

/**
 * Namespace prefixes a {@link ReservedValueDisposition} slot may use.
 *
 * A slot key is a registry-internal alias, not a document prefix: folio resolves
 * document markup by namespace URI plus local name, and this table is the one
 * place the alias binds to a URI.
 */
export const RESERVED_VALUE_NAMESPACE_URIS = {
  a: "http://schemas.openxmlformats.org/drawingml/2006/main",
  lc: "http://schemas.openxmlformats.org/drawingml/2006/lockedCanvas",
  m: "http://schemas.openxmlformats.org/officeDocument/2006/math",
  pic: "http://schemas.openxmlformats.org/drawingml/2006/picture",
  w: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
  wp: "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
} as const;

/** A prefix the registry may spell a slot with. */
export type ReservedValueNamespacePrefix = keyof typeof RESERVED_VALUE_NAMESPACE_URIS;

/**
 * The slots one field covers, as `"<prefix>:<element>@<attribute>"`, several
 * separated by `|`.
 *
 * The attribute carries no prefix because a WordprocessingML attribute is
 * declared in no namespace; the element's prefix identifies the vocabulary. One
 * field often covers a family: `ColorValue.auto` is the same `ST_HexColor`
 * decision under `w:color`, `w:shd` and every border side.
 *
 * Plain `string` rather than a template-literal type: a template only
 * constrains the first alternative, and the registry needs every one of them
 * checked against the committed schema graph anyway.
 * `scripts/check-reserved-value-coverage.ts` expands the prefixes and resolves
 * each slot, so an unknown prefix or a slot no schema declares fails there.
 */
type ReservedValueSlots = string;

/**
 * The reserved value(s) the slot accepts, exactly as written in the markup.
 *
 * Alternatives are separated by `|` (`"nil|none"`). A few rules have no literal
 * to name — `"absent"` (a `w:vMerge` with no `@w:val` is a continuation cell),
 * `"both-present"` (`w:ind` carrying `@w:hanging` and `@w:firstLine` at once),
 * `"meaningless-under-auto"`, `"unresolvable-styleid"` — and those spellings
 * are fixed so the lint can tell a literal sentinel from a structural one.
 */
type ReservedValueSentinel = string;

/**
 * A slot folio models, read through exactly one function.
 *
 * `reader` comes from `RESERVED_VALUE_READERS` and carries the module path as
 * well as the name, because a bare name is not unique here. The reserved-value
 * lint exempts that module and flags a bare comparison against the sentinel
 * anywhere else.
 */
type ReaderOwnedSlot = {
  readonly disposition: "reader-owned";
  readonly slot: ReservedValueSlots;
  readonly sentinel: ReservedValueSentinel;
  readonly reader: string;
  /** Id of the `specifications/evidence` record that pins the claim, where the rule is prose-only. */
  readonly evidence?: string;
};

/**
 * A slot whose reserved value the model has no representation for.
 *
 * The parse boundary maps it into a named arm of a union instead — a
 * `w:outlineLvl w:val="9"` becomes `{ kind: "bodyText" }`, not the number 9 —
 * so no field downstream can hold the sentinel and no comparison against it
 * exists to be bare. This is the disposition to reach for once a reader-owned
 * slot's sentinel has been designed out: `carrier` names the type that
 * absorbed it, which is what a reviewer checks the claim against.
 */
type UnrepresentableSlot = {
  readonly disposition: "unrepresentable";
  readonly slot: ReservedValueSlots;
  readonly sentinel: ReservedValueSentinel;
  /** The model type whose arm the parse boundary maps the sentinel into. */
  readonly carrier: string;
  /** Id of the `specifications/evidence` record that pins the claim, where the rule is prose-only. */
  readonly evidence?: string;
};

/** A slot whose reserved value folio does not model, and why that is deliberate. */
type NotModelledSlot = {
  readonly disposition: "not-modelled";
  readonly slot: ReservedValueSlots;
  readonly sentinel: ReservedValueSentinel;
  readonly reason: string;
  /** Id of the `specifications/evidence` record that pins the claim, where the rule is prose-only. */
  readonly evidence?: string;
};

/**
 * `"no-reserved-value"` says the field's slot has none: every value it accepts
 * means itself, so no reader owns it and no comparison against it is bare.
 */
export type ReservedValueDisposition =
  | "no-reserved-value"
  | ReaderOwnedSlot
  | UnrepresentableSlot
  | NotModelledSlot;

/**
 * Every key of every member of a union type.
 *
 * `keyof` over a union gives only the keys every member shares, which would let
 * a variant-only field (`DrawingContent`'s `rawImageFingerprint`) slip through
 * with no decision.
 */
export type UnionFields<Source> = Source extends unknown ? keyof Source : never;

/**
 * The constructors below annotate their return type on purpose.
 *
 * This project is budgeted like any other (`bun run typecheck:budget`). Written
 * as `as const` object literals, a few hundred entries mint a few hundred
 * anonymous object types and check each one against a three-member union;
 * returned as {@link ReservedValueDisposition}, they are one type, checked once
 * per constructor. The maps keep their exact keys either way, which is where
 * the totality guarantee lives.
 */
export const NO_RESERVED_VALUE: ReservedValueDisposition = "no-reserved-value";

type ReaderOwnedOptions = {
  slot: ReservedValueSlots;
  sentinel: ReservedValueSentinel;
  reader: string;
  evidence?: string;
};

export const readerOwned = ({
  slot,
  sentinel,
  reader,
  evidence,
}: ReaderOwnedOptions): ReservedValueDisposition =>
  evidence === undefined
    ? { disposition: "reader-owned", slot, sentinel, reader }
    : { disposition: "reader-owned", slot, sentinel, reader, evidence };

type UnrepresentableOptions = {
  slot: ReservedValueSlots;
  sentinel: ReservedValueSentinel;
  carrier: string;
  evidence?: string;
};

export const unrepresentable = ({
  slot,
  sentinel,
  carrier,
  evidence,
}: UnrepresentableOptions): ReservedValueDisposition =>
  evidence === undefined
    ? { disposition: "unrepresentable", slot, sentinel, carrier }
    : { disposition: "unrepresentable", slot, sentinel, carrier, evidence };

type NotModelledOptions = {
  slot: ReservedValueSlots;
  sentinel: ReservedValueSentinel;
  reason: string;
  evidence?: string;
};

export const notModelled = ({
  slot,
  sentinel,
  reason,
  evidence,
}: NotModelledOptions): ReservedValueDisposition =>
  evidence === undefined
    ? { disposition: "not-modelled", slot, sentinel, reason }
    : { disposition: "not-modelled", slot, sentinel, reason, evidence };

/**
 * A toggle property: `w:val` accepts `0|false|off` as an explicit "not set",
 * which is not the same as the element being absent, and ECMA-376 combines two
 * levels of the style hierarchy by XOR rather than by override.
 */
export const toggle = (slot: ReservedValueSlots): ReservedValueDisposition =>
  readerOwned({
    slot,
    sentinel: "0|false|off",
    reader: RESERVED_VALUE_READERS.onOffValue,
    evidence: "toggle-property-xor",
  });

/** Every field decision recorded for one model type. */
export type ReservedValueMap = Readonly<Record<string, ReservedValueDisposition>>;
