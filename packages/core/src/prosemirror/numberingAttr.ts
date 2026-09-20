import { panic } from "better-result";
import {
  NO_PARAGRAPH_NUMBERING,
  type ParagraphNumberingOverride,
  paragraphNumberingFromSlots,
} from "@stll/docx-core/model";

declare const PARAGRAPH_NUMBERING_ATTR: unique symbol;

/**
 * The only value a persisted `numPr` attr slot may hold.
 *
 * It is the model's union, minted by {@link paragraphNumberingAttr} and by
 * nothing else. The brand is what makes the mint unavoidable: a
 * `ParagraphFormatting.numPr` read off a model object, or a record assembled
 * by hand, does not carry it and so cannot be assigned to an attr. That is the
 * defect the two tiers kept trading — a model value stored verbatim in an attr
 * (`toProseDoc`'s `_propertyChanges` projection) and an attr value stored
 * verbatim in the model-typed `_originalFormatting`
 * (`paragraphRejectOriginalFormatting`) — each of which read as the other's
 * shape on the far side.
 *
 * Reading needs no unbranding: the brand is a phantom property, so an attr
 * value is already a `ParagraphNumberingOverride` to every consumer.
 */
// The intersection adds a `unique symbol` phantom member, not a field: nothing
// is stored under it and no projection can drop it. The rule guards against
// widening a model type with real data, which is the opposite of what a brand
// does, and declaring the brand on `ParagraphNumberingOverride` would put an
// editor concern inside the model.
// oxlint-disable-next-line folio-model-types/no-model-intersection-widening -- phantom brand, no data field
export type ParagraphNumberingAttr = ParagraphNumberingOverride & {
  readonly [PARAGRAPH_NUMBERING_ATTR]: true;
};

// SAFETY: the brand is a phantom property that exists only in the type. The
// runtime value is the union itself, and this is the one function allowed to
// state that a value came from the codec.
const branded = (numbering: ParagraphNumberingOverride): ParagraphNumberingAttr =>
  numbering as ParagraphNumberingAttr;

/** Mint the attr value for a stated override: the one model-to-attr crossing. */
export const paragraphNumberingAttr = (
  numbering: ParagraphNumberingOverride,
): ParagraphNumberingAttr => {
  switch (numbering.kind) {
    case "none":
      return branded(NO_PARAGRAPH_NUMBERING);
    case "reference":
      return branded(
        numbering.ilvl === undefined
          ? { kind: "reference", numId: numbering.numId }
          : { kind: "reference", numId: numbering.numId, ilvl: numbering.ilvl },
      );
    case "levelOnly":
      return branded({ kind: "levelOnly", ilvl: numbering.ilvl });
    default: {
      const unhandled: never = numbering;
      return panic(`Unhandled paragraph numbering ${JSON.stringify(unhandled)}`);
    }
  }
};

/** A `<w:numPr>` slot value: an ordinal, so a non-integer or a negative is not one. */
const isNumberingSlot = (value: unknown): value is number =>
  typeof value === "number" && Number.isInteger(value) && value >= 0;

/**
 * The one shape check over a persisted `numPr` attr.
 *
 * The attr validator and the readers that go straight to `node.attrs` share
 * it, so a snapshot shape can be accepted in one place and rejected in the
 * other only by deleting this function.
 *
 * It is deliberately strict. ProseMirror's `computeAttrs` copies whatever a
 * stored document holds into the node without validating, so the pre-union
 * shape — the two `<w:numPr>` slots, where `numId` 0 meant a cancellation —
 * would otherwise arrive as a `numPr` whose `kind` is `undefined`, which every
 * `switch` over the union reads as the arm it is not. Returning `null` for it
 * is what lets the caller refuse.
 *
 * `ilvl` is bounded below and not above: real packages carry levels past the
 * nine the format defines, and docx-core preserves them.
 */
export const paragraphNumberingFromAttrValue = (value: unknown): ParagraphNumberingAttr | null => {
  if (value === null || typeof value !== "object" || Array.isArray(value)) {
    return null;
  }
  const kind: unknown = Reflect.get(value, "kind");
  if (kind === "none") {
    return branded(NO_PARAGRAPH_NUMBERING);
  }
  const ilvl: unknown = Reflect.get(value, "ilvl");
  if (kind === "levelOnly") {
    return isNumberingSlot(ilvl) ? branded({ kind: "levelOnly", ilvl }) : null;
  }
  if (kind !== "reference") {
    return null;
  }
  const numId: unknown = Reflect.get(value, "numId");
  if (!isNumberingSlot(numId) || (ilvl !== undefined && !isNumberingSlot(ilvl))) {
    return null;
  }
  // The model's own constructor decides what the id means, so the reserved
  // value has no spelling here: a stored `numId` it reads as a cancellation is
  // not a reference and this shape check refuses it.
  const stated = paragraphNumberingFromSlots({ ilvl, numId });
  return stated?.kind === "reference" ? branded(stated) : null;
};

/**
 * A paragraph's stated numbering, read straight from an unvalidated attrs
 * record. `null` is the attr's absent state, which is what the node spec's
 * default stores.
 *
 * A value that is present and not a {@link ParagraphNumberingAttr} panics
 * rather than being ignored: it can only come from a document written under an
 * attr schema this build does not read, and `yjsDocumentMetadata`'s version
 * gate exists to make that unreachable.
 */
export const readParagraphNumberingAttr = (value: unknown): ParagraphNumberingAttr | null => {
  if (value === undefined || value === null) {
    return null;
  }
  const numbering = paragraphNumberingFromAttrValue(value);
  if (numbering === null) {
    panic(`Invalid paragraph numPr attr: ${JSON.stringify(value)}`);
  }
  return numbering;
};
