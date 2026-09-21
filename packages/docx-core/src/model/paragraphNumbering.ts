import { panic } from "better-result";

/**
 * `w:numPr`: what a tier states about a paragraph's numbering, and what the
 * cascade resolves that to.
 *
 * Two facts shape the types, and the old `{ numId?: number; ilvl?: number }`
 * pair expressed neither.
 *
 * **`w:numId` and `w:ilvl` inherit independently** (ECMA-376 17.3.1.19). A tier
 * that states only the level keeps the id it inherits, which is what Word
 * writes whenever a styled list paragraph is demoted, so a direct `w:numPr` is
 * a partial override rather than a whole replacement. That is the `levelOnly`
 * arm, and it is why the override type has three arms and not two.
 *
 * **`w:numId w:val="0"` is reserved** (17.9.18): it names no `w:num`, it
 * switches numbering off, and on a style it cancels the numbering the style
 * would otherwise inherit through `w:basedOn`. It is the `none` arm, and a
 * `reference` therefore cannot carry it — which is the whole point, because
 * the sentinel used to read as "this paragraph is in list 0" in the layout and
 * marker tiers while the sanctioned reader said it was in no list at all.
 */

/** The reserved `w:numId` value that means "no numbering". */
export const NO_NUMBERING_NUM_ID = 0;

/**
 * Whether a `w:numId` names a numbering definition to resolve.
 *
 * Still the reader for the raw ids the numbering part itself carries
 * (`w:num@numId`, a style's own reference before it is read into the union).
 * A paragraph's numbering no longer goes through it: after
 * {@link paragraphNumberingFromSlots} the sentinel has no representation.
 */
export const isNumberingReference = (numId: number | undefined): numId is number =>
  numId !== undefined && numId !== NO_NUMBERING_NUM_ID;

declare const PARAGRAPH_NUMBERING_REFERENCE: unique symbol;

/** A non-reserved numbering reference minted by {@link paragraphNumberingReference}. */
export type ParagraphNumberingReference = {
  readonly kind: "reference";
  readonly numId: number;
  readonly ilvl?: number;
  readonly [PARAGRAPH_NUMBERING_REFERENCE]: true;
};

/**
 * What one tier's `<w:numPr>` states. The field being absent is the fourth
 * state and means the tier states nothing at all.
 */
export type ParagraphNumberingOverride =
  | { readonly kind: "none" }
  | ParagraphNumberingReference
  | { readonly kind: "levelOnly"; readonly ilvl: number };

/** Numbering after the cascade: what actually renders. */
export type ResolvedParagraphNumbering =
  | { readonly kind: "none" }
  | (ParagraphNumberingReference & { readonly ilvl: number });

/** The arm a `w:numId w:val="0"` reads as, and the arm a cascade cancels to. */
export const NO_PARAGRAPH_NUMBERING = { kind: "none" } as const;

/** The level an absent `w:ilvl` stands for. */
const IMPLICIT_NUMBERING_LEVEL = 0;

/**
 * The two slots a `<w:numPr>` carries, as read off the element.
 *
 * An absent `w:ilvl` and `w:ilvl="0"` mean the same level but are different
 * bytes, so `ilvl` stays optional inside `reference`: collapsing it turns an
 * untouched save into a document that states something its source did not.
 */
export type ParagraphNumberingSlots = {
  numId?: number;
  ilvl?: number;
};

type ParagraphNumberingReferenceOptions<NumId extends number, Level extends number | undefined> = {
  numId: Exclude<NumId, typeof NO_NUMBERING_NUM_ID>;
  ilvl?: Level | undefined;
};

type ParagraphNumberingReferenceResult<Level extends number | undefined> =
  ParagraphNumberingReference & (Level extends number ? { readonly ilvl: Level } : unknown);

/** Mint a reference after excluding the reserved id from the model. */
export const paragraphNumberingReference = <
  const NumId extends number,
  const Level extends number | undefined = undefined,
>({
  numId,
  ilvl,
}: ParagraphNumberingReferenceOptions<NumId, Level>): ParagraphNumberingReferenceResult<Level> => {
  if (!isNumberingReference(numId)) {
    return panic(`Paragraph numbering reference cannot use reserved numId ${numId}`);
  }
  const reference =
    ilvl === undefined ? { kind: "reference", numId } : { kind: "reference", numId, ilvl };
  // SAFETY: this constructor rejects the only id excluded by the reference
  // arm. The symbol is a phantom brand; the serialized model stays plain.
  return reference as ParagraphNumberingReferenceResult<Level>;
};

/**
 * The one mapping from a `<w:numPr>`'s two slots onto the union.
 *
 * `undefined` is an element that states neither slot, which states nothing.
 * Every parse boundary goes through this, which is what stops the reserved id
 * reaching the model at all.
 */
export const paragraphNumberingFromSlots = ({
  numId,
  ilvl,
}: {
  numId?: number | undefined;
  ilvl?: number | undefined;
}): ParagraphNumberingOverride | undefined => {
  if (numId === NO_NUMBERING_NUM_ID) {
    return NO_PARAGRAPH_NUMBERING;
  }
  if (numId !== undefined) {
    return paragraphNumberingReference({ numId, ilvl });
  }
  return ilvl === undefined ? undefined : { kind: "levelOnly", ilvl };
};

/**
 * The slots to write back for a stated override, so the emit has one owner
 * too. `none` writes the reserved id, because deleting the element would
 * uncover the tier below and hand the paragraph its numbering back.
 */
export const paragraphNumberingSlots = (
  numbering: ParagraphNumberingOverride,
): ParagraphNumberingSlots => {
  switch (numbering.kind) {
    case "none":
      return { numId: NO_NUMBERING_NUM_ID };
    case "reference":
      return numbering.ilvl === undefined
        ? { numId: numbering.numId }
        : { numId: numbering.numId, ilvl: numbering.ilvl };
    case "levelOnly":
      return { ilvl: numbering.ilvl };
    default: {
      const unhandled: never = numbering;
      return panic(`Unhandled paragraph numbering ${JSON.stringify(unhandled)}`);
    }
  }
};

/**
 * Fold one tier over the tier beneath it, direct over style.
 *
 * This is ECMA-376 17.3.1.19's independent inheritance written once. Before
 * the union it was four object spreads in three languages, each of them a
 * comment saying what the spread happened to do.
 */
export const mergeParagraphNumbering = (
  inherited: ParagraphNumberingOverride | undefined,
  stated: ParagraphNumberingOverride | undefined,
): ParagraphNumberingOverride | undefined => {
  if (stated === undefined) {
    return inherited;
  }
  switch (stated.kind) {
    case "none":
      return stated;
    case "reference": {
      if (stated.ilvl !== undefined || inherited === undefined || inherited.kind === "none") {
        return stated;
      }
      const inheritedLevel = inherited.ilvl;
      return inheritedLevel === undefined
        ? stated
        : paragraphNumberingReference({ numId: stated.numId, ilvl: inheritedLevel });
    }
    case "levelOnly": {
      if (inherited === undefined) {
        return stated;
      }
      switch (inherited.kind) {
        case "none":
          // A cancelled tier states there is no id to keep the level of.
          return NO_PARAGRAPH_NUMBERING;
        case "levelOnly":
          return stated;
        case "reference":
          return paragraphNumberingReference({ numId: inherited.numId, ilvl: stated.ilvl });
        default: {
          const unhandled: never = inherited;
          return panic(`Unhandled inherited numbering ${JSON.stringify(unhandled)}`);
        }
      }
    }
    default: {
      const unhandled: never = stated;
      return panic(`Unhandled stated numbering ${JSON.stringify(unhandled)}`);
    }
  }
};

/**
 * Collapse a stated override to what renders.
 *
 * A `levelOnly` that nothing supplied an id for numbers nothing, the same as a
 * cancellation does, which is why the resolved type has two arms.
 */
export const resolveParagraphNumbering = (
  numbering: ParagraphNumberingOverride | undefined,
): ResolvedParagraphNumbering => {
  if (numbering === undefined) {
    return NO_PARAGRAPH_NUMBERING;
  }
  switch (numbering.kind) {
    case "none":
    case "levelOnly":
      return NO_PARAGRAPH_NUMBERING;
    case "reference":
      return paragraphNumberingReference({
        numId: numbering.numId,
        ilvl: numbering.ilvl ?? IMPLICIT_NUMBERING_LEVEL,
      });
    default: {
      const unhandled: never = numbering;
      return panic(`Unhandled paragraph numbering ${JSON.stringify(unhandled)}`);
    }
  }
};

/** The level a stated override renders at, absent `w:ilvl` read as zero. */
export const paragraphNumberingLevel = (
  numbering: ParagraphNumberingOverride | undefined,
): number | undefined => {
  if (numbering === undefined) {
    return undefined;
  }
  switch (numbering.kind) {
    case "none":
      return undefined;
    case "levelOnly":
      return numbering.ilvl;
    case "reference":
      return numbering.ilvl ?? IMPLICIT_NUMBERING_LEVEL;
    default: {
      const unhandled: never = numbering;
      return panic(`Unhandled paragraph numbering ${JSON.stringify(unhandled)}`);
    }
  }
};

/** The numbering definition a stated override names, or undefined for none. */
export const paragraphNumberingReferenceId = (
  numbering: ParagraphNumberingOverride | undefined,
): number | undefined => (numbering?.kind === "reference" ? numbering.numId : undefined);

/**
 * Whether two stated overrides say the same thing, byte for byte: an absent
 * `w:ilvl` is not a stated `w:val="0"` here, which is what keeps an untouched
 * save from stating something its source did not.
 */
export const sameStatedParagraphNumbering = (
  left: ParagraphNumberingOverride | undefined,
  right: ParagraphNumberingOverride | undefined,
): boolean => {
  if (left === undefined || right === undefined) {
    return left === right;
  }
  switch (left.kind) {
    case "none":
      return right.kind === "none";
    case "reference":
      return right.kind === "reference" && left.numId === right.numId && left.ilvl === right.ilvl;
    case "levelOnly":
      return right.kind === "levelOnly" && left.ilvl === right.ilvl;
    default: {
      const unhandled: never = left;
      return panic(`Unhandled paragraph numbering ${JSON.stringify(unhandled)}`);
    }
  }
};

/** Whether two stated overrides render the same, an absent `w:ilvl` read as zero. */
export const sameEffectiveParagraphNumbering = (
  left: ParagraphNumberingOverride | undefined,
  right: ParagraphNumberingOverride | undefined,
): boolean => {
  if (left === undefined || right === undefined) {
    return left === right;
  }
  const resolvedLeft = resolveParagraphNumbering(left);
  const resolvedRight = resolveParagraphNumbering(right);
  if (resolvedLeft.kind === "none" || resolvedRight.kind === "none") {
    return resolvedLeft.kind === resolvedRight.kind;
  }
  return resolvedLeft.numId === resolvedRight.numId && resolvedLeft.ilvl === resolvedRight.ilvl;
};
