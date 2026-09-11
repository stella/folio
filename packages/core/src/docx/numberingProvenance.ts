import {
  isNonEmptyNumberingLevelIndentGeometry,
  type NumberingIdentity,
  type NumberingLevelIndentGeometry,
} from "@stll/docx-core/model";
import { panic } from "better-result";

import type { ParagraphFormatting, ParagraphPropertyChange } from "../types/document";

export type NumberingIdRemapper =
  | ReadonlyMap<number, number>
  | ((numId: number) => number | undefined);

const remappedNumId = (remapper: NumberingIdRemapper, numId: number): number | undefined =>
  typeof remapper === "function" ? remapper(numId) : remapper.get(numId);

type NumberingLevelIndentProvenanceInput = {
  numId: number;
  ilvl: number;
  baseline: NumberingLevelIndentGeometry;
  owned?: NumberingLevelIndentGeometry;
};

/** Construct the only valid numbering-indent provenance branches. */
export const createNumberingLevelIndentProvenance = ({
  numId,
  ilvl,
  baseline,
  owned,
}: NumberingLevelIndentProvenanceInput): ParagraphFormatting["numberingLevelIndent"] => {
  if (!isNonEmptyNumberingLevelIndentGeometry(baseline)) {
    return undefined;
  }
  if (!isNonEmptyNumberingLevelIndentGeometry(owned)) {
    return { type: "latent", numId, ilvl, baseline };
  }
  for (const key of [
    "indentLeft",
    "indentRight",
    "indentFirstLine",
    "hangingIndent",
  ] as const) {
    if (owned[key] !== undefined && owned[key] !== baseline[key]) {
      panic("Numbering-owned indentation must match its level baseline.");
    }
  }
  return { type: "owned", numId, ilvl, baseline, owned };
};

/** Remap every numbering identity carried by one paragraph-formatting snapshot. */
export const remapParagraphFormattingNumbering = (
  formatting: ParagraphFormatting | undefined,
  numIdRemap: NumberingIdRemapper,
): ParagraphFormatting | undefined => {
  if (!formatting) {
    return undefined;
  }

  let next = formatting;
  const sourceNumIds = new Set(
    [
      formatting.numPr,
      formatting.numPrFromStyle,
      formatting.numberingLevelIndent,
    ]
      .map((identity) => identity?.numId)
      .filter((numId): numId is number => numId !== undefined && numId !== 0),
  );
  for (const numId of sourceNumIds) {
    if (remappedNumId(numIdRemap, numId) === undefined) {
      next = dropParagraphNumberingReference(next, numId);
    }
  }

  const remapIdentity = <T extends NumberingIdentity>(identity: T | undefined): T | undefined => {
    const numId = identity?.numId;
    if (numId === undefined || numId === 0) {
      return identity;
    }
    const remapped = remappedNumId(numIdRemap, numId);
    return remapped === undefined || remapped === numId
      ? identity
      : { ...identity, numId: remapped };
  };
  const numPr = remapIdentity(next.numPr);
  const numPrFromStyle = remapIdentity(next.numPrFromStyle);
  const numberingLevelIndent = remapIdentity(next.numberingLevelIndent);
  if (
    numPr === next.numPr &&
    numPrFromStyle === next.numPrFromStyle &&
    numberingLevelIndent === next.numberingLevelIndent
  ) {
    return next;
  }
  const remapped = { ...next };
  const assignIdentity = <
    K extends "numPr" | "numPrFromStyle" | "numberingLevelIndent",
  >(
    key: K,
    identity: ParagraphFormatting[K],
  ): void => {
    if (identity === undefined) {
      Reflect.deleteProperty(remapped, key);
      return;
    }
    Object.assign(remapped, { [key]: identity });
  };
  assignIdentity("numPr", numPr);
  assignIdentity("numPrFromStyle", numPrFromStyle);
  assignIdentity("numberingLevelIndent", numberingLevelIndent);
  return remapped;
};

/** Remap current and historical paragraph-property numbering identities together. */
export const remapParagraphPropertyChangeNumbering = (
  changes: ParagraphPropertyChange[] | undefined,
  numIdRemap: NumberingIdRemapper,
): ParagraphPropertyChange[] | undefined => {
  if (!changes) {
    return undefined;
  }
  let changed = false;
  const remapped = changes.map((change) => {
    const previousFormatting = remapParagraphFormattingNumbering(
      change.previousFormatting,
      numIdRemap,
    );
    const currentFormatting = remapParagraphFormattingNumbering(
      change.currentFormatting,
      numIdRemap,
    );
    if (
      previousFormatting === change.previousFormatting &&
      currentFormatting === change.currentFormatting
    ) {
      return change;
    }
    changed = true;
    return {
      ...change,
      ...(previousFormatting === undefined ? {} : { previousFormatting }),
      ...(currentFormatting === undefined ? {} : { currentFormatting }),
    };
  });
  return changed ? remapped : changes;
};

const omitOwnedIndentValue = <K extends "indentLeft" | "indentRight">(
  formatting: ParagraphFormatting,
  provenance: NumberingLevelIndentGeometry,
  key: K,
): void => {
  if (provenance[key] !== undefined && formatting[key] === provenance[key]) {
    Reflect.deleteProperty(formatting, key);
  }
};

/** Drop an invalid active numbering reference and only the effective indents it owned. */
export const dropParagraphNumberingReference = (
  formatting: ParagraphFormatting,
  numId: number,
): ParagraphFormatting => {
  const next = { ...formatting };
  const provenance = formatting.numberingLevelIndent;
  if (formatting.numPr?.numId === numId) {
    Reflect.deleteProperty(next, "numPr");
  }
  if (provenance?.numId === numId) {
    const owned = provenance.type === "owned" ? provenance.owned : undefined;
    if (owned) {
      omitOwnedIndentValue(next, owned, "indentLeft");
      omitOwnedIndentValue(next, owned, "indentRight");
    }
    const hasOwnedFirstLinePair =
      owned?.indentFirstLine !== undefined || owned?.hangingIndent !== undefined;
    const ownsFirstLinePair =
      hasOwnedFirstLinePair &&
      formatting.indentFirstLine === owned?.indentFirstLine &&
      (formatting.hangingIndent ?? false) === (owned?.hangingIndent ?? false);
    if (ownsFirstLinePair) {
      Reflect.deleteProperty(next, "indentFirstLine");
      Reflect.deleteProperty(next, "hangingIndent");
    }
    Reflect.deleteProperty(next, "numberingLevelIndent");
  }
  if (formatting.numPrFromStyle?.numId === numId) {
    Reflect.deleteProperty(next, "numPrFromStyle");
  }
  return next;
};

/** Remove numbering-owned effective indents, leaving direct/style fields intact. */
export const paragraphFormattingWithoutNumberingOwnedIndent = (
  formatting: ParagraphFormatting,
): ParagraphFormatting => {
  const result = { ...formatting };
  const provenance = formatting.numberingLevelIndent;
  if (provenance?.type !== "owned") {
    Reflect.deleteProperty(result, "numberingLevelIndent");
    return result;
  }
  const owned = provenance.owned;
  if (owned.indentLeft !== undefined && result.indentLeft === owned.indentLeft) {
    Reflect.deleteProperty(result, "indentLeft");
  }
  if (owned.indentRight !== undefined && result.indentRight === owned.indentRight) {
    Reflect.deleteProperty(result, "indentRight");
  }
  if (
    (owned.indentFirstLine !== undefined || owned.hangingIndent !== undefined) &&
    result.indentFirstLine === owned.indentFirstLine &&
    (result.hangingIndent ?? false) === (owned.hangingIndent ?? false)
  ) {
    Reflect.deleteProperty(result, "indentFirstLine");
    Reflect.deleteProperty(result, "hangingIndent");
  }
  Reflect.deleteProperty(result, "numberingLevelIndent");
  return result;
};

export type NumberingIndentSide = "indentLeft" | "indentRight" | "firstLine";

/** Reclassify rewritten numbering-derived geometry as direct paragraph formatting. */
export const retireNumberingIndentOwnership = (
  formatting: ParagraphFormatting,
  sides: ReadonlySet<NumberingIndentSide>,
): ParagraphFormatting => {
  const provenance = formatting.numberingLevelIndent;
  if (provenance?.type !== "owned") {
    return formatting;
  }
  const nextOwned: NumberingLevelIndentGeometry = { ...provenance.owned };
  if (sides.has("indentLeft")) {
    Reflect.deleteProperty(nextOwned, "indentLeft");
  }
  if (sides.has("indentRight")) {
    Reflect.deleteProperty(nextOwned, "indentRight");
  }
  if (sides.has("firstLine")) {
    Reflect.deleteProperty(nextOwned, "indentFirstLine");
    Reflect.deleteProperty(nextOwned, "hangingIndent");
  }
  const next = { ...formatting };
  next.numberingLevelIndent = createNumberingLevelIndentProvenance({
    numId: provenance.numId,
    ilvl: provenance.ilvl,
    baseline: provenance.baseline,
    owned: nextOwned,
  });
  return next;
};
