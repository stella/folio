import type { ParagraphFormatting } from "../types/document";
import type { FolioAIParagraphSpacing } from "../ai-edits/types";
import { setAutospacingBaseValue } from "./autospacingBase";
import type { ParagraphSpacingInheritance } from "./paragraphPropertyContext";
import { expectParagraphPropertyState } from "./paragraphPropertyState";
import type { ParagraphAttrs } from "./schema/nodes";

/** The complete modeled attribute set of one direct `w:pPr/w:spacing`. */
export const DIRECT_PARAGRAPH_SPACING_KEYS = [
  "spaceBefore",
  "spaceAfter",
  "lineSpacing",
  "lineSpacingRule",
  "beforeAutospacing",
  "afterAutospacing",
] as const satisfies readonly (keyof ParagraphFormatting)[];

export type DirectParagraphSpacing = FolioAIParagraphSpacing;

type ExactLineSpacingProvenance = Exclude<
  NonNullable<ParagraphAttrs["lineSpacingExplicit"]>,
  boolean
>;

export const lineSpacingProvenanceFromSpacing = (
  spacing: DirectParagraphSpacing | null | undefined,
): ExactLineSpacingProvenance | undefined => {
  const hasValue = spacing?.lineSpacing !== undefined;
  const hasRule = spacing?.lineSpacingRule !== undefined;
  if (hasValue && hasRule) {
    return "both";
  }
  if (hasValue) {
    return "value";
  }
  return hasRule ? "rule" : undefined;
};

/** Project only authored `w:spacing` attributes, preserving explicit zero and false. */
export const paragraphSpacingFromFormatting = (
  formatting: DirectParagraphSpacing | null | undefined,
): DirectParagraphSpacing | undefined => {
  if (!formatting) {
    return undefined;
  }
  const spacing: DirectParagraphSpacing = {
    ...(formatting.spaceBefore !== undefined ? { spaceBefore: formatting.spaceBefore } : {}),
    ...(formatting.spaceAfter !== undefined ? { spaceAfter: formatting.spaceAfter } : {}),
    ...(formatting.lineSpacing !== undefined ? { lineSpacing: formatting.lineSpacing } : {}),
    ...(formatting.lineSpacingRule !== undefined
      ? { lineSpacingRule: formatting.lineSpacingRule }
      : {}),
    ...(formatting.beforeAutospacing !== undefined
      ? { beforeAutospacing: formatting.beforeAutospacing }
      : {}),
    ...(formatting.afterAutospacing !== undefined
      ? { afterAutospacing: formatting.afterAutospacing }
      : {}),
  };
  return Object.keys(spacing).length > 0 ? spacing : undefined;
};

/**
 * Read direct spacing independently of the effective values used for layout.
 * Mandatory authored state is the sole source. Effective spacing and its
 * auto-spacing baseline are projections, never fallback provenance.
 */
export const directParagraphSpacing = (
  attrs: ParagraphAttrs,
): DirectParagraphSpacing | undefined =>
  paragraphSpacingFromFormatting(
    expectParagraphPropertyState(attrs._paragraphPropertyState).authoredPPr,
  );

/** Structural equality for the canonical, fixed-order spacing projection. */
export const paragraphSpacingEqual = (
  left: DirectParagraphSpacing | null | undefined,
  right: DirectParagraphSpacing | null | undefined,
): boolean =>
  (left?.spaceBefore ?? null) === (right?.spaceBefore ?? null) &&
  (left?.spaceAfter ?? null) === (right?.spaceAfter ?? null) &&
  (left?.lineSpacing ?? null) === (right?.lineSpacing ?? null) &&
  (left?.lineSpacingRule ?? null) === (right?.lineSpacingRule ?? null) &&
  (left?.beforeAutospacing ?? null) === (right?.beforeAutospacing ?? null) &&
  (left?.afterAutospacing ?? null) === (right?.afterAutospacing ?? null);

/** Replace the complete direct `w:spacing` cluster in canonical formatting. */
export const withDirectParagraphSpacing = (
  formatting: ParagraphFormatting | null | undefined,
  spacing: DirectParagraphSpacing | null | undefined,
): ParagraphFormatting | undefined => {
  const result: ParagraphFormatting = { ...formatting };
  for (const key of DIRECT_PARAGRAPH_SPACING_KEYS) {
    Reflect.deleteProperty(result, key);
  }
  Reflect.deleteProperty(result, "spacingExplicit");
  if (spacing) {
    Object.assign(result, spacing);
    const spacingExplicit = {
      ...(spacing.spaceBefore !== undefined ? { before: true } : {}),
      ...(spacing.spaceAfter !== undefined ? { after: true } : {}),
    };
    if (Object.keys(spacingExplicit).length > 0) {
      result.spacingExplicit = spacingExplicit;
    }
  }
  return Object.keys(result).length > 0 ? result : undefined;
};

type ParagraphSpacingAttrPatchOptions = {
  direct: DirectParagraphSpacing | null | undefined;
  inherited: DirectParagraphSpacing | null | undefined;
  inheritance: ParagraphSpacingInheritance;
};

/** Project direct-over-inherited spacing into the attrs consumed by layout. */
export const paragraphSpacingAttrPatch = ({
  direct,
  inherited,
  inheritance,
}: ParagraphSpacingAttrPatchOptions): Record<string, unknown> => {
  const effective = { ...inherited, ...direct };
  const spacingExplicit = {
    ...(direct?.spaceBefore !== undefined ? { before: true } : {}),
    ...(direct?.spaceAfter !== undefined ? { after: true } : {}),
  };
  const spacingFromDocDefaults = {
    ...(direct?.spaceBefore === undefined && inheritance.before === "document-default"
      ? { before: true }
      : {}),
    ...(direct?.spaceAfter === undefined && inheritance.after === "document-default"
      ? { after: true }
      : {}),
  };
  const spacingFromImplicitDefaultStyle = {
    ...(direct?.spaceBefore === undefined && inheritance.before === "implicit-default-style"
      ? { before: true }
      : {}),
    ...(direct?.spaceAfter === undefined && inheritance.after === "implicit-default-style"
      ? { after: true }
      : {}),
  };
  const autospacingBase: NonNullable<ParagraphAttrs["_autospacingBase"]> = {};
  if (effective.beforeAutospacing === true) {
    setAutospacingBaseValue(autospacingBase, "before", effective.spaceBefore);
  }
  if (effective.afterAutospacing === true) {
    setAutospacingBaseValue(autospacingBase, "after", effective.spaceAfter);
  }
  return {
    spaceBefore: effective.spaceBefore ?? null,
    spaceAfter: effective.spaceAfter ?? null,
    lineSpacing: effective.lineSpacing ?? null,
    lineSpacingRule: effective.lineSpacingRule ?? null,
    lineSpacingExplicit: lineSpacingProvenanceFromSpacing(direct) ?? null,
    spacingExplicit: Object.keys(spacingExplicit).length > 0 ? spacingExplicit : null,
    spacingFromDocDefaults:
      Object.keys(spacingFromDocDefaults).length > 0 ? spacingFromDocDefaults : null,
    spacingFromImplicitDefaultStyle:
      Object.keys(spacingFromImplicitDefaultStyle).length > 0
        ? spacingFromImplicitDefaultStyle
        : null,
    _autospacingBase: Object.keys(autospacingBase).length > 0 ? autospacingBase : null,
  };
};
