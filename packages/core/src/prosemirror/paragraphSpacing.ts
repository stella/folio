import type { ParagraphFormatting } from "../types/document";
import type { FolioAIParagraphSpacing } from "../ai-edits/types";
import {
  autospacingMatchesBase,
  hasAutospacingBaseSide,
  setAutospacingBaseValue,
} from "./autospacingBase";
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
 * Imported and command-authored paragraphs keep their source in
 * `_originalFormatting`; explicit PM attrs cover newly created content.
 */
export const directParagraphSpacing = (
  attrs: ParagraphAttrs,
): DirectParagraphSpacing | undefined => {
  const spacing = paragraphSpacingFromFormatting(attrs._originalFormatting) ?? {};
  const spaceBefore: unknown = attrs.spaceBefore;
  const spaceAfter: unknown = attrs.spaceAfter;
  const beforeHasBase = hasAutospacingBaseSide(attrs._autospacingBase, "before");
  const afterHasBase = hasAutospacingBaseSide(attrs._autospacingBase, "after");
  const beforeAutospacingEdited = beforeHasBase
    ? !autospacingMatchesBase(attrs._autospacingBase, "before", spaceBefore)
    : spacing.beforeAutospacing === true && attrs._autospacingBase == null;
  const afterAutospacingEdited = afterHasBase
    ? !autospacingMatchesBase(attrs._autospacingBase, "after", spaceAfter)
    : spacing.afterAutospacing === true && attrs._autospacingBase == null;

  if (beforeAutospacingEdited) {
    spacing.beforeAutospacing = false;
    if (typeof spaceBefore === "number") {
      spacing.spaceBefore = spaceBefore;
    } else {
      Reflect.deleteProperty(spacing, "spaceBefore");
    }
  }
  if (afterAutospacingEdited) {
    spacing.afterAutospacing = false;
    if (typeof spaceAfter === "number") {
      spacing.spaceAfter = spaceAfter;
    } else {
      Reflect.deleteProperty(spacing, "spaceAfter");
    }
  }
  if (spacing.spaceBefore !== undefined || attrs.spacingExplicit?.before === true) {
    if (typeof spaceBefore === "number") {
      spacing.spaceBefore = spaceBefore;
    } else {
      Reflect.deleteProperty(spacing, "spaceBefore");
    }
  }
  if (spacing.spaceAfter !== undefined || attrs.spacingExplicit?.after === true) {
    if (typeof spaceAfter === "number") {
      spacing.spaceAfter = spaceAfter;
    } else {
      Reflect.deleteProperty(spacing, "spaceAfter");
    }
  }
  if (attrs.lineSpacingExplicit === true || spacing.lineSpacing !== undefined) {
    if (typeof attrs.lineSpacing === "number") {
      spacing.lineSpacing = attrs.lineSpacing;
    } else {
      Reflect.deleteProperty(spacing, "lineSpacing");
    }
  }
  if (attrs.lineSpacingRuleExplicit === true || spacing.lineSpacingRule !== undefined) {
    if (attrs.lineSpacingRule !== undefined) {
      spacing.lineSpacingRule = attrs.lineSpacingRule;
    } else {
      Reflect.deleteProperty(spacing, "lineSpacingRule");
    }
  }
  return Object.keys(spacing).length > 0 ? spacing : undefined;
};

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
};

/** Project direct-over-inherited spacing into the attrs consumed by layout. */
export const paragraphSpacingAttrPatch = ({
  direct,
  inherited,
}: ParagraphSpacingAttrPatchOptions): Record<string, unknown> => {
  const effective = { ...inherited, ...direct };
  const spacingExplicit = {
    ...(direct?.spaceBefore !== undefined ? { before: true } : {}),
    ...(direct?.spaceAfter !== undefined ? { after: true } : {}),
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
    lineSpacingExplicit: direct?.lineSpacing !== undefined ? true : null,
    lineSpacingRuleExplicit: direct?.lineSpacingRule !== undefined ? true : null,
    spacingExplicit: Object.keys(spacingExplicit).length > 0 ? spacingExplicit : null,
    _autospacingBase: Object.keys(autospacingBase).length > 0 ? autospacingBase : null,
  };
};
