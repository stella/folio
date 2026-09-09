import type { Mark, Node as PMNode } from "prosemirror-model";

import type { StyleEngine } from "../style-engine";
import type { TextFormatting } from "../types/document";
import { mergeTextFormatting, STYLE_TOGGLE_KEYS } from "../utils/textFormattingMerge";
import { expectCharacterStyleMarkAttrs, expectParagraphAttrs } from "./attrs";
import {
  cascadeStyleFromResolvedBase,
  cascadeStyleTextFormatting,
  type StyleToggleKey,
} from "./styles/styleToggleCascade";

const PARAGRAPH_MARK_BOOLEAN_KEYS = [
  "bold",
  "italic",
  "strike",
  "doubleStrike",
  "allCaps",
  "smallCaps",
  "hidden",
  "emboss",
  "imprint",
  "shadow",
  "outline",
  "rtl",
] as const satisfies readonly (keyof TextFormatting)[];

export const hasDirectRunFormatting = (formatting: TextFormatting | undefined): boolean => {
  if (!formatting) {
    return false;
  }
  return Object.entries(formatting).some(
    ([property, value]) => property !== "styleId" && value !== undefined,
  );
};

export const stripParagraphMarkOnlyFormatting = (
  formatting: TextFormatting,
): TextFormatting | undefined => {
  const {
    allCaps: _allCaps,
    highlight: _highlight,
    shading: _shading,
    smallCaps: _smallCaps,
    vertAlign: _vertAlign,
    ...formattingForBody
  } = formatting;
  return Object.keys(formattingForBody).length > 0 ? formattingForBody : undefined;
};

export const stripParagraphMarkFormattingForBodyRuns = (
  formatting: TextFormatting,
): TextFormatting | undefined => {
  const paragraphMarkFormatting = stripParagraphMarkOnlyFormatting(formatting);
  if (!paragraphMarkFormatting) {
    return undefined;
  }
  const { fontFamily: _fontFamily, ...bodyRunFormatting } = paragraphMarkFormatting;
  return Object.keys(bodyRunFormatting).length > 0 ? bodyRunFormatting : undefined;
};

type SuppressParagraphMarkFormattingOptions = {
  baseFormatting: TextFormatting | undefined;
  directFormatting: TextFormatting | undefined;
  paragraphMarkFormatting: TextFormatting | undefined;
  paragraphMarkPrecedesStyle?: boolean;
};

export const suppressParagraphMarkFormatting = ({
  baseFormatting,
  directFormatting,
  paragraphMarkFormatting,
  paragraphMarkPrecedesStyle = false,
}: SuppressParagraphMarkFormattingOptions): TextFormatting | undefined => {
  if (!paragraphMarkFormatting) {
    return baseFormatting;
  }
  const result =
    (paragraphMarkPrecedesStyle
      ? mergeTextFormatting(paragraphMarkFormatting, baseFormatting)
      : mergeTextFormatting(baseFormatting, paragraphMarkFormatting)) ?? {};
  for (const property of PARAGRAPH_MARK_BOOLEAN_KEYS) {
    if (
      paragraphMarkFormatting[property] !== undefined &&
      directFormatting?.[property] === undefined
    ) {
      Reflect.set(result, property, baseFormatting?.[property] ?? false);
    }
  }
  for (const property of ["fontSize", "fontSizeCs"] as const) {
    if (
      paragraphMarkFormatting[property] !== undefined &&
      directFormatting?.[property] === undefined &&
      baseFormatting?.[property] !== undefined
    ) {
      result[property] = baseFormatting[property];
    }
  }
  if (
    paragraphMarkFormatting.underline !== undefined &&
    directFormatting?.underline === undefined
  ) {
    result.underline = { style: "none" };
  }
  if (paragraphMarkFormatting.spacing !== undefined && directFormatting?.spacing === undefined) {
    result.spacing = 0;
  }
  return Object.keys(result).length > 0 ? result : undefined;
};

type GetParagraphMarkSuppressionOverridesOptions = {
  directFormatting: TextFormatting | undefined;
  paragraphMarkFormatting: TextFormatting | undefined;
  suppressedFormatting: TextFormatting | undefined;
};

export const getParagraphMarkSuppressionOverrides = ({
  directFormatting,
  paragraphMarkFormatting,
  suppressedFormatting,
}: GetParagraphMarkSuppressionOverridesOptions): TextFormatting | undefined => {
  if (!paragraphMarkFormatting) {
    return undefined;
  }
  const overrides: TextFormatting = {};
  for (const property of PARAGRAPH_MARK_BOOLEAN_KEYS) {
    if (
      paragraphMarkFormatting[property] !== undefined &&
      directFormatting?.[property] === undefined &&
      suppressedFormatting?.[property] === false
    ) {
      Reflect.set(overrides, property, false);
    }
  }
  if (
    paragraphMarkFormatting.underline !== undefined &&
    directFormatting?.underline === undefined &&
    suppressedFormatting?.underline?.style === "none"
  ) {
    overrides.underline = { style: "none" };
  }
  return Object.keys(overrides).length > 0 ? overrides : undefined;
};

export type ParagraphRunStyleContext = {
  baseParagraphFormatting: TextFormatting | undefined;
  paragraphFormatting: TextFormatting | undefined;
  paragraphMarkFormatting: TextFormatting | undefined;
  paragraphMarkPrecedesStyle: boolean;
};

export type RunStyleResolver = Pick<
  StyleEngine,
  | "getDefaultCharacterStyle"
  | "getDocDefaults"
  | "getRunStyleOwnProperties"
  | "resolveParagraphStyle"
>;

export const paragraphRunStyleContext = (
  paragraph: PMNode,
  styleResolver?: RunStyleResolver | null,
): ParagraphRunStyleContext => {
  const attrs = expectParagraphAttrs(paragraph);
  const paragraphMarkFormatting = attrs._originalFormatting?.runProperties;
  const characterStyleFormatting = paragraphMarkFormatting?.styleId
    ? styleResolver?.getRunStyleOwnProperties(paragraphMarkFormatting.styleId)
    : undefined;
  const resolvedParagraphMarkFormatting = cascadeStyleTextFormatting([
    { formatting: characterStyleFormatting, type: "style" },
    { formatting: paragraphMarkFormatting, type: "direct" },
  ]).formatting;
  const inheritableParagraphMarkFormatting = resolvedParagraphMarkFormatting
    ? stripParagraphMarkFormattingForBodyRuns(resolvedParagraphMarkFormatting)
    : undefined;
  return {
    baseParagraphFormatting: styleResolver?.resolveParagraphStyle(attrs.styleId).runFormatting,
    paragraphFormatting: attrs.defaultTextFormatting ?? undefined,
    paragraphMarkFormatting: inheritableParagraphMarkFormatting,
    paragraphMarkPrecedesStyle: attrs.styleId !== undefined,
  };
};

export const paragraphFormattingForRun = (
  marks: readonly Mark[],
  context: ParagraphRunStyleContext,
  directFormatting?: TextFormatting,
): TextFormatting | undefined => {
  if (
    !marks.some(
      ({ type }) => type.name === "runFormattingOverride" || type.name === "characterStyle",
    )
  ) {
    return context.paragraphFormatting;
  }
  return suppressParagraphMarkFormatting({
    baseFormatting: context.baseParagraphFormatting,
    directFormatting,
    paragraphMarkFormatting: context.paragraphMarkFormatting,
    paragraphMarkPrecedesStyle: context.paragraphMarkPrecedesStyle,
  });
};

export const paragraphRunStyleContextAt = (
  doc: PMNode,
  pos: number,
  styleResolver?: RunStyleResolver | null,
): ParagraphRunStyleContext => {
  const resolved = doc.resolve(pos);
  for (let depth = resolved.depth; depth >= 0; depth--) {
    const ancestor = resolved.node(depth);
    if (ancestor.type.name === "paragraph") {
      return paragraphRunStyleContext(ancestor, styleResolver);
    }
  }
  return {
    baseParagraphFormatting: undefined,
    paragraphFormatting: undefined,
    paragraphMarkFormatting: undefined,
    paragraphMarkPrecedesStyle: false,
  };
};

type ResolveEffectiveRunStyleFormattingOptions = {
  marks: readonly Mark[];
  paragraphFormatting: TextFormatting | undefined;
  styleResolver?: RunStyleResolver | null;
};

/** Resolve the non-direct run cascade from shared document styles and the current paragraph. */
export const resolveEffectiveRunStyleFormatting = ({
  marks,
  paragraphFormatting,
  styleResolver,
}: ResolveEffectiveRunStyleFormattingOptions): TextFormatting | undefined => {
  const characterStyle = marks.find(({ type }) => type.name === "characterStyle");
  const hasDirectFormattingCarrier = marks.some(
    ({ type }) => type.name === "runFormattingOverride",
  );
  let styleFormatting: TextFormatting | undefined;
  if (characterStyle) {
    styleFormatting = styleResolver?.getRunStyleOwnProperties(
      expectCharacterStyleMarkAttrs(characterStyle).styleId,
    );
  } else if (hasDirectFormattingCarrier) {
    styleFormatting = styleResolver?.getDefaultCharacterStyle()?.rPr;
  }
  if (!styleFormatting) {
    return paragraphFormatting;
  }
  const docDefaults = styleResolver?.getDocDefaults()?.rPr;
  const defaultsActiveProperties = STYLE_TOGGLE_KEYS.filter(
    (property): property is StyleToggleKey => docDefaults?.[property] === true,
  );
  return cascadeStyleFromResolvedBase({
    baseFormatting: paragraphFormatting,
    defaultsActiveProperties,
    styleFormatting,
  });
};
