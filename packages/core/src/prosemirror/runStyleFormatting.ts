import type { Mark, Node as PMNode } from "prosemirror-model";
import type { ParagraphMarkProperties } from "@stll/docx-core/model";

import type { StyleEngine } from "../style-engine";
import type { TextFormatting } from "../types/document";
import {
  mergeStyleTextFormatting,
  mergeTextFormatting,
  STYLE_TOGGLE_KEYS,
} from "../utils/textFormattingMerge";
import { expectCharacterStyleMarkAttrs, expectParagraphAttrs } from "./attrs";
import { expectParagraphPropertyState } from "./paragraphPropertyState";
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
  | "getDefaultParagraphStyle"
  | "getDocDefaults"
  | "getStyle"
  | "getRunStyleOwnProperties"
  | "resolveParagraphStyle"
>;

export const resolveRunFormattingWithoutDefaults = (
  formatting: TextFormatting | undefined,
  styleResolver: Pick<StyleEngine, "getRunStyleOwnProperties"> | null,
): TextFormatting | undefined => {
  if (!formatting || !styleResolver) {
    return formatting;
  }
  const characterStyleFormatting = formatting.styleId
    ? styleResolver.getRunStyleOwnProperties(formatting.styleId)
    : undefined;
  return cascadeStyleTextFormatting([
    { formatting: characterStyleFormatting, type: "style" },
    { formatting, type: "direct" },
  ]).formatting;
};

const resolveParagraphStyleRunFormatting = (
  styleId: string | undefined,
  styleResolver: RunStyleResolver,
): TextFormatting | undefined => {
  let style = styleId ? styleResolver.getStyle(styleId) : styleResolver.getDefaultParagraphStyle();
  const visited = new Set<string>();
  const chain: TextFormatting[] = [];
  while (style?.type === "paragraph" && !visited.has(style.styleId)) {
    visited.add(style.styleId);
    if (style.rPr) {
      chain.push(style.rPr);
    }
    style = style.basedOn ? styleResolver.getStyle(style.basedOn) : undefined;
  }
  let formatting: TextFormatting | undefined;
  for (const inherited of chain.reverse()) {
    formatting = mergeStyleTextFormatting(formatting, inherited);
  }
  return formatting;
};

export const resolveParagraphRunStyleBase = (
  styleId: string | undefined,
  styleResolver: RunStyleResolver | null | undefined,
  tableRunFormatting?: TextFormatting,
): ReturnType<typeof cascadeStyleTextFormatting> => {
  const paragraphStyleRunFormatting = styleResolver
    ? resolveParagraphStyleRunFormatting(styleId, styleResolver)
    : undefined;
  const styleRunFormatting = cascadeStyleTextFormatting([
    { formatting: styleResolver?.getDocDefaults()?.rPr, type: "defaults" },
    { formatting: paragraphStyleRunFormatting, type: "style" },
  ]).formatting;
  const ordinaryStyleFormatting =
    styleId === undefined
      ? mergeTextFormatting(styleRunFormatting, tableRunFormatting)
      : mergeTextFormatting(tableRunFormatting, styleRunFormatting);
  const cascade = cascadeStyleTextFormatting(
    [
      { formatting: styleResolver?.getDocDefaults()?.rPr, type: "defaults" },
      { formatting: tableRunFormatting, type: "style" },
      { formatting: paragraphStyleRunFormatting, type: "style" },
    ],
    { ordinaryFormatting: ordinaryStyleFormatting },
  );
  const paragraphStyleFontFamily = paragraphStyleRunFormatting?.fontFamily;
  if (!paragraphStyleFontFamily) {
    return cascade;
  }
  return {
    ...cascade,
    formatting: mergeTextFormatting(cascade.formatting, {
      fontFamily: paragraphStyleFontFamily,
    }),
  };
};

type ResolveParagraphRunStyleFormattingOptions = {
  styleId: string | undefined;
  styleResolver: RunStyleResolver | null | undefined;
  tableRunFormatting: TextFormatting | undefined;
  inheritableParagraphMarkFormatting: TextFormatting | undefined;
};

export const resolveParagraphRunStyleFormatting = ({
  styleId,
  styleResolver,
  tableRunFormatting,
  inheritableParagraphMarkFormatting,
}: ResolveParagraphRunStyleFormattingOptions) => {
  const baseStyleCascade = resolveParagraphRunStyleBase(styleId, styleResolver, tableRunFormatting);
  const defaultCharacterFormatting = styleResolver?.getDefaultCharacterStyle()?.rPr;
  const baseWithDefaultCharacter = mergeTextFormatting(
    defaultCharacterFormatting,
    baseStyleCascade.formatting,
  );
  const defaultCharacterStyleCascade = cascadeStyleTextFormatting(
    [
      { cascade: baseStyleCascade, type: "carried" },
      { formatting: defaultCharacterFormatting, type: "style" },
    ],
    { ordinaryFormatting: baseWithDefaultCharacter },
  );
  const defaultRunFormatting = mergeTextFormatting(
    baseWithDefaultCharacter,
    inheritableParagraphMarkFormatting,
  );
  const defaultRunCascade = cascadeStyleTextFormatting(
    [
      { cascade: defaultCharacterStyleCascade, type: "carried" },
      { formatting: inheritableParagraphMarkFormatting, type: "direct" },
    ],
    { ordinaryFormatting: defaultRunFormatting },
  );
  return {
    baseStyleCascade,
    baseWithDefaultCharacter,
    defaultCharacterStyleCascade,
    defaultRunCascade,
  };
};

type ResolveEffectiveParagraphMarkFormattingOptions = {
  authored: ParagraphMarkProperties;
  styleId: string | undefined;
  styleResolver: RunStyleResolver | null | undefined;
  tableRunFormatting: TextFormatting | undefined;
};

/** Resolve authored paragraph-mark rPr over the complete inherited run cascade. */
export const resolveEffectiveParagraphMarkFormatting = ({
  authored,
  styleId,
  styleResolver,
  tableRunFormatting,
}: ResolveEffectiveParagraphMarkFormattingOptions): TextFormatting | undefined => {
  const characterStyleFormatting = authored.runProperties?.styleId
    ? styleResolver?.getRunStyleOwnProperties(authored.runProperties.styleId)
    : undefined;
  const resolvedParagraphMarkFormatting = cascadeStyleTextFormatting([
    { formatting: characterStyleFormatting, type: "style" },
    { formatting: authored.runProperties, type: "direct" },
  ]).formatting;
  const inheritableParagraphMarkFormatting = resolvedParagraphMarkFormatting
    ? stripParagraphMarkFormattingForBodyRuns(resolvedParagraphMarkFormatting)
    : undefined;
  return resolveParagraphRunStyleFormatting({
    styleId,
    styleResolver,
    tableRunFormatting,
    inheritableParagraphMarkFormatting,
  }).defaultRunCascade.formatting;
};

export const paragraphRunStyleContext = (
  paragraph: PMNode,
  styleResolver?: RunStyleResolver | null,
  tableRunFormatting?: TextFormatting | null,
): ParagraphRunStyleContext => {
  const attrs = expectParagraphAttrs(paragraph);
  const propertyState = expectParagraphPropertyState(attrs._paragraphPropertyState);
  const paragraphMarkFormatting = propertyState.context.paragraphMark.authored.runProperties;
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
  const tableStyleContext =
    tableRunFormatting === undefined
      ? undefined
      : resolveParagraphRunStyleFormatting({
          styleId: attrs.styleId,
          styleResolver,
          tableRunFormatting: tableRunFormatting ?? undefined,
          inheritableParagraphMarkFormatting:
            attrs.styleId === undefined && attrs._tableOfContentsLevel === undefined
              ? inheritableParagraphMarkFormatting
              : undefined,
        });
  return {
    baseParagraphFormatting:
      tableStyleContext === undefined
        ? resolveParagraphRunStyleBase(attrs.styleId, styleResolver).formatting
        : tableStyleContext.baseStyleCascade.formatting,
    paragraphFormatting:
      propertyState.context.paragraphMark.effective.defaultTextFormatting ?? undefined,
    paragraphMarkFormatting: inheritableParagraphMarkFormatting,
    paragraphMarkPrecedesStyle: attrs.styleId !== undefined,
  };
};

type ParagraphFormattingForRunOptions = {
  context: ParagraphRunStyleContext;
  directFormatting?: TextFormatting;
  marks: readonly Mark[];
};

export const paragraphFormattingForRun = ({
  context,
  directFormatting,
  marks,
}: ParagraphFormattingForRunOptions): TextFormatting | undefined => {
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

type ParagraphRunStyleContextAtOptions = {
  doc: PMNode;
  pos: number;
  styleResolver?: RunStyleResolver | null;
};

export const paragraphRunStyleContextAt = ({
  doc,
  pos,
  styleResolver,
}: ParagraphRunStyleContextAtOptions): ParagraphRunStyleContext => {
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
