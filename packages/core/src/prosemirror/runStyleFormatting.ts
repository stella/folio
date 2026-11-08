import type { Mark, Node as PMNode } from "prosemirror-model";

import type { StyleEngine } from "../style-engine";
import type { TextFormatting } from "../types/document";
import {
  mergeTextFormatting,
  STYLE_TOGGLE_KEYS,
} from "../utils/textFormattingMerge";
import { expectCharacterStyleMarkAttrs, expectParagraphAttrs } from "./attrs";
import {
  cascadeStyleFromResolvedBase,
  cascadeStyleTextFormatting,
  type StyleToggleKey,
} from "./styles/styleToggleCascade";

export const hasDirectRunFormatting = (
  formatting: TextFormatting | undefined,
): boolean => {
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
    // The mark's own unmodelled markup belongs to the mark's `w:rPr` and to no
    // body run: a `w:bdr` written once on the paragraph mark would come back
    // on every run in the paragraph.
    preserved: _preserved,
    ...formattingForBody
  } = formatting;
  return Object.keys(formattingForBody).length > 0
    ? formattingForBody
    : undefined;
};

/**
 * What a paragraph lends its body runs. The paragraph mark's own run
 * properties format the mark alone and lend nothing.
 */
export type ParagraphRunStyleContext = {
  baseParagraphFormatting: TextFormatting | undefined;
  paragraphFormatting: TextFormatting | undefined;
};

export type RunStyleResolver = Pick<
  StyleEngine,
  | "builtInStyles"
  | "getDefaultCharacterStyle"
  | "getDefaultParagraphStyle"
  | "getDocDefaults"
  | "getRunStyleOwnProperties"
  | "getStyle"
  | "resolveParagraphStyle"
>;

const withoutFontFamily = (
  formatting: TextFormatting | undefined,
): TextFormatting | undefined => {
  if (!formatting?.fontFamily) {
    return formatting;
  }
  const { fontFamily: _fontFamily, ...withoutFont } = formatting;
  return Object.keys(withoutFont).length > 0 ? withoutFont : undefined;
};

const paragraphStyleFontFamily = (
  styleId: string | undefined,
  styleResolver: RunStyleResolver,
): TextFormatting["fontFamily"] | undefined => {
  let style =
    (styleId ? styleResolver.getStyle(styleId) : undefined) ??
    styleResolver.getDefaultParagraphStyle();
  const visited = new Set<string>();
  const styleChain: TextFormatting[] = [];
  while (style?.type === "paragraph" && !visited.has(style.styleId)) {
    visited.add(style.styleId);
    if (style.rPr?.fontFamily) {
      styleChain.push({ fontFamily: style.rPr.fontFamily });
    }
    style = style.basedOn ? styleResolver.getStyle(style.basedOn) : undefined;
  }
  let formatting: TextFormatting | undefined;
  for (const styleFormatting of styleChain.toReversed()) {
    formatting = mergeTextFormatting(formatting, styleFormatting);
  }
  return formatting?.fontFamily;
};

type ResolveParagraphBodyRunFormattingOptions = {
  styleId: string | undefined;
  styleResolver?: RunStyleResolver | null | undefined;
  tableRunFormatting?: TextFormatting | undefined;
};

/** Resolve the table, paragraph, and character style cascade before paragraph-mark properties. */
export const resolveParagraphBodyRunFormatting = ({
  styleId,
  styleResolver,
  tableRunFormatting,
}: ResolveParagraphBodyRunFormattingOptions): {
  baseFormatting: TextFormatting | undefined;
  baseToggleCascade: ReturnType<typeof cascadeStyleTextFormatting>;
  defaultFormatting: TextFormatting | undefined;
  defaultToggleCascade: ReturnType<typeof cascadeStyleTextFormatting>;
} => {
  if (!styleResolver) {
    const defaultToggleCascade = cascadeStyleTextFormatting(
      [{ formatting: tableRunFormatting, type: "style" }],
      { ordinaryFormatting: tableRunFormatting },
    );
    return {
      baseFormatting: tableRunFormatting,
      baseToggleCascade: defaultToggleCascade,
      defaultFormatting: tableRunFormatting,
      defaultToggleCascade,
    };
  }
  const resolved = styleResolver.resolveParagraphStyle(styleId);
  const styleRunFormatting =
    tableRunFormatting === undefined
      ? resolved.runFormatting
      : withoutFontFamily(resolved.runFormatting);
  const paragraphStyle = styleId
    ? (styleResolver.getStyle(styleId) ??
      styleResolver.getDefaultParagraphStyle())
    : styleResolver.getDefaultParagraphStyle();
  const paragraphStyleRunFormatting =
    paragraphStyle?.type === "paragraph" ? paragraphStyle.rPr : undefined;
  const ordinaryStyleFormatting =
    styleId === undefined
      ? mergeTextFormatting(styleRunFormatting, tableRunFormatting)
      : mergeTextFormatting(tableRunFormatting, styleRunFormatting);
  const orderedToggleFormatting = cascadeStyleTextFormatting(
    [
      { formatting: styleResolver.getDocDefaults()?.rPr, type: "defaults" },
      { formatting: tableRunFormatting, type: "style" },
      { formatting: paragraphStyleRunFormatting, type: "style" },
    ],
    { ordinaryFormatting: ordinaryStyleFormatting },
  );
  let baseRunFormatting = orderedToggleFormatting.formatting;
  const fontFamily = paragraphStyleFontFamily(styleId, styleResolver);
  if (fontFamily) {
    baseRunFormatting = mergeTextFormatting(baseRunFormatting, { fontFamily });
  }
  const defaultCharacterFormatting =
    styleResolver.getDefaultCharacterStyle()?.rPr;
  const defaultFormatting = mergeTextFormatting(
    defaultCharacterFormatting,
    baseRunFormatting,
  );
  const defaultToggleCascade = cascadeStyleTextFormatting(
    [
      { cascade: orderedToggleFormatting, type: "carried" },
      { formatting: defaultCharacterFormatting, type: "style" },
    ],
    { ordinaryFormatting: defaultFormatting },
  );
  return {
    baseFormatting: baseRunFormatting,
    baseToggleCascade: orderedToggleFormatting,
    defaultFormatting,
    defaultToggleCascade,
  };
};

/**
 * The formatting a paragraph's runs inherit, with the toggles of the style
 * cascade resolved as they stand: what a run that states nothing shows.
 */
export const resolveParagraphInheritedRunFormatting = (
  options: ResolveParagraphBodyRunFormattingOptions,
) => {
  const body = resolveParagraphBodyRunFormatting(options);
  const inheritedToggleCascade = cascadeStyleTextFormatting(
    [
      { cascade: body.defaultToggleCascade, type: "carried" },
      { formatting: undefined, type: "direct" },
    ],
    { ordinaryFormatting: body.defaultFormatting },
  );
  return {
    ...body,
    inheritedFormatting: inheritedToggleCascade.formatting,
    inheritedToggleCascade,
  };
};

export const paragraphRunStyleContext = (
  paragraph: PMNode,
  styleResolver?: RunStyleResolver | null,
): ParagraphRunStyleContext => {
  const attrs = expectParagraphAttrs(paragraph);
  return {
    baseParagraphFormatting:
      attrs._tableRunFormatting === undefined
        ? styleResolver?.resolveParagraphStyle(attrs.styleId).runFormatting
        : resolveParagraphBodyRunFormatting({
            styleId: attrs.styleId,
            tableRunFormatting: attrs._tableRunFormatting,
            styleResolver,
          }).defaultFormatting,
    paragraphFormatting: attrs.defaultTextFormatting ?? undefined,
  };
};

type ParagraphFormattingForRunOptions = {
  context: ParagraphRunStyleContext;
  marks: readonly Mark[];
};

/**
 * What a run inherits from its paragraph: the caret defaults for a run with
 * no stated formatting, the paragraph style cascade for one that states some.
 */
export const paragraphFormattingForRun = ({
  context,
  marks,
}: ParagraphFormattingForRunOptions): TextFormatting | undefined => {
  if (
    !marks.some(
      ({ type }) =>
        type.name === "runFormattingOverride" || type.name === "characterStyle",
    )
  ) {
    return context.paragraphFormatting;
  }
  return context.baseParagraphFormatting;
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
  return nearestParagraphRunStyleContext(null, styleResolver);
};

/**
 * {@link paragraphRunStyleContextAt} for a caller that already knows the
 * nearest paragraph, or that there is none. A document walk knows it for free,
 * where `doc.resolve` re-descends from the root once per run.
 */
export const nearestParagraphRunStyleContext = (
  paragraph: PMNode | null,
  styleResolver?: RunStyleResolver | null,
): ParagraphRunStyleContext =>
  paragraph
    ? paragraphRunStyleContext(paragraph, styleResolver)
    : {
        baseParagraphFormatting: undefined,
        paragraphFormatting: undefined,
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
  const characterStyle = marks.find(
    ({ type }) => type.name === "characterStyle",
  );
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
