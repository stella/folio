import type { Paragraph, TextFormatting } from "../types/document";
import { mergeTextFormatting } from "../utils/textFormattingMerge";
import {
  getParagraphMarkSuppressionOverrides,
  hasDirectRunFormatting,
  stripParagraphMarkFormattingForBodyRuns,
  suppressParagraphMarkFormatting,
} from "../prosemirror/runStyleFormatting";
import { cascadeStyleTextFormatting } from "../prosemirror/styles/styleToggleCascade";
import type { StyleEngine } from "./styleEngine";

export type ParagraphDefaultFormattingResolver = Pick<
  StyleEngine,
  | "getStyle"
  | "getDocDefaults"
  | "getDefaultParagraphStyle"
  | "getDefaultCharacterStyle"
  | "getRunStyleOwnProperties"
  | "resolveParagraphStyle"
>;

export type ResolvedRunFormatting = {
  readonly formatting: TextFormatting | undefined;
  readonly implicitCharacterStyleApplied?: true;
  readonly paragraphMarkOverrides?: TextFormatting;
  readonly toggleCascade: ReturnType<typeof cascadeStyleTextFormatting>;
};

export type ResolvedRunPresentation = {
  readonly effective: TextFormatting | undefined;
  readonly inherited: TextFormatting | undefined;
};

/** Resolve an embedded character-style reference without importing doc defaults. */
export const resolveRunFormattingWithoutDefaults = (
  formatting: TextFormatting | undefined,
  styleResolver: Pick<StyleEngine, "getRunStyleOwnProperties"> | null,
): TextFormatting | undefined => {
  if (!formatting || !styleResolver) return formatting;
  const characterStyleFormatting = formatting.styleId
    ? styleResolver.getRunStyleOwnProperties(formatting.styleId)
    : undefined;
  return cascadeStyleTextFormatting([
    { formatting: characterStyleFormatting, type: "style" },
    { formatting, type: "direct" },
  ]).formatting;
};

const resolveParagraphStyleFontFamily = (
  styleId: string | undefined,
  styleResolver: ParagraphDefaultFormattingResolver,
): TextFormatting["fontFamily"] | undefined => {
  let style = styleId ? styleResolver.getStyle(styleId) : styleResolver.getDefaultParagraphStyle();
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

const withoutFontFamily = (formatting: TextFormatting | undefined): TextFormatting | undefined => {
  if (!formatting?.fontFamily) return formatting;
  const { fontFamily: _fontFamily, ...withoutFont } = formatting;
  return Object.keys(withoutFont).length > 0 ? withoutFont : undefined;
};

type ParagraphRunFormattingResolution = {
  readonly defaultFormatting: TextFormatting | undefined;
  readonly resolve: (
    formatting: TextFormatting | undefined,
    fieldType?: string,
  ) => ResolvedRunFormatting;
};

/**
 * The single live-model cascade for paragraph-inherited run presentation.
 * Import and comparison call this same resolver; no PM cache is an authority.
 */
export const createParagraphRunFormattingResolver = ({
  paragraph,
  styleResolver,
  extraRunFormatting,
  isTocParagraph,
}: {
  readonly paragraph: Paragraph;
  readonly styleResolver: ParagraphDefaultFormattingResolver | null;
  readonly extraRunFormatting?: TextFormatting;
  readonly isTocParagraph: boolean;
}): ParagraphRunFormattingResolution => {
  let styleRunFormatting: TextFormatting | undefined;
  let paragraphStyleRunFormatting: TextFormatting | undefined;
  let paragraphStyleFontFamily: TextFormatting["fontFamily"] | undefined;
  if (styleResolver) {
    const resolved = styleResolver.resolveParagraphStyle(paragraph.formatting?.styleId);
    // Table style run defaults outrank document-default fonts. Remove those
    // defaults here when a table contribution exists; explicit paragraph-style
    // font slots are restored from their authored chain below.
    styleRunFormatting =
      extraRunFormatting === undefined
        ? resolved.runFormatting
        : withoutFontFamily(resolved.runFormatting);
    const paragraphStyle = paragraph.formatting?.styleId
      ? (styleResolver.getStyle(paragraph.formatting.styleId) ??
        styleResolver.getDefaultParagraphStyle())
      : styleResolver.getDefaultParagraphStyle();
    paragraphStyleRunFormatting =
      paragraphStyle?.type === "paragraph" ? paragraphStyle.rPr : undefined;
    paragraphStyleFontFamily = resolveParagraphStyleFontFamily(
      paragraph.formatting?.styleId,
      styleResolver,
    );
  }
  const paragraphRunFormatting = resolveRunFormattingWithoutDefaults(
    paragraph.formatting?.runProperties,
    styleResolver,
  );
  let inheritableParagraphRunFormatting: TextFormatting | undefined;
  if (paragraphRunFormatting && !isTocParagraph && paragraph.formatting?.styleId === undefined) {
    inheritableParagraphRunFormatting =
      stripParagraphMarkFormattingForBodyRuns(paragraphRunFormatting);
  }
  const ordinaryStyleFormatting =
    paragraph.formatting?.styleId === undefined
      ? mergeTextFormatting(styleRunFormatting, extraRunFormatting)
      : mergeTextFormatting(extraRunFormatting, styleRunFormatting);
  const orderedToggleFormatting = cascadeStyleTextFormatting(
    [
      { formatting: styleResolver?.getDocDefaults()?.rPr, type: "defaults" },
      { formatting: extraRunFormatting, type: "style" },
      { formatting: paragraphStyleRunFormatting, type: "style" },
    ],
    { ordinaryFormatting: ordinaryStyleFormatting },
  );
  let baseRunFormatting = orderedToggleFormatting.formatting;
  if (paragraphStyleFontFamily) {
    baseRunFormatting = mergeTextFormatting(baseRunFormatting, {
      fontFamily: paragraphStyleFontFamily,
    });
  }
  const defaultCharacterFormatting = styleResolver?.getDefaultCharacterStyle()?.rPr;
  const ordinaryBaseWithDefaultCharacter = mergeTextFormatting(
    defaultCharacterFormatting,
    baseRunFormatting,
  );
  const defaultCharacterStyleCascade = cascadeStyleTextFormatting(
    [
      { cascade: orderedToggleFormatting, type: "carried" },
      { formatting: defaultCharacterFormatting, type: "style" },
    ],
    { ordinaryFormatting: ordinaryBaseWithDefaultCharacter },
  );
  const ordinaryDefaultRunFormatting = mergeTextFormatting(
    ordinaryBaseWithDefaultCharacter,
    inheritableParagraphRunFormatting,
  );
  const defaultToggleCascade = cascadeStyleTextFormatting(
    [
      { cascade: defaultCharacterStyleCascade, type: "carried" },
      { formatting: inheritableParagraphRunFormatting, type: "direct" },
    ],
    { ordinaryFormatting: ordinaryDefaultRunFormatting },
  );
  const defaultFormatting = defaultToggleCascade.formatting;
  return {
    defaultFormatting,
    resolve: (formatting, fieldType) => {
      const hasCharacterStyle = formatting?.styleId !== undefined;
      const inheritedBaseFormatting = hasCharacterStyle
        ? baseRunFormatting
        : ordinaryBaseWithDefaultCharacter;
      const inheritedToggleCascade = hasCharacterStyle
        ? orderedToggleFormatting
        : defaultCharacterStyleCascade;
      if (fieldType === "TOC") {
        return {
          formatting: hasDirectRunFormatting(formatting)
            ? suppressParagraphMarkFormatting({
                baseFormatting: inheritedBaseFormatting,
                directFormatting: formatting,
                paragraphMarkFormatting: undefined,
              })
            : inheritedBaseFormatting,
          ...(!hasCharacterStyle ? { implicitCharacterStyleApplied: true as const } : {}),
          toggleCascade: inheritedToggleCascade,
        };
      }
      const hasExplicitRunFormatting =
        hasDirectRunFormatting(formatting) || formatting?.styleId !== undefined;
      if (!hasExplicitRunFormatting) {
        return {
          formatting: defaultFormatting,
          implicitCharacterStyleApplied: true,
          toggleCascade: defaultToggleCascade,
        };
      }
      const suppressedFormatting = suppressParagraphMarkFormatting({
        baseFormatting: inheritedBaseFormatting,
        directFormatting: formatting,
        paragraphMarkFormatting: inheritableParagraphRunFormatting,
      });
      const paragraphMarkOverrides =
        getParagraphMarkSuppressionOverrides({
          directFormatting: formatting,
          paragraphMarkFormatting: inheritableParagraphRunFormatting,
          suppressedFormatting,
        }) ?? (inheritableParagraphRunFormatting ? {} : undefined);
      return {
        formatting: suppressedFormatting,
        ...(!hasCharacterStyle ? { implicitCharacterStyleApplied: true as const } : {}),
        ...(paragraphMarkOverrides ? { paragraphMarkOverrides } : {}),
        toggleCascade: inheritedToggleCascade,
      };
    },
  };
};

/** Exact final presentation for one authored run under the shared inherited cascade. */
export const resolveEffectiveRunPresentation = (
  runFormatting: TextFormatting | undefined,
  inherited: ResolvedRunFormatting,
  styleResolver: ParagraphDefaultFormattingResolver | null,
): ResolvedRunPresentation => {
  const styleId = runFormatting?.styleId;
  const characterStyleFormatting = (() => {
    if (styleId) return styleResolver?.getRunStyleOwnProperties(styleId);
    if (!inherited.implicitCharacterStyleApplied) {
      return styleResolver?.getDefaultCharacterStyle()?.rPr;
    }
    return undefined;
  })();
  let ordinaryRunStyleFormatting = inherited.formatting ? { ...inherited.formatting } : {};
  if (styleId) {
    ordinaryRunStyleFormatting =
      mergeTextFormatting(inherited.formatting, characterStyleFormatting) ?? {};
  }
  const cascadedStyleFormatting = cascadeStyleTextFormatting(
    [
      { cascade: inherited.toggleCascade, type: "carried" },
      { formatting: characterStyleFormatting, type: "style" },
    ],
    { ordinaryFormatting: ordinaryRunStyleFormatting },
  );
  const runStyleFormatting = cascadedStyleFormatting.formatting;
  const effective = cascadeStyleTextFormatting(
    [
      { cascade: cascadedStyleFormatting, type: "carried" },
      { formatting: runFormatting, type: "direct" },
    ],
    { ordinaryFormatting: mergeTextFormatting(runStyleFormatting, runFormatting) },
  ).formatting;
  return { effective, inherited: runStyleFormatting };
};
