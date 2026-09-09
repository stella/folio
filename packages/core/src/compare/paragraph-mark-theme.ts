import type { FolioDocumentStoryHandle } from "../ai-edits/headless";
import { paragraphMarkRunFormattingOf } from "../ai-edits/snapshot";
import type { FolioAIEditSnapshot } from "../ai-edits/types";
import { resolveThemeFontRef } from "../docx/themeParser";
import { createStyleEngine, type StyleEngine } from "../style-engine/styleEngine";
import type {
  ColorValue,
  StyleDefinitions,
  TextFormatting,
  Theme,
} from "../types/document";
import { canonicalJson } from "../utils/canonicalJson";
import { resolveColor } from "../utils/colorResolver";

type ParagraphMarkThemePair = {
  baseStory: FolioDocumentStoryHandle;
  targetSnapshot: FolioAIEditSnapshot;
};

export type ParagraphMarkPackageDependencyFailure = {
  story: FolioDocumentStoryHandle;
  detail: string;
};

const referencedColors = (formatting: TextFormatting): readonly ColorValue[] =>
  [
    formatting.color,
    formatting.underline?.color,
    formatting.shading?.color,
    formatting.shading?.fill,
  ].filter((color): color is ColorValue => color?.themeColor !== undefined);

const referencedFonts = (formatting: TextFormatting): readonly string[] => {
  const fonts = formatting.fontFamily;
  if (!fonts) {
    return [];
  }
  return [fonts.asciiTheme, fonts.hAnsiTheme, fonts.eastAsiaTheme, fonts.csTheme].filter(
    (reference): reference is string => reference !== undefined,
  );
};

type TargetThemeDependenciesOptions = {
  formatting: TextFormatting;
  baseTheme: Theme | null | undefined;
  targetTheme: Theme | null | undefined;
};

const targetThemeDependenciesResolveInBase = ({
  formatting,
  baseTheme,
  targetTheme,
}: TargetThemeDependenciesOptions): boolean =>
  referencedColors(formatting).every(
    (color) => resolveColor(color, baseTheme) === resolveColor(color, targetTheme),
  ) &&
  referencedFonts(formatting).every(
    (reference) =>
      resolveThemeFontRef(baseTheme, reference) === resolveThemeFontRef(targetTheme, reference),
  );

type CharacterStyleDependencyOptions = {
  styleId: string;
  baseEngine: StyleEngine;
  targetEngine: StyleEngine;
  baseTheme: Theme | null | undefined;
  targetTheme: Theme | null | undefined;
};

const characterStyleDependencyIsTransportable = ({
  styleId,
  baseEngine,
  targetEngine,
  baseTheme,
  targetTheme,
}: CharacterStyleDependencyOptions): boolean => {
  const baseStyle = baseEngine.getStyle(styleId);
  const targetStyle = targetEngine.getStyle(styleId);
  if (!baseStyle || !targetStyle) {
    return baseStyle === targetStyle;
  }
  if (baseStyle.type !== "character" || targetStyle.type !== "character") {
    return false;
  }
  const baseFormatting = baseEngine.resolveRunStyle(styleId);
  const targetFormatting = targetEngine.resolveRunStyle(styleId);
  return (
    canonicalJson(baseFormatting ?? {}) === canonicalJson(targetFormatting ?? {}) &&
    (targetFormatting === undefined ||
      targetThemeDependenciesResolveInBase({
        formatting: targetFormatting,
        baseTheme,
        targetTheme,
      }))
  );
};

type ParagraphMarkPackageDependencyOptions = {
  pairs: readonly ParagraphMarkThemePair[];
  baseTheme: Theme | null | undefined;
  targetTheme: Theme | null | undefined;
  baseStyles: StyleDefinitions | undefined;
  targetStyles: StyleDefinitions | undefined;
};

/**
 * A comparison keeps the base package theme in both terminal views. Target
 * paragraph-mark references are writable only when that theme gives them the
 * same meaning as the target package's theme.
 */
export const paragraphMarkPackageDependencyFailures = ({
  pairs,
  baseTheme,
  targetTheme,
  baseStyles,
  targetStyles,
}: ParagraphMarkPackageDependencyOptions): readonly ParagraphMarkPackageDependencyFailure[] => {
  const failures: ParagraphMarkPackageDependencyFailure[] = [];
  let styleEngines: { base: StyleEngine; target: StyleEngine } | undefined;
  const getStyleEngines = (): { base: StyleEngine; target: StyleEngine } => {
    styleEngines ??= {
      base: createStyleEngine(baseStyles),
      target: createStyleEngine(targetStyles),
    };
    return styleEngines;
  };
  for (const { baseStory, targetSnapshot } of pairs) {
    const formatting = [...paragraphMarkRunFormattingOf(targetSnapshot).values()].flatMap(
      (projection) => (projection.formatting ? [projection.formatting] : []),
    );
    if (
      formatting.some(
        (properties) =>
          !targetThemeDependenciesResolveInBase({
            formatting: properties,
            baseTheme,
            targetTheme,
          }),
      )
    ) {
      failures.push({
        story: baseStory,
        detail:
          "target paragraph-mark theme references have different semantics in the base package",
      });
    }
    const characterStyleMismatch = formatting.some(({ styleId }) => {
      if (styleId === undefined) {
        return false;
      }
      const engines = getStyleEngines();
      return !characterStyleDependencyIsTransportable({
        styleId,
        baseEngine: engines.base,
        targetEngine: engines.target,
        baseTheme,
        targetTheme,
      });
    });
    if (characterStyleMismatch) {
      failures.push({
        story: baseStory,
        detail:
          "target paragraph-mark character styles have different semantics in the base package",
      });
    }
  }
  return failures;
};
