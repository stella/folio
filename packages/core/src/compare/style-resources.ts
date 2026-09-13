import { createStyleResolver } from "../prosemirror/styles/styleResolver";
import type { ColorValue, Style, StyleDefinitions, TextFormatting, Theme } from "../types/document";
import { resolveColorToHex } from "../utils/colorResolver";
import { canonicalJson } from "../utils/canonicalJson";
import { resolveThemeFont } from "../utils/fontResolver";

export type ImportReferencedStyleDefinitionsOptions = {
  sourceStyles: StyleDefinitions | undefined;
  destinationStyles: StyleDefinitions | undefined;
  sourceTheme: Theme | undefined;
  destinationTheme: Theme | undefined;
  referencedStyleIds: readonly string[];
  reservedStyleIds?: readonly string[];
  materializeDefaultParagraphStyle?: boolean;
};

export type ImportReferencedStyleDefinitionsResult =
  | {
      status: "unchanged" | "imported";
      styles: StyleDefinitions | undefined;
      styleIdMap: ReadonlyMap<string, string>;
      importedStyleIds: readonly string[];
      defaultParagraphStyleId?: string;
    }
  | { status: "unalignable"; detail: string };

const styleDependencies = ({ basedOn, link, next }: StyleDefinitions["styles"][number]) =>
  [basedOn, link, next].filter((styleId): styleId is string => styleId !== undefined);

const themesMatch = (
  sourceTheme: Theme | undefined,
  destinationTheme: Theme | undefined,
): boolean => canonicalJson(sourceTheme) === canonicalJson(destinationTheme);

const hasThemeReference = (value: unknown): boolean => {
  if (Array.isArray(value)) return value.some(hasThemeReference);
  if (typeof value !== "object" || value === null) return false;
  return Object.entries(value).some(
    ([key, nested]) => key.toLowerCase().includes("theme") || hasThemeReference(nested),
  );
};

const materializeThemeColor = (
  color: ColorValue | undefined,
  theme: Theme | undefined,
): ColorValue | undefined => {
  if (!color || !color.themeColor) return color;
  const rgb = resolveColorToHex(color, theme);
  if (!rgb) return undefined;
  const {
    themeColor: _themeColor,
    themeTint: _themeTint,
    themeShade: _themeShade,
    ...plain
  } = color;
  return { ...plain, rgb };
};

type ThemeFontSlot = "ascii" | "hAnsi" | "eastAsia" | "cs";

const languageForThemeFontSlot = ({
  formatting,
  slot,
}: {
  formatting: TextFormatting;
  slot: ThemeFontSlot;
}): string | undefined => {
  if (slot === "eastAsia") return formatting.language?.eastAsia ?? formatting.language?.val;
  if (slot === "cs") return formatting.language?.bidi ?? formatting.language?.val;
  return formatting.language?.val;
};

const scriptForLanguage = (language: string | undefined): string | undefined => {
  if (!language || typeof Intl.Locale !== "function") return undefined;
  try {
    return new Intl.Locale(language).maximize().script;
  } catch {
    return undefined;
  }
};

const resolveThemeFontForFormatting = ({
  formatting,
  slot,
  theme,
  themeRef,
}: {
  formatting: TextFormatting;
  slot: ThemeFontSlot;
  theme: Theme | undefined;
  themeRef: string;
}): string | null => {
  const script = scriptForLanguage(languageForThemeFontSlot({ formatting, slot }));
  const themeFont = themeRef.toLowerCase().startsWith("major")
    ? theme?.fontScheme?.majorFont
    : theme?.fontScheme?.minorFont;
  const scriptFont = script ? themeFont?.fonts?.[script] : undefined;
  return scriptFont || resolveThemeFont(themeRef, theme?.fontScheme);
};

type MaterializedTextFormattingResult =
  | { status: "ok"; formatting: TextFormatting | undefined }
  | { status: "unalignable" };

const materializeThemeTextFormatting = (
  formatting: TextFormatting | undefined,
  theme: Theme | undefined,
  materializeThemeReferences: boolean,
): MaterializedTextFormattingResult => {
  if (!formatting) return { status: "ok", formatting };
  if (!materializeThemeReferences) return { status: "ok", formatting };
  if (hasThemeReference(formatting.shading)) return { status: "unalignable" };
  const color = materializeThemeColor(formatting.color, theme);
  if (formatting.color?.themeColor && !color) return { status: "unalignable" };
  const underlineColor = materializeThemeColor(formatting.underline?.color, theme);
  if (formatting.underline?.color?.themeColor && !underlineColor) {
    return { status: "unalignable" };
  }
  const {
    asciiTheme: _asciiTheme,
    hAnsiTheme: _hAnsiTheme,
    eastAsiaTheme: _eastAsiaTheme,
    csTheme: _csTheme,
    ...fontFamily
  } = formatting.fontFamily ?? {};
  const themeFonts = formatting.fontFamily;
  const resolvedAscii = themeFonts?.asciiTheme
    ? resolveThemeFontForFormatting({
        formatting,
        slot: "ascii",
        theme,
        themeRef: themeFonts.asciiTheme,
      })
    : null;
  const resolvedHAnsi = themeFonts?.hAnsiTheme
    ? resolveThemeFontForFormatting({
        formatting,
        slot: "hAnsi",
        theme,
        themeRef: themeFonts.hAnsiTheme,
      })
    : null;
  const resolvedEastAsia = themeFonts?.eastAsiaTheme
    ? resolveThemeFontForFormatting({
        formatting,
        slot: "eastAsia",
        theme,
        themeRef: themeFonts.eastAsiaTheme,
      })
    : null;
  const resolvedCs = themeFonts?.csTheme
    ? resolveThemeFontForFormatting({
        formatting,
        slot: "cs",
        theme,
        themeRef: themeFonts.csTheme,
      })
    : null;
  if (
    (themeFonts?.asciiTheme && !resolvedAscii) ||
    (themeFonts?.hAnsiTheme && !resolvedHAnsi) ||
    (themeFonts?.eastAsiaTheme && !resolvedEastAsia) ||
    (themeFonts?.csTheme && !resolvedCs)
  ) {
    return { status: "unalignable" };
  }
  return {
    status: "ok",
    formatting: {
      ...formatting,
      ...(formatting.color && color && { color }),
      ...(formatting.underline && {
        underline: {
          ...formatting.underline,
          ...(underlineColor && { color: underlineColor }),
        },
      }),
      ...(formatting.fontFamily && {
        fontFamily: {
          ...fontFamily,
          ...(resolvedAscii && { ascii: resolvedAscii }),
          ...(resolvedHAnsi && { hAnsi: resolvedHAnsi }),
          ...(resolvedEastAsia && { eastAsia: resolvedEastAsia }),
          ...(resolvedCs && { cs: resolvedCs }),
        },
      }),
    },
  };
};

type MaterializedStyleResult = { status: "ok"; style: Style } | { status: "unalignable" };

const materializeEffectiveStyle = ({
  style,
  styles,
  theme,
  materializeThemeReferences,
}: {
  style: Style;
  styles: StyleDefinitions;
  theme: Theme | undefined;
  materializeThemeReferences: boolean;
}): MaterializedStyleResult => {
  if (style.type === "numbering") return { status: "unalignable" };
  const resolver = createStyleResolver(styles);
  const resolved =
    style.type === "character"
      ? { paragraphFormatting: undefined, runFormatting: resolver.resolveRunStyle(style.styleId) }
      : resolver.resolveParagraphStyle(style.styleId);
  if (materializeThemeReferences && hasThemeReference(resolved.paragraphFormatting)) {
    return { status: "unalignable" };
  }
  const runFormatting = materializeThemeTextFormatting(
    resolved.runFormatting,
    theme,
    materializeThemeReferences,
  );
  if (runFormatting.status === "unalignable") return runFormatting;
  const conditionalStyles: NonNullable<Style["tblStylePr"]> = [];
  for (const conditional of style.tblStylePr ?? []) {
    if (materializeThemeReferences && hasThemeReference(conditional.pPr)) {
      return { status: "unalignable" };
    }
    const conditionalRunFormatting = materializeThemeTextFormatting(
      conditional.rPr,
      theme,
      materializeThemeReferences,
    );
    if (conditionalRunFormatting.status === "unalignable") return { status: "unalignable" };
    conditionalStyles.push({
      ...conditional,
      ...(conditionalRunFormatting.formatting && { rPr: conditionalRunFormatting.formatting }),
    });
  }
  const { basedOn: _basedOn, ...independent } = style;
  return {
    status: "ok",
    style: {
      ...independent,
      ...((style.type === "paragraph" || style.type === "table") &&
        resolved.paragraphFormatting && {
          pPr: resolved.paragraphFormatting,
        }),
      ...(runFormatting.formatting && { rPr: runFormatting.formatting }),
      ...(style.tblStylePr && { tblStylePr: conditionalStyles }),
    },
  };
};

const materializeEffectiveDefaultParagraphStyle = ({
  styles,
  destinationStyles,
  theme,
  styleId,
  materializeThemeReferences,
}: {
  styles: StyleDefinitions;
  destinationStyles: StyleDefinitions;
  theme: Theme | undefined;
  styleId: string;
  materializeThemeReferences: boolean;
}): MaterializedStyleResult => {
  const resolved = createStyleResolver(styles).resolveParagraphStyle(undefined);
  const destinationResolved =
    createStyleResolver(destinationStyles).resolveParagraphStyle(undefined);
  if (materializeThemeReferences && hasThemeReference(resolved.paragraphFormatting)) {
    return { status: "unalignable" };
  }
  const runFormatting = materializeThemeTextFormatting(
    resolved.runFormatting,
    theme,
    materializeThemeReferences,
  );
  if (runFormatting.status === "unalignable") return runFormatting;
  const sourceParagraphFormatting = resolved.paragraphFormatting;
  const destinationParagraphFormatting = destinationResolved.paragraphFormatting;
  const materializedSpacing = {
    ...(sourceParagraphFormatting?.spaceBefore === undefined &&
      destinationParagraphFormatting?.spaceBefore !== undefined && { spaceBefore: 0 }),
    ...(sourceParagraphFormatting?.spaceAfter === undefined &&
      destinationParagraphFormatting?.spaceAfter !== undefined && { spaceAfter: 0 }),
    ...(sourceParagraphFormatting?.lineSpacing === undefined &&
      destinationParagraphFormatting?.lineSpacing !== undefined && { lineSpacing: 240 }),
    ...(sourceParagraphFormatting?.lineSpacingRule === undefined &&
      destinationParagraphFormatting?.lineSpacingRule !== undefined && {
        lineSpacingRule: "auto" as const,
      }),
    ...(sourceParagraphFormatting?.beforeAutospacing === undefined &&
      destinationParagraphFormatting?.beforeAutospacing !== undefined && {
        beforeAutospacing: false,
      }),
    ...(sourceParagraphFormatting?.afterAutospacing === undefined &&
      destinationParagraphFormatting?.afterAutospacing !== undefined && {
        afterAutospacing: false,
      }),
  };
  const paragraphFormatting =
    sourceParagraphFormatting || Object.keys(materializedSpacing).length > 0
      ? { ...sourceParagraphFormatting, ...materializedSpacing }
      : undefined;
  return {
    status: "ok",
    style: {
      styleId,
      type: "paragraph",
      ...(paragraphFormatting && { pPr: paragraphFormatting }),
      ...(runFormatting.formatting && { rPr: runFormatting.formatting }),
    },
  };
};

const effectiveDefaultParagraphFormattingMatches = ({
  sourceStyles,
  destinationStyles,
  sourceTheme,
  destinationTheme,
  materializeThemeReferences,
}: {
  sourceStyles: StyleDefinitions;
  destinationStyles: StyleDefinitions;
  sourceTheme: Theme | undefined;
  destinationTheme: Theme | undefined;
  materializeThemeReferences: boolean;
}): boolean => {
  const sourceResolved = createStyleResolver(sourceStyles).resolveParagraphStyle(undefined);
  const destinationResolved =
    createStyleResolver(destinationStyles).resolveParagraphStyle(undefined);
  const sourceRunFormatting = materializeThemeTextFormatting(
    sourceResolved.runFormatting,
    sourceTheme,
    materializeThemeReferences,
  );
  const destinationRunFormatting = materializeThemeTextFormatting(
    destinationResolved.runFormatting,
    destinationTheme,
    materializeThemeReferences,
  );
  if (
    sourceRunFormatting.status === "unalignable" ||
    destinationRunFormatting.status === "unalignable" ||
    (materializeThemeReferences && hasThemeReference(sourceResolved.paragraphFormatting)) ||
    (materializeThemeReferences && hasThemeReference(destinationResolved.paragraphFormatting))
  ) {
    return false;
  }
  return (
    canonicalJson({
      pPr: sourceResolved.paragraphFormatting,
      rPr: sourceRunFormatting.formatting,
    }) ===
    canonicalJson({
      pPr: destinationResolved.paragraphFormatting,
      rPr: destinationRunFormatting.formatting,
    })
  );
};

const nextAvailableStyleAlias = ({
  taken,
  sequence,
}: {
  taken: ReadonlySet<string>;
  sequence: number;
}): string => {
  let candidate = `FolioImportedStyle${sequence}`;
  let suffix = sequence;
  while (taken.has(candidate)) {
    suffix += 1;
    candidate = `FolioImportedStyle${suffix}`;
  }
  return candidate;
};

const remapImportedStyle = ({
  style,
  styleIdMap,
  isolatedDefaultStyleIds,
}: {
  style: Style;
  styleIdMap: ReadonlyMap<string, string>;
  isolatedDefaultStyleIds: ReadonlySet<string>;
}): Style => {
  const { default: _default, ...withoutDefault } = style;
  return {
    ...(isolatedDefaultStyleIds.has(style.styleId) ? withoutDefault : style),
    styleId: styleIdMap.get(style.styleId) ?? style.styleId,
    ...(style.basedOn !== undefined && { basedOn: styleIdMap.get(style.basedOn) ?? style.basedOn }),
    ...(style.link !== undefined && { link: styleIdMap.get(style.link) ?? style.link }),
    ...(style.next !== undefined && { next: styleIdMap.get(style.next) ?? style.next }),
  };
};

/**
 * Add the target style definitions required by a comparison without changing
 * existing package resources. A referenced style is only portable when both
 * packages have the same document defaults and theme: those resources take
 * part in OOXML style resolution even when a style definition itself is new.
 */
export const importReferencedStyleDefinitions = ({
  sourceStyles,
  destinationStyles,
  sourceTheme,
  destinationTheme,
  referencedStyleIds,
  reservedStyleIds = [],
  materializeDefaultParagraphStyle = false,
}: ImportReferencedStyleDefinitionsOptions): ImportReferencedStyleDefinitionsResult => {
  const requested = [...new Set(referencedStyleIds)].toSorted();
  if (requested.length === 0 && !materializeDefaultParagraphStyle) {
    return {
      status: "unchanged",
      styles: destinationStyles,
      styleIdMap: new Map(),
      importedStyleIds: [],
    };
  }
  if (!sourceStyles) {
    return { status: "unalignable", detail: "the target has no style definitions" };
  }
  if (!destinationStyles) {
    return { status: "unalignable", detail: "the base has no style definitions" };
  }

  const sourceById = new Map(sourceStyles.styles.map((style) => [style.styleId, style]));
  const destinationById = new Map(destinationStyles.styles.map((style) => [style.styleId, style]));
  const reserved = new Set(reservedStyleIds);
  const required = new Set<string>();
  const pending = [...requested];
  while (pending.length > 0) {
    const styleId = pending.pop();
    if (!styleId || required.has(styleId)) {
      continue;
    }
    const source = sourceById.get(styleId);
    if (!source) {
      return { status: "unalignable", detail: "a referenced target style is missing" };
    }
    required.add(styleId);
    pending.push(...styleDependencies(source));
  }

  const imported = [...required].toSorted();
  const taken = new Set([...sourceById.keys(), ...destinationById.keys(), ...reserved]);
  const styleIdMap = new Map<string, string>();
  const isolatedDefaultStyleIds = new Set<string>();
  let aliasSequence = 1;
  for (const styleId of imported) {
    const source = sourceById.get(styleId);
    const destination = destinationById.get(styleId);
    if (!source) {
      return { status: "unalignable", detail: "a referenced target style is missing" };
    }
    const conflictsWithBase =
      destination !== undefined && canonicalJson(source) !== canonicalJson(destination);
    if (conflictsWithBase || (!destination && (reserved.has(styleId) || source.default))) {
      const alias = nextAvailableStyleAlias({ taken, sequence: aliasSequence });
      aliasSequence += 1;
      taken.add(alias);
      styleIdMap.set(styleId, alias);
      if (source.default) isolatedDefaultStyleIds.add(styleId);
      continue;
    }
    styleIdMap.set(styleId, styleId);
  }

  const contextMatches =
    canonicalJson(sourceStyles.docDefaults) === canonicalJson(destinationStyles.docDefaults) &&
    themesMatch(sourceTheme, destinationTheme);
  const materializeThemeReferences = !themesMatch(sourceTheme, destinationTheme);
  const additions = sourceStyles.styles.filter(
    ({ styleId }) =>
      required.has(styleId) &&
      (!destinationById.has(styleId) || styleIdMap.get(styleId) !== styleId),
  );
  if (contextMatches) {
    if (additions.length === 0) {
      return { status: "unchanged", styles: destinationStyles, styleIdMap, importedStyleIds: [] };
    }
    return {
      status: "imported",
      styles: {
        ...destinationStyles,
        styles: [
          ...destinationStyles.styles,
          ...additions.map((style) =>
            remapImportedStyle({ style, styleIdMap, isolatedDefaultStyleIds }),
          ),
        ],
      },
      styleIdMap,
      importedStyleIds: additions.map(({ styleId }) => styleId),
    };
  }
  const materializedStyles: Style[] = [];
  for (const style of additions) {
    const result = materializeEffectiveStyle({
      style,
      styles: sourceStyles,
      theme: sourceTheme,
      materializeThemeReferences,
    });
    if (result.status === "unalignable") {
      return {
        status: "unalignable",
        detail: "a referenced style has formatting that cannot be materialized",
      };
    }
    materializedStyles.push(
      remapImportedStyle({ style: result.style, styleIdMap, isolatedDefaultStyleIds }),
    );
  }
  let defaultParagraphStyleId: string | undefined;
  let materializedDefaultParagraphStyle: Style | undefined;
  const needsDefaultParagraphStyle =
    materializeDefaultParagraphStyle &&
    !effectiveDefaultParagraphFormattingMatches({
      sourceStyles,
      destinationStyles,
      sourceTheme,
      destinationTheme,
      materializeThemeReferences,
    });
  if (needsDefaultParagraphStyle) {
    defaultParagraphStyleId = nextAvailableStyleAlias({ taken, sequence: aliasSequence });
    const result = materializeEffectiveDefaultParagraphStyle({
      styles: sourceStyles,
      destinationStyles,
      theme: sourceTheme,
      styleId: defaultParagraphStyleId,
      materializeThemeReferences,
    });
    if (result.status === "unalignable") {
      return {
        status: "unalignable",
        detail: "default paragraph formatting cannot be isolated from the base formatting context",
      };
    }
    materializedDefaultParagraphStyle = result.style;
    materializedStyles.push(result.style);
  }
  if (materializedStyles.length === 0) {
    return { status: "unchanged", styles: destinationStyles, styleIdMap, importedStyleIds: [] };
  }
  const candidateStyles = {
    ...destinationStyles,
    styles: [...destinationStyles.styles, ...materializedStyles],
  };
  const sourceResolver = createStyleResolver(sourceStyles);
  const candidateResolver = createStyleResolver(candidateStyles);
  for (const style of additions) {
    const candidateStyleId = styleIdMap.get(style.styleId) ?? style.styleId;
    const sourceResolved =
      style.type === "character"
        ? {
            paragraphFormatting: undefined,
            runFormatting: sourceResolver.resolveRunStyle(style.styleId),
          }
        : sourceResolver.resolveParagraphStyle(style.styleId);
    const candidateResolved =
      style.type === "character"
        ? {
            paragraphFormatting: undefined,
            runFormatting: candidateResolver.resolveRunStyle(candidateStyleId),
          }
        : candidateResolver.resolveParagraphStyle(candidateStyleId);
    const sourceRunFormatting = materializeThemeTextFormatting(
      sourceResolved.runFormatting,
      sourceTheme,
      materializeThemeReferences,
    );
    const candidateRunFormatting = materializeThemeTextFormatting(
      candidateResolved.runFormatting,
      destinationTheme,
      materializeThemeReferences,
    );
    if (
      sourceRunFormatting.status === "unalignable" ||
      candidateRunFormatting.status === "unalignable" ||
      (materializeThemeReferences && hasThemeReference(sourceResolved.paragraphFormatting)) ||
      (materializeThemeReferences && hasThemeReference(candidateResolved.paragraphFormatting)) ||
      canonicalJson({
        pPr: sourceResolved.paragraphFormatting,
        rPr: sourceRunFormatting.formatting,
      }) !==
        canonicalJson({
          pPr: candidateResolved.paragraphFormatting,
          rPr: candidateRunFormatting.formatting,
        })
    ) {
      return {
        status: "unalignable",
        detail: "a referenced style differs after materializing its formatting",
      };
    }
  }
  if (defaultParagraphStyleId !== undefined) {
    const sourceResolved = {
      paragraphFormatting: materializedDefaultParagraphStyle?.pPr,
      runFormatting: materializedDefaultParagraphStyle?.rPr,
    };
    const candidateResolved = candidateResolver.resolveParagraphStyle(defaultParagraphStyleId);
    const sourceRunFormatting = { status: "ok" as const, formatting: sourceResolved.runFormatting };
    const candidateRunFormatting = materializeThemeTextFormatting(
      candidateResolved.runFormatting,
      destinationTheme,
      materializeThemeReferences,
    );
    if (
      candidateRunFormatting.status === "unalignable" ||
      (materializeThemeReferences && hasThemeReference(candidateResolved.paragraphFormatting)) ||
      canonicalJson({
        pPr: sourceResolved.paragraphFormatting,
        rPr: sourceRunFormatting.formatting,
      }) !==
        canonicalJson({
          pPr: candidateResolved.paragraphFormatting,
          rPr: candidateRunFormatting.formatting,
        })
    ) {
      return {
        status: "unalignable",
        detail: "default paragraph formatting cannot be isolated from the base formatting context",
      };
    }
  }
  return {
    status: "imported",
    styles: candidateStyles,
    styleIdMap,
    importedStyleIds: materializedStyles.map(({ styleId }) => styleId),
    ...(defaultParagraphStyleId !== undefined && { defaultParagraphStyleId }),
  };
};
