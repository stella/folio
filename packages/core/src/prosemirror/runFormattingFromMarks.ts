import type { Mark } from "prosemirror-model";
import { normalizeHorizontalScalePercent } from "../utils/horizontalScale";
import type { TextFormatting, ColorValue } from "../types/document";
import {
  expectCharacterSpacingMarkAttrs,
  expectCharacterStyleMarkAttrs,
  expectEmphasisMarkAttrs,
  expectTextEffectMarkAttrs,
  expectFontFamilyMarkAttrs,
  expectLanguageMarkAttrs,
  expectFontSizeMarkAttrs,
  expectHighlightMarkAttrs,
  expectRunShadingMarkAttrs,
  expectRunFormattingOverrideMarkAttrs,
  expectStrikeMarkAttrs,
  expectTextColorMarkAttrs,
  expectUnderlineMarkAttrs,
} from "./attrs";
import {
  authoredRunFormattingFromAttrs,
  hasAuthoredRunFormattingProvenance,
} from "./runFormattingProvenance";
import {
  paragraphFormattingForRun,
  resolveEffectiveRunStyleFormatting,
  suppressParagraphMarkFormatting,
  type RunStyleResolver,
} from "./runStyleFormatting";
import {
  applyRunFormattingOverrideAttrs,
  buildRunFormattingOverrideAttrs,
} from "./extensions/marks/RunFormattingOverrideExtension";
import type { RunFormattingOverrideAttrs } from "./schema/marks";
import { runShadingAttrsToShading, shadingToRunShadingAttrs } from "./conversion/runShadingMark";
import { mergeTextFormatting } from "../utils/textFormattingMerge";

/**
 * Convert ProseMirror marks to TextFormatting
 */
export type MarksToTextFormattingOptions = {
  baseParagraphFormatting?: TextFormatting | undefined;
  inheritedFormatting?: TextFormatting | undefined;
  paragraphMarkFormatting?: TextFormatting | undefined;
  paragraphMarkPrecedesStyle?: boolean | undefined;
  styleResolver?: RunStyleResolver | null | undefined;
};

const RUN_FORMATTING_VISUAL_GROUPS = {
  bold: "bold",
  boldCs: null,
  italic: "italic",
  italicCs: null,
  underline: "underline",
  strike: "strike",
  doubleStrike: "strike",
  vertAlign: "vertAlign",
  smallCaps: "smallCaps",
  allCaps: "allCaps",
  hidden: "hidden",
  noProof: null,
  color: "color",
  highlight: "highlight",
  shading: "shading",
  fontSize: "fontSize",
  fontSizeCs: null,
  fontFamily: "fontFamily",
  language: "language",
  spacing: "characterSpacing",
  position: "characterSpacing",
  scale: "characterSpacing",
  kerning: "characterSpacing",
  effect: "effect",
  emphasisMark: "emphasisMark",
  emboss: "emboss",
  imprint: "imprint",
  outline: "outline",
  shadow: "shadow",
  rtl: "rtl",
  cs: null,
  styleId: null,
  preserved: null,
} as const satisfies Record<keyof TextFormatting, string | null>;

type VisualFormattingGroup = Exclude<
  (typeof RUN_FORMATTING_VISUAL_GROUPS)[keyof typeof RUN_FORMATTING_VISUAL_GROUPS],
  null
>;

const RUN_FORMATTING_FAST_PATH_DISPOSITION = {
  bold: "visual",
  boldCs: "structural",
  italic: "visual",
  italicCs: "structural",
  underline: "visual",
  strike: "visual",
  doubleStrike: "visual",
  vertAlign: "visual",
  smallCaps: "visual",
  allCaps: "visual",
  hidden: "visual",
  noProof: "structural",
  color: "visual",
  highlight: "visual",
  shading: "visual",
  fontSize: "visual",
  fontSizeCs: "structural",
  fontFamily: "visual",
  language: "visual",
  spacing: "visual",
  position: "visual",
  scale: "visual",
  kerning: "visual",
  effect: "visual",
  emphasisMark: "visual",
  emboss: "visual",
  imprint: "visual",
  outline: "visual",
  shadow: "visual",
  rtl: "visual",
  cs: "structural",
  styleId: "character-style",
  // Bytes, not a visual: a run holding one has formatting no mark carries, so
  // it must not take the fast path that rebuilds the run from marks alone.
  preserved: "structural",
} as const satisfies Record<keyof TextFormatting, "character-style" | "structural" | "visual">;

const hasOnlyCarrierlessVisualFormatting = (formatting: TextFormatting): boolean => {
  for (const property of Object.keys(formatting) as (keyof TextFormatting)[]) {
    if (RUN_FORMATTING_FAST_PATH_DISPOSITION[property] !== "visual") {
      return false;
    }
  }
  return true;
};

const visibleUnderline = (
  formatting: TextFormatting | undefined,
): TextFormatting["underline"] | undefined => {
  const underline = formatting?.underline;
  return underline?.style === "none" ? undefined : underline;
};

const visibleStrike = (formatting: TextFormatting | undefined): "double" | "single" | undefined => {
  if (formatting?.doubleStrike === true) {
    return "double";
  }
  return formatting?.strike === true ? "single" : undefined;
};

const visibleColor = (
  formatting: TextFormatting | undefined,
): TextFormatting["color"] | undefined =>
  formatting?.color?.auto === true ? undefined : formatting?.color;

const visibleValue = <Value>(value: Value | "none" | undefined): Value | undefined =>
  value === "none" ? undefined : value;

/**
 * Compare the exact presentation encoded by PM run marks without allocating the
 * intermediate maps used by reconciliation. Non-visual authored state is
 * deliberately absent here: the carrierless fast path below excludes every
 * structural carrier before calling this predicate.
 */
const sameVisualFormatting = (
  left: TextFormatting | undefined,
  right: TextFormatting | undefined,
): boolean => {
  if (
    (left?.bold === true) !== (right?.bold === true) ||
    (left?.italic === true) !== (right?.italic === true) ||
    !sameFormattingValue(visibleUnderline(left), visibleUnderline(right)) ||
    visibleStrike(left) !== visibleStrike(right) ||
    !sameFormattingValue(visibleColor(left), visibleColor(right)) ||
    visibleValue(left?.highlight) !== visibleValue(right?.highlight) ||
    (left?.fontSize || undefined) !== (right?.fontSize || undefined) ||
    !sameFormattingValue(left?.fontFamily, right?.fontFamily) ||
    !sameFormattingValue(left?.language, right?.language) ||
    (left?.vertAlign === "superscript" || left?.vertAlign === "subscript"
      ? left.vertAlign
      : undefined) !==
      (right?.vertAlign === "superscript" || right?.vertAlign === "subscript"
        ? right.vertAlign
        : undefined) ||
    (left?.allCaps === true) !== (right?.allCaps === true) ||
    (left?.smallCaps === true) !== (right?.smallCaps === true) ||
    (left?.hidden === true) !== (right?.hidden === true) ||
    (left?.emboss === true) !== (right?.emboss === true) ||
    (left?.imprint === true) !== (right?.imprint === true) ||
    (left?.shadow === true) !== (right?.shadow === true) ||
    (left?.outline === true) !== (right?.outline === true) ||
    (left?.rtl === true) !== (right?.rtl === true) ||
    (typeof left?.spacing === "number" ? left.spacing : null) !==
      (typeof right?.spacing === "number" ? right.spacing : null) ||
    (typeof left?.position === "number" ? left.position : null) !==
      (typeof right?.position === "number" ? right.position : null) ||
    (normalizeHorizontalScalePercent(left?.scale) ?? null) !==
      (normalizeHorizontalScalePercent(right?.scale) ?? null) ||
    (typeof left?.kerning === "number" ? left.kerning : null) !==
      (typeof right?.kerning === "number" ? right.kerning : null) ||
    visibleValue(left?.effect) !== visibleValue(right?.effect) ||
    visibleValue(left?.emphasisMark) !== visibleValue(right?.emphasisMark)
  ) {
    return false;
  }

  return sameFormattingValue(
    shadingToRunShadingAttrs(left?.shading),
    shadingToRunShadingAttrs(right?.shading),
  );
};

const visualFormattingGroups = (
  formatting: TextFormatting | undefined,
): ReadonlyMap<VisualFormattingGroup, unknown> => {
  const groups = new Map<VisualFormattingGroup, unknown>();
  if (!formatting) {
    return groups;
  }
  if (formatting.bold) {
    groups.set("bold", true);
  }
  if (formatting.italic) {
    groups.set("italic", true);
  }
  if (formatting.underline && formatting.underline.style !== "none") {
    groups.set("underline", {
      style: formatting.underline.style || "single",
      ...(formatting.underline.color ? { color: formatting.underline.color } : {}),
    });
  }
  if (formatting.strike || formatting.doubleStrike) {
    groups.set("strike", formatting.doubleStrike ? "double" : "single");
  }
  if (formatting.color && !formatting.color.auto) {
    const { rgb, themeColor, themeTint, themeShade } = formatting.color;
    groups.set("color", {
      ...(rgb ? { rgb } : {}),
      ...(themeColor ? { themeColor } : {}),
      ...(themeTint ? { themeTint } : {}),
      ...(themeShade ? { themeShade } : {}),
    });
  }
  if (formatting.highlight && formatting.highlight !== "none") {
    groups.set("highlight", formatting.highlight);
  }
  const shadingAttrs = shadingToRunShadingAttrs(formatting.shading);
  if (shadingAttrs) {
    groups.set("shading", runShadingAttrsToShading(shadingAttrs));
  }
  if (formatting.fontSize) {
    groups.set("fontSize", formatting.fontSize);
  }
  if (formatting.fontFamily) {
    const { ascii, hAnsi, eastAsia, cs, hint, asciiTheme, hAnsiTheme, eastAsiaTheme, csTheme } =
      formatting.fontFamily;
    groups.set("fontFamily", {
      ...(ascii ? { ascii } : {}),
      ...(hAnsi ? { hAnsi } : {}),
      ...(eastAsia ? { eastAsia } : {}),
      ...(cs ? { cs } : {}),
      ...(hint ? { hint } : {}),
      ...(asciiTheme ? { asciiTheme } : {}),
      ...(hAnsiTheme ? { hAnsiTheme } : {}),
      ...(eastAsiaTheme ? { eastAsiaTheme } : {}),
      ...(csTheme ? { csTheme } : {}),
    });
  }
  if (formatting.language) {
    const { val, eastAsia, bidi } = formatting.language;
    groups.set("language", {
      ...(val ? { val } : {}),
      ...(eastAsia ? { eastAsia } : {}),
      ...(bidi ? { bidi } : {}),
    });
  }
  if (formatting.vertAlign === "superscript" || formatting.vertAlign === "subscript") {
    groups.set("vertAlign", formatting.vertAlign);
  }
  for (const [property, group] of [
    ["allCaps", "allCaps"],
    ["smallCaps", "smallCaps"],
    ["hidden", "hidden"],
    ["emboss", "emboss"],
    ["imprint", "imprint"],
    ["shadow", "shadow"],
    ["outline", "outline"],
    ["rtl", "rtl"],
  ] as const) {
    if (formatting[property]) {
      groups.set(group, true);
    }
  }
  const spacing = typeof formatting.spacing === "number" ? formatting.spacing : null;
  const position = typeof formatting.position === "number" ? formatting.position : null;
  const scale = normalizeHorizontalScalePercent(formatting.scale) ?? null;
  const kerning = typeof formatting.kerning === "number" ? formatting.kerning : null;
  if (spacing !== null || position !== null || scale !== null || kerning !== null) {
    groups.set("characterSpacing", { spacing, position, scale, kerning });
  }
  if (formatting.effect && formatting.effect !== "none") {
    groups.set("effect", formatting.effect);
  }
  if (formatting.emphasisMark && formatting.emphasisMark !== "none") {
    groups.set("emphasisMark", formatting.emphasisMark);
  }
  return groups;
};

const expectedRunFormattingOverrideAttrs = (
  authoredFormatting: TextFormatting,
): RunFormattingOverrideAttrs | undefined => {
  const attrs = buildRunFormattingOverrideAttrs(authoredFormatting, {
    type: "authored-baseline",
    formatting: authoredFormatting,
  });
  const hasDirectFormatting = Object.keys(authoredFormatting).some(
    (property) => property !== "styleId",
  );
  if (!attrs && !hasDirectFormatting) {
    return undefined;
  }

  const expected: RunFormattingOverrideAttrs = { ...attrs };
  const directFontProperties = (["color", "fontFamily", "fontSize"] as const).filter(
    (property) => authoredFormatting[property] !== undefined,
  );
  if (directFontProperties.length > 0) {
    expected.directFontProperties = directFontProperties;
  }
  const complexScriptPropertyAbsences = (
    [
      ["bold", "boldCs"],
      ["italic", "italicCs"],
      ["fontSize", "fontSizeCs"],
    ] as const
  )
    .filter(
      ([ordinary, complex]) =>
        authoredFormatting[ordinary] !== undefined && authoredFormatting[complex] === undefined,
    )
    .map(([, complex]) => complex);
  if (complexScriptPropertyAbsences.length > 0) {
    expected.complexScriptPropertyAbsences = complexScriptPropertyAbsences;
  }
  return Object.keys(expected).length > 0 ? expected : undefined;
};

const changedVisualFormattingGroups = (
  observedFormatting: TextFormatting,
  expectedFormatting: TextFormatting | undefined,
): ReadonlySet<VisualFormattingGroup> => {
  const actual = visualFormattingGroups(observedFormatting);
  const expected = visualFormattingGroups(expectedFormatting);
  const changed = new Set<VisualFormattingGroup>();
  for (const group of new Set([...actual.keys(), ...expected.keys()])) {
    const actualValue = actual.get(group);
    const expectedValue = expected.get(group);
    if (!sameFormattingValue(actualValue, expectedValue)) {
      changed.add(group);
    }
  }
  return changed;
};

const VISUAL_BOOLEAN_FORMATTING_PROPERTIES = new Set<keyof TextFormatting>([
  "allCaps",
  "bold",
  "doubleStrike",
  "emboss",
  "hidden",
  "imprint",
  "italic",
  "outline",
  "rtl",
  "shadow",
  "smallCaps",
  "strike",
]);

const fontFamilyDifference = (
  actual: NonNullable<TextFormatting["fontFamily"]>,
  inherited: TextFormatting["fontFamily"],
): NonNullable<TextFormatting["fontFamily"]> | undefined => {
  const difference: NonNullable<TextFormatting["fontFamily"]> = {};
  for (const property of [
    "ascii",
    "hAnsi",
    "eastAsia",
    "cs",
    "hint",
    "asciiTheme",
    "hAnsiTheme",
    "eastAsiaTheme",
    "csTheme",
  ] as const) {
    const value = actual[property];
    if (value !== undefined && value !== inherited?.[property]) {
      Reflect.set(difference, property, value);
    }
  }
  return Object.keys(difference).length > 0 ? difference : undefined;
};

type ReconcileAuthoredFormattingOptions = {
  authoredFormatting: TextFormatting;
  carrierlessContext: "ordinary" | "paragraph-mark";
  currentOverrideAttrs: RunFormattingOverrideAttrs | undefined;
  inheritedFormatting: TextFormatting | undefined;
  observedFormatting: TextFormatting;
};

const DIRECT_OVERRIDE_FORMATTING_PROPERTIES = [
  "allCaps",
  "bold",
  "boldCs",
  "cs",
  "doubleStrike",
  "emboss",
  "fontSizeCs",
  "hidden",
  "noProof",
  "imprint",
  "italic",
  "italicCs",
  "outline",
  "rtl",
  "shadow",
  "smallCaps",
  "strike",
  "underline",
] as const satisfies readonly (keyof TextFormatting & keyof RunFormattingOverrideAttrs)[];

const POSITIVE_OVERRIDE_PROPERTIES_REPRESENTED_BY_VISUAL_MARKS = new Set<
  (typeof DIRECT_OVERRIDE_FORMATTING_PROPERTIES)[number]
>(["allCaps", "emboss", "hidden", "imprint", "outline", "rtl", "shadow", "smallCaps", "strike"]);

const setFormattingFromOverrideSignal = (
  formatting: TextFormatting,
  property: (typeof DIRECT_OVERRIDE_FORMATTING_PROPERTIES)[number],
  value: RunFormattingOverrideAttrs[typeof property],
): void => {
  Reflect.deleteProperty(formatting, property);
  if (value === undefined) {
    return;
  }
  if (property === "underline") {
    formatting.underline = { style: "none" };
    return;
  }
  Reflect.set(formatting, property, value);
};

export const sameFormattingValue = (left: unknown, right: unknown): boolean => {
  if (left === right) {
    return true;
  }
  if (typeof left !== "object" || left === null || typeof right !== "object" || right === null) {
    return false;
  }
  if (Array.isArray(left) || Array.isArray(right)) {
    if (!Array.isArray(left) || !Array.isArray(right) || left.length !== right.length) {
      return false;
    }
    return left.every((value, index) => sameFormattingValue(value, right[index]));
  }
  const leftKeys = Object.keys(left);
  const rightKeys = Object.keys(right);
  if (leftKeys.length !== rightKeys.length) {
    return false;
  }
  return leftKeys.every(
    (key) =>
      Object.prototype.hasOwnProperty.call(right, key) &&
      sameFormattingValue(Reflect.get(left, key), Reflect.get(right, key)),
  );
};

const reconcileOverrideSignals = (
  formatting: TextFormatting,
  current: RunFormattingOverrideAttrs | undefined,
  expected: RunFormattingOverrideAttrs | undefined,
  inheritedFormatting: TextFormatting | undefined,
  observedFormatting: TextFormatting,
): void => {
  for (const property of DIRECT_OVERRIDE_FORMATTING_PROPERTIES) {
    if (sameFormattingValue(current?.[property], expected?.[property])) {
      continue;
    }
    if (
      current?.[property] === undefined &&
      expected?.[property] === true &&
      POSITIVE_OVERRIDE_PROPERTIES_REPRESENTED_BY_VISUAL_MARKS.has(property)
    ) {
      continue;
    }
    if (
      current &&
      hasAuthoredRunFormattingProvenance(current) &&
      expected?.[property] === undefined &&
      sameFormattingValue(current[property], inheritedFormatting?.[property])
    ) {
      continue;
    }
    setFormattingFromOverrideSignal(formatting, property, current?.[property]);
  }

  for (const property of ["color", "fontFamily", "fontSize"] as const) {
    const currentIsDirect = current?.directFontProperties?.includes(property) === true;
    const expectedIsDirect = expected?.directFontProperties?.includes(property) === true;
    if (currentIsDirect === expectedIsDirect) {
      continue;
    }
    if (!currentIsDirect) {
      Reflect.deleteProperty(formatting, property);
      continue;
    }
    const observed = observedFormatting[property];
    if (observed !== undefined) {
      Reflect.set(formatting, property, observed);
    }
  }

  for (const property of ["boldCs", "italicCs", "fontSizeCs"] as const) {
    const currentIsAbsent = current?.complexScriptPropertyAbsences?.includes(property) === true;
    const expectedIsAbsent = expected?.complexScriptPropertyAbsences?.includes(property) === true;
    if (currentIsAbsent !== expectedIsAbsent && currentIsAbsent) {
      Reflect.deleteProperty(formatting, property);
    }
  }
};

/**
 * Reconcile source-authored provenance with the current visual mark set at the
 * one save boundary shared by editor commands, raw transactions and compare.
 * Unchanged groups keep exact source state (including equal direct values and
 * otherwise invisible sentinels); changed groups derive the smallest direct
 * override needed relative to the current style cascade.
 */
const reconcileAuthoredFormatting = ({
  authoredFormatting,
  carrierlessContext,
  currentOverrideAttrs,
  inheritedFormatting,
  observedFormatting,
}: ReconcileAuthoredFormattingOptions): TextFormatting => {
  const hasAuthoredFormatting = Object.keys(authoredFormatting).length > 0;
  if (
    carrierlessContext === "ordinary" &&
    currentOverrideAttrs === undefined &&
    !hasAuthoredFormatting &&
    hasOnlyCarrierlessVisualFormatting(observedFormatting)
  ) {
    if (sameVisualFormatting(observedFormatting, inheritedFormatting)) {
      return authoredFormatting;
    }
    if (
      observedFormatting.fontFamily === undefined &&
      sameVisualFormatting(inheritedFormatting, undefined)
    ) {
      return observedFormatting;
    }
  }
  const expectedFormatting = hasAuthoredFormatting
    ? mergeTextFormatting(inheritedFormatting, authoredFormatting)
    : inheritedFormatting;
  const changedGroups = changedVisualFormattingGroups(observedFormatting, expectedFormatting);
  const reconciled: TextFormatting = { ...authoredFormatting };
  for (const property of Object.keys(RUN_FORMATTING_VISUAL_GROUPS) as Array<keyof TextFormatting>) {
    const group = RUN_FORMATTING_VISUAL_GROUPS[property];
    if (group === null || !changedGroups.has(group)) {
      continue;
    }
    Reflect.deleteProperty(reconciled, property);
    if (VISUAL_BOOLEAN_FORMATTING_PROPERTIES.has(property)) {
      const observed = observedFormatting[property] === true;
      const inherited = inheritedFormatting?.[property] === true;
      if (observed !== inherited) {
        Reflect.set(reconciled, property, observed);
      }
      continue;
    }

    const observed = observedFormatting[property];
    if (observed === undefined) {
      continue;
    }
    if (property === "fontFamily") {
      const difference = fontFamilyDifference(
        observed as NonNullable<TextFormatting["fontFamily"]>,
        inheritedFormatting?.fontFamily,
      );
      if (difference) {
        reconciled.fontFamily = difference;
      }
      continue;
    }
    if (!sameFormattingValue(observed, inheritedFormatting?.[property])) {
      Reflect.set(reconciled, property, observed);
    }
  }
  if (currentOverrideAttrs !== undefined) {
    reconcileOverrideSignals(
      reconciled,
      currentOverrideAttrs,
      expectedRunFormattingOverrideAttrs(authoredFormatting),
      inheritedFormatting,
      observedFormatting,
    );
  }
  return reconciled;
};

const formattingAuthorshipCost = (formatting: TextFormatting): number => {
  let cost = 0;
  const visit = (value: unknown): void => {
    if (value === undefined) {
      return;
    }
    if (value === null || typeof value !== "object") {
      cost++;
      return;
    }
    const entries = Object.entries(value);
    if (entries.length === 0) {
      cost++;
      return;
    }
    for (const [, nested] of entries) {
      visit(nested);
    }
  };
  for (const [property, value] of Object.entries(formatting)) {
    if (property !== "styleId") {
      visit(value);
    }
  }
  return cost;
};

export function marksToTextFormatting(
  marks: readonly Mark[],
  options?: MarksToTextFormattingOptions,
): TextFormatting {
  const formatting: TextFormatting = {};
  let directOverrideFormatting: TextFormatting | undefined;
  let directFontProperties: RunFormattingOverrideAttrs["directFontProperties"];
  let characterStyleId: string | undefined;
  let runFormattingOverrideMark: Mark | undefined;
  let currentOverrideAttrs: RunFormattingOverrideAttrs | undefined;

  for (const mark of marks) {
    switch (mark.type.name) {
      case "bold":
        formatting.bold = true;
        break;

      case "italic":
        formatting.italic = true;
        break;

      case "underline": {
        const attrs = expectUnderlineMarkAttrs(mark);
        const uline: NonNullable<TextFormatting["underline"]> = {
          style: attrs.style || "single",
        };
        if (attrs.color) {
          uline.color = attrs.color;
        }
        formatting.underline = uline;
        break;
      }

      case "strike":
        if (expectStrikeMarkAttrs(mark).double) {
          formatting.doubleStrike = true;
        } else {
          formatting.strike = true;
        }
        break;

      case "textColor": {
        const attrs = expectTextColorMarkAttrs(mark);
        const colorVal: ColorValue = {};
        if (attrs.rgb) {
          colorVal.rgb = attrs.rgb;
        }
        if (attrs.themeColor) {
          colorVal.themeColor = attrs.themeColor;
        }
        if (attrs.themeTint) {
          colorVal.themeTint = attrs.themeTint;
        }
        if (attrs.themeShade) {
          colorVal.themeShade = attrs.themeShade;
        }
        formatting.color = colorVal;
        break;
      }

      case "highlight":
        formatting.highlight = expectHighlightMarkAttrs(mark).color;
        break;

      case "runShading":
        // Rebuild the model `w:shd` fill so the run serializer re-emits it.
        formatting.shading = runShadingAttrsToShading(expectRunShadingMarkAttrs(mark));
        break;

      case "fontSize": {
        const attrs = expectFontSizeMarkAttrs(mark);
        formatting.fontSize = attrs.size;
        break;
      }

      case "fontFamily": {
        const attrs = expectFontFamilyMarkAttrs(mark);
        const ff: NonNullable<TextFormatting["fontFamily"]> = {};
        if (attrs.ascii) {
          ff.ascii = attrs.ascii;
        }
        if (attrs.hAnsi) {
          ff.hAnsi = attrs.hAnsi;
        }
        if (attrs.eastAsia) {
          ff.eastAsia = attrs.eastAsia;
        }
        if (attrs.hint) {
          ff.hint = attrs.hint;
        }
        if (attrs.cs) {
          ff.cs = attrs.cs;
        }
        // asciiTheme needs to be cast to the proper type
        if (attrs.asciiTheme) {
          ff.asciiTheme = attrs.asciiTheme as NonNullable<
            NonNullable<TextFormatting["fontFamily"]>["asciiTheme"]
          >;
        }
        if (attrs.hAnsiTheme) {
          ff.hAnsiTheme = attrs.hAnsiTheme;
        }
        if (attrs.eastAsiaTheme) {
          ff.eastAsiaTheme = attrs.eastAsiaTheme;
        }
        if (attrs.csTheme) {
          ff.csTheme = attrs.csTheme;
        }
        formatting.fontFamily = ff;
        break;
      }

      case "language": {
        const attrs = expectLanguageMarkAttrs(mark);
        formatting.language = {
          ...(attrs.val ? { val: attrs.val } : {}),
          ...(attrs.eastAsia ? { eastAsia: attrs.eastAsia } : {}),
          ...(attrs.bidi ? { bidi: attrs.bidi } : {}),
        };
        break;
      }

      case "superscript":
        formatting.vertAlign = "superscript";
        break;

      case "subscript":
        formatting.vertAlign = "subscript";
        break;

      case "allCaps":
        formatting.allCaps = true;
        break;

      case "smallCaps":
        formatting.smallCaps = true;
        break;

      case "characterSpacing": {
        const attrs = expectCharacterSpacingMarkAttrs(mark);
        if (attrs.spacing !== undefined) {
          formatting.spacing = attrs.spacing;
        }
        if (attrs.position !== undefined) {
          formatting.position = attrs.position;
        }
        const horizontalScale = normalizeHorizontalScalePercent(attrs.scale);
        if (horizontalScale !== undefined) {
          formatting.scale = horizontalScale;
        }
        if (attrs.kerning !== undefined) {
          formatting.kerning = attrs.kerning;
        }
        break;
      }

      case "emboss":
        formatting.emboss = true;
        break;

      case "imprint":
        formatting.imprint = true;
        break;

      case "hidden":
        // eigenpal #424 (w:vanish gap 9): mark closes the round-trip so
        // `<w:vanish/>` survives parse → PM → serialize.
        formatting.hidden = true;
        break;

      case "textShadow":
        formatting.shadow = true;
        break;

      case "emphasisMark":
        formatting.emphasisMark = expectEmphasisMarkAttrs(mark).type || "dot";
        break;

      case "textOutline":
        formatting.outline = true;
        break;

      case "rtl":
        formatting.rtl = true;
        break;

      case "textEffect":
        formatting.effect = expectTextEffectMarkAttrs(mark).effect;
        break;

      case "runFormattingOverride":
        runFormattingOverrideMark = mark;
        break;

      case "characterStyle": {
        const attrs = expectCharacterStyleMarkAttrs(mark);
        formatting.styleId = attrs.styleId;
        characterStyleId = attrs.styleId;
        break;
      }

      // hyperlink is handled separately
      default:
        break;
    }
  }

  let authoredFormatting: TextFormatting = {};
  if (runFormattingOverrideMark) {
    const overrideAttrs = expectRunFormattingOverrideMarkAttrs(runFormattingOverrideMark);
    currentOverrideAttrs = overrideAttrs;
    directFontProperties = overrideAttrs.directFontProperties;
    directOverrideFormatting = {};
    applyRunFormattingOverrideAttrs(directOverrideFormatting, overrideAttrs);
    for (const property of directFontProperties ?? []) {
      const value = formatting[property];
      if (value !== undefined) {
        Reflect.set(directOverrideFormatting, property, value);
      }
    }
    authoredFormatting = authoredRunFormattingFromAttrs(overrideAttrs) ?? directOverrideFormatting;
  }

  if (characterStyleId !== undefined && !options?.styleResolver) {
    const conservativeFormatting = mergeTextFormatting(formatting, directOverrideFormatting) ?? {};
    for (const property of currentOverrideAttrs?.complexScriptPropertyAbsences ?? []) {
      Reflect.deleteProperty(conservativeFormatting, property);
    }
    return { ...conservativeFormatting, styleId: characterStyleId };
  }

  const runContext = {
    baseParagraphFormatting: options?.baseParagraphFormatting,
    paragraphFormatting: options?.inheritedFormatting,
    paragraphMarkFormatting: options?.paragraphMarkFormatting,
    paragraphMarkPrecedesStyle: options?.paragraphMarkPrecedesStyle ?? false,
  };
  const paragraphFormatting = paragraphFormattingForRun({
    context: runContext,
    directFormatting: authoredFormatting,
    marks,
  });
  const inheritedFormatting = resolveEffectiveRunStyleFormatting({
    marks,
    paragraphFormatting,
    ...(options?.styleResolver !== undefined ? { styleResolver: options.styleResolver } : {}),
  });
  let reconciled = reconcileAuthoredFormatting({
    authoredFormatting,
    carrierlessContext:
      options?.paragraphMarkFormatting === undefined ? "ordinary" : "paragraph-mark",
    currentOverrideAttrs,
    inheritedFormatting,
    observedFormatting: formatting,
  });
  if (
    runFormattingOverrideMark === undefined &&
    characterStyleId === undefined &&
    options?.paragraphMarkFormatting !== undefined
  ) {
    const suppressedParagraphFormatting = suppressParagraphMarkFormatting({
      baseFormatting: options.baseParagraphFormatting,
      directFormatting: authoredFormatting,
      paragraphMarkFormatting: options.paragraphMarkFormatting,
      paragraphMarkPrecedesStyle: options.paragraphMarkPrecedesStyle ?? false,
    });
    const suppressedInheritedFormatting = resolveEffectiveRunStyleFormatting({
      marks,
      paragraphFormatting: suppressedParagraphFormatting,
      ...(options.styleResolver !== undefined ? { styleResolver: options.styleResolver } : {}),
    });
    const suppressed = reconcileAuthoredFormatting({
      authoredFormatting,
      carrierlessContext: "paragraph-mark",
      currentOverrideAttrs,
      inheritedFormatting: suppressedInheritedFormatting,
      observedFormatting: formatting,
    });
    if (formattingAuthorshipCost(suppressed) < formattingAuthorshipCost(reconciled)) {
      reconciled = suppressed;
    }
  }
  return {
    ...reconciled,
    ...(characterStyleId !== undefined ? { styleId: characterStyleId } : {}),
  };
}
