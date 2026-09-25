/**
 * Run-formatting merge rules: layering direct formatting over character
 * styles, character-style toggle properties, explicit run-formatting override
 * marks, and default-black text colour provenance.
 */

import type { Mark } from "prosemirror-model";
import type { RunFormatting } from "../../layout-engine/types";
import { STYLE_TOGGLE_KEYS } from "../../utils/textFormattingMerge";
import { expectCharacterStyleMarkAttrs } from "../../prosemirror/attrs";
import {
  resolveEffectiveRunStyleFormatting,
  type RunStyleResolver,
} from "../../prosemirror/runStyleFormatting";
import type { RunFormattingOverrideAttrs } from "../../prosemirror/schema/marks";
import type { TextFormatting } from "../../types/document";

const DEFAULT_BLACK_TEXT_COLOR_VALUES = new Set(["000000", "000"]);

function normalizeResolvedTextColor(color: string): string {
  return color.trim().toLowerCase().replace(/^#/u, "");
}

function isDefaultBlackResolvedTextColor(color: string): boolean {
  return DEFAULT_BLACK_TEXT_COLOR_VALUES.has(normalizeResolvedTextColor(color));
}

function areResolvedTextColorsEqual(left: string, right: string): boolean {
  return normalizeResolvedTextColor(left) === normalizeResolvedTextColor(right);
}

export function markDefaultBlackTextColorSource(
  formatting: RunFormatting,
  paraDefaults: RunFormatting,
): RunFormatting {
  if (
    formatting.textColorSource === "direct" ||
    formatting.color === undefined ||
    paraDefaults.color === undefined ||
    !isDefaultBlackResolvedTextColor(formatting.color) ||
    !areResolvedTextColorsEqual(formatting.color, paraDefaults.color)
  ) {
    return formatting;
  }

  return {
    ...formatting,
    textColorSource: "paragraphDefault",
  };
}

export function mergeRunFormatting(
  paraDefaults: RunFormatting,
  formatting: RunFormatting,
): RunFormatting {
  // A deletion represents the pre-change run. Word does not apply the current
  // paragraph mark's run properties to that historical content; the projected
  // deletion marks already carry its resolved original formatting. Let only
  // the document terminal fallback fill properties that were absent at every
  // original style level.
  if (formatting.isDeletion && formatting.usesHistoricalFormatting) {
    return formatting;
  }
  const merged = {
    ...paraDefaults,
    ...markDefaultBlackTextColorSource(formatting, paraDefaults),
  };
  if (formatting.fontFamily !== undefined && formatting.alternateFontFamily === undefined) {
    delete merged.alternateFontFamily;
  }
  if (
    formatting.eastAsiaFontFamily !== undefined &&
    formatting.eastAsiaAlternateFontFamily === undefined
  ) {
    delete merged.eastAsiaAlternateFontFamily;
  }
  if (
    formatting.complexScriptFontFamily !== undefined &&
    formatting.complexScriptAlternateFontFamily === undefined
  ) {
    delete merged.complexScriptAlternateFontFamily;
  }
  if (merged.letterSpacing === 0) {
    delete merged.letterSpacing;
  }
  if (formatting.horizontalScale === 100 && paraDefaults.horizontalScale === undefined) {
    delete merged.horizontalScale;
  }
  return merged;
}

type ApplyCharacterStyleToggleFormattingOptions = {
  formatting: RunFormatting;
  marks: readonly Mark[];
  paragraphFormatting: TextFormatting | undefined;
  styleResolver: RunStyleResolver;
};

const CHARACTER_STYLE_TOGGLE_LAYOUT_PROPERTIES = {
  allCaps: "allCaps",
  bold: "bold",
  boldCs: "complexScriptBold",
  emboss: "emboss",
  hidden: "hidden",
  imprint: "imprint",
  italic: "italic",
  italicCs: "complexScriptItalic",
  outline: "textOutline",
  shadow: "textShadow",
  smallCaps: "smallCaps",
  strike: "strike",
} as const satisfies Record<(typeof STYLE_TOGGLE_KEYS)[number], keyof RunFormatting>;

/** Restore character-style toggle values that plain visual marks cannot represent. */
export function applyCharacterStyleToggleFormatting({
  formatting,
  marks,
  paragraphFormatting,
  styleResolver,
}: ApplyCharacterStyleToggleFormattingOptions): void {
  const characterStyleMark = marks.find((mark) => mark.type.name === "characterStyle");
  const styleRPr = characterStyleMark
    ? styleResolver.getRunStyleOwnProperties(
        expectCharacterStyleMarkAttrs(characterStyleMark).styleId,
      )
    : styleResolver.getDefaultCharacterStyle()?.rPr;
  if (!styleRPr) {
    return;
  }

  const effectiveStyleFormatting = resolveEffectiveRunStyleFormatting({
    marks,
    paragraphFormatting,
    styleResolver,
  });
  for (const property of STYLE_TOGGLE_KEYS) {
    if (styleRPr[property] === undefined) {
      continue;
    }
    const layoutProperty = CHARACTER_STYLE_TOGGLE_LAYOUT_PROPERTIES[property];
    if (formatting[layoutProperty] === undefined) {
      Reflect.set(formatting, layoutProperty, effectiveStyleFormatting?.[property] ?? false);
    }
  }
}

type RunFormattingOverrideHandler = (formatting: RunFormatting, value: unknown) => void;

const suppressInheritedRunFormatting = (
  formatting: RunFormatting,
  property: keyof RunFormatting,
): void => {
  // The enumerable own property must survive the later object spread and
  // replace the paragraph default with the property's neutral state.
  Reflect.set(formatting, property, undefined);
};

/**
 * Project every structural run-formatting override into layout. This total map
 * keeps save/reopen semantics and the live layout in lockstep: adding a new
 * override attr requires an explicit rendering decision here.
 */
const RUN_FORMATTING_OVERRIDE_HANDLERS = {
  _authoredOff() {
    // Serialization provenance; current rendering signals own layout.
  },
  _authoredOn() {
    // Serialization provenance; current rendering signals own layout.
  },
  _authoredValues() {
    // Serialization provenance; current rendering signals own layout.
  },
  allCaps(formatting, value) {
    if (typeof value === "boolean") {
      formatting.allCaps = value;
    }
  },
  bold(formatting, value) {
    if (typeof value === "boolean") {
      formatting.bold = value;
    }
  },
  boldCs(formatting, value) {
    if (typeof value === "boolean") {
      formatting.complexScriptBold = value;
    }
  },
  color(formatting, value) {
    if (value === "auto") {
      suppressInheritedRunFormatting(formatting, "color");
      suppressInheritedRunFormatting(formatting, "textColorSource");
    }
  },
  complexScriptPropertyAbsences() {
    // Serialization provenance; current rendering signals own layout.
  },
  cs(formatting, value) {
    if (typeof value === "boolean") {
      formatting.forceComplexScript = value;
    }
  },
  directFontProperties() {
    // Serialization provenance only; concrete font and color marks own layout.
  },
  doubleStrike(formatting, value) {
    if (value === false && formatting.strike === undefined) {
      // Layout currently paints single and double strike through one field. A
      // copied single-strike mark wins; otherwise this cancels inherited strike.
      formatting.strike = false;
    }
  },
  effect(formatting, value) {
    if (value === "none") {
      suppressInheritedRunFormatting(formatting, "textEffect");
    }
  },
  emboss(formatting, value) {
    if (typeof value === "boolean") {
      formatting.emboss = value;
    }
  },
  emphasisMark(formatting, value) {
    if (value === "none") {
      suppressInheritedRunFormatting(formatting, "emphasisMark");
    }
  },
  fontSizeCs(formatting, value) {
    if (typeof value === "number") {
      formatting.complexScriptFontSize = value / 2;
    }
  },
  noProof() {
    // Proofing metadata does not affect layout.
  },
  hidden(formatting, value) {
    if (typeof value === "boolean") {
      formatting.hidden = value;
    }
  },
  highlight(formatting, value) {
    if (value === "none") {
      suppressInheritedRunFormatting(formatting, "highlight");
    }
  },
  imprint(formatting, value) {
    if (typeof value === "boolean") {
      formatting.imprint = value;
    }
  },
  italic(formatting, value) {
    if (typeof value === "boolean") {
      formatting.italic = value;
    }
  },
  italicCs(formatting, value) {
    if (typeof value === "boolean") {
      formatting.complexScriptItalic = value;
    }
  },
  kerning(formatting, value) {
    if (value === 0) {
      formatting.kerningMinPt = 0;
    }
  },
  outline(formatting, value) {
    if (typeof value === "boolean") {
      formatting.textOutline = value;
    }
  },
  position(formatting, value) {
    if (value === 0) {
      formatting.positionPx = 0;
    }
  },
  rtl(formatting, value) {
    if (value === false) {
      formatting.rtl = false;
    }
  },
  scale(formatting, value) {
    if (value === 100) {
      formatting.horizontalScale = 100;
    }
  },
  shading(formatting, value) {
    if (typeof value === "object" && value !== null && Reflect.get(value, "pattern") === "nil") {
      suppressInheritedRunFormatting(formatting, "shading");
    }
  },
  shadow(formatting, value) {
    if (typeof value === "boolean") {
      formatting.textShadow = value;
    }
  },
  smallCaps(formatting, value) {
    if (typeof value === "boolean") {
      formatting.smallCaps = value;
    }
  },
  spacing(formatting, value) {
    if (value === 0) {
      formatting.letterSpacing = 0;
    }
  },
  strike(formatting, value) {
    if (typeof value === "boolean") {
      formatting.strike = value;
    }
  },
  underline(formatting, value) {
    if (value === "none") {
      formatting.underline = false;
    }
  },
  vertAlign(formatting, value) {
    if (value === "baseline") {
      formatting.superscript = false;
      formatting.subscript = false;
    }
  },
} satisfies Record<keyof RunFormattingOverrideAttrs, RunFormattingOverrideHandler>;

const RUN_FORMATTING_OVERRIDE_HANDLER_ENTRIES = Object.entries(RUN_FORMATTING_OVERRIDE_HANDLERS);

export function applyRunFormattingOverrides(
  formatting: RunFormatting,
  attrs: RunFormattingOverrideAttrs,
): void {
  for (const [key, handler] of RUN_FORMATTING_OVERRIDE_HANDLER_ENTRIES) {
    handler(formatting, Reflect.get(attrs, key));
  }
}
