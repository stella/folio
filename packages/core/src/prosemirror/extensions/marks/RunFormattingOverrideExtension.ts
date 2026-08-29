/**
 * Run formatting override mark.
 *
 * OOXML can explicitly set inherited boolean/defaultable run properties. Plain
 * PM marks cannot distinguish inherited visuals from authored direct values, so
 * this mark carries those direct toggle overrides.
 */

import type { Mark } from "prosemirror-model";

import type { TextFormatting } from "../../../types/document";
import { expectRunFormattingOverrideMarkAttrs } from "../../attrs";
import type { RunFormattingOverrideAttrs } from "../../schema/marks";
import { createMarkExtension } from "../create";

export function buildRunFormattingOverrideAttrs(
  formatting: TextFormatting | undefined,
): RunFormattingOverrideAttrs | undefined {
  if (!formatting) {
    return undefined;
  }

  const attrs: RunFormattingOverrideAttrs = {};

  if (formatting.bold !== undefined) {
    attrs.bold = formatting.bold;
  }
  if (formatting.italic !== undefined) {
    attrs.italic = formatting.italic;
  }
  if (formatting.underline?.style === "none") {
    attrs.underline = "none";
  }
  if (formatting.strike !== undefined) {
    attrs.strike = formatting.strike;
  }
  if (formatting.doubleStrike === false) {
    attrs.doubleStrike = false;
  }
  if (formatting.allCaps !== undefined) {
    attrs.allCaps = formatting.allCaps;
  }
  if (formatting.smallCaps !== undefined) {
    attrs.smallCaps = formatting.smallCaps;
  }
  if (formatting.hidden !== undefined) {
    attrs.hidden = formatting.hidden;
  }
  if (formatting.emboss !== undefined) {
    attrs.emboss = formatting.emboss;
  }
  if (formatting.imprint !== undefined) {
    attrs.imprint = formatting.imprint;
  }
  if (formatting.shadow !== undefined) {
    attrs.shadow = formatting.shadow;
  }
  if (formatting.outline !== undefined) {
    attrs.outline = formatting.outline;
  }
  if (formatting.rtl === false) {
    attrs.rtl = false;
  }
  if (formatting.boldCs !== undefined) {
    attrs.boldCs = formatting.boldCs;
  }
  if (formatting.italicCs !== undefined) {
    attrs.italicCs = formatting.italicCs;
  }
  if (formatting.fontSizeCs !== undefined && formatting.fontSizeCs !== formatting.fontSize) {
    attrs.fontSizeCs = formatting.fontSizeCs;
  }
  if (formatting.cs !== undefined) {
    attrs.cs = formatting.cs;
  }

  return Object.keys(attrs).length > 0 ? attrs : undefined;
}

export function applyRunFormattingOverrideAttrs(
  formatting: TextFormatting,
  attrs: RunFormattingOverrideAttrs,
): void {
  if (attrs.bold !== undefined) {
    formatting.bold = attrs.bold;
  }
  if (attrs.italic !== undefined) {
    formatting.italic = attrs.italic;
  }
  if (attrs.underline === "none") {
    formatting.underline = { style: "none" };
  }
  if (attrs.strike !== undefined) {
    formatting.strike = attrs.strike;
  }
  if (attrs.doubleStrike === false) {
    formatting.doubleStrike = false;
  }
  if (attrs.allCaps !== undefined) {
    formatting.allCaps = attrs.allCaps;
  }
  if (attrs.smallCaps !== undefined) {
    formatting.smallCaps = attrs.smallCaps;
  }
  if (attrs.hidden !== undefined) {
    formatting.hidden = attrs.hidden;
  }
  if (attrs.emboss !== undefined) {
    formatting.emboss = attrs.emboss;
  }
  if (attrs.imprint !== undefined) {
    formatting.imprint = attrs.imprint;
  }
  if (attrs.shadow !== undefined) {
    formatting.shadow = attrs.shadow;
  }
  if (attrs.outline !== undefined) {
    formatting.outline = attrs.outline;
  }
  if (attrs.rtl === false) {
    formatting.rtl = false;
  }
  if (attrs.boldCs !== undefined) {
    formatting.boldCs = attrs.boldCs;
  }
  if (attrs.italicCs !== undefined) {
    formatting.italicCs = attrs.italicCs;
  }
  if (attrs.fontSizeCs !== undefined) {
    formatting.fontSizeCs = attrs.fontSizeCs;
  }
  if (attrs.cs !== undefined) {
    formatting.cs = attrs.cs;
  }
}

export function applyRunFormattingOverrideMark(formatting: TextFormatting, mark: Mark): void {
  applyRunFormattingOverrideAttrs(formatting, expectRunFormattingOverrideMarkAttrs(mark));
}

export const RunFormattingOverrideExtension = createMarkExtension({
  name: "runFormattingOverride",
  schemaMarkName: "runFormattingOverride",
  markSpec: {
    attrs: {
      bold: { default: null },
      italic: { default: null },
      underline: { default: null },
      strike: { default: null },
      doubleStrike: { default: null },
      allCaps: { default: null },
      smallCaps: { default: null },
      hidden: { default: null },
      emboss: { default: null },
      imprint: { default: null },
      shadow: { default: null },
      outline: { default: null },
      rtl: { default: null },
      boldCs: { default: null },
      italicCs: { default: null },
      fontSizeCs: { default: null },
      cs: { default: null },
    },
    toDOM(mark) {
      const attrs = expectRunFormattingOverrideMarkAttrs(mark);
      const styles: string[] = [];

      if (attrs.bold === false) {
        styles.push("font-weight: normal");
      }
      if (attrs.italic === false) {
        styles.push("font-style: normal");
      }
      if (attrs.underline === "none" || attrs.strike === false) {
        styles.push("text-decoration: none");
      }
      if (attrs.allCaps === false) {
        styles.push("text-transform: none");
      }
      if (attrs.smallCaps === false) {
        styles.push("font-variant-caps: normal");
      }
      if (attrs.hidden === false) {
        styles.push("visibility: visible");
      }

      return ["span", styles.length > 0 ? { style: styles.join("; ") } : {}, 0];
    },
  },
});
