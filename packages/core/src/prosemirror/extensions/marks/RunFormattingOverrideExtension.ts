/**
 * Run formatting override mark.
 *
 * OOXML can explicitly turn inherited boolean/defaultable run properties off
 * (for example <w:b w:val="0"/>). Plain PM marks cannot distinguish "inherit"
 * from "explicitly off", so this mark carries those negative overrides.
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

  if (formatting.bold === false) {
    attrs.bold = false;
  }
  if (formatting.italic === false) {
    attrs.italic = false;
  }
  if (formatting.underline?.style === "none") {
    attrs.underline = "none";
  }
  if (formatting.strike === false) {
    attrs.strike = false;
  }
  if (formatting.doubleStrike === false) {
    attrs.doubleStrike = false;
  }
  if (formatting.allCaps === false) {
    attrs.allCaps = false;
  }
  if (formatting.smallCaps === false) {
    attrs.smallCaps = false;
  }
  if (formatting.hidden === false) {
    attrs.hidden = false;
  }
  if (formatting.emboss === false) {
    attrs.emboss = false;
  }
  if (formatting.imprint === false) {
    attrs.imprint = false;
  }
  if (formatting.shadow === false) {
    attrs.shadow = false;
  }
  if (formatting.outline === false) {
    attrs.outline = false;
  }
  if (formatting.rtl === false) {
    attrs.rtl = false;
  }
  if (formatting.boldCs !== undefined && formatting.boldCs !== formatting.bold) {
    attrs.boldCs = formatting.boldCs;
  }
  if (formatting.italicCs !== undefined && formatting.italicCs !== formatting.italic) {
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
  if (attrs.bold === false) {
    formatting.bold = false;
  }
  if (attrs.italic === false) {
    formatting.italic = false;
  }
  if (attrs.underline === "none") {
    formatting.underline = { style: "none" };
  }
  if (attrs.strike === false) {
    formatting.strike = false;
  }
  if (attrs.doubleStrike === false) {
    formatting.doubleStrike = false;
  }
  if (attrs.allCaps === false) {
    formatting.allCaps = false;
  }
  if (attrs.smallCaps === false) {
    formatting.smallCaps = false;
  }
  if (attrs.hidden === false) {
    formatting.hidden = false;
  }
  if (attrs.emboss === false) {
    formatting.emboss = false;
  }
  if (attrs.imprint === false) {
    formatting.imprint = false;
  }
  if (attrs.shadow === false) {
    formatting.shadow = false;
  }
  if (attrs.outline === false) {
    formatting.outline = false;
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
