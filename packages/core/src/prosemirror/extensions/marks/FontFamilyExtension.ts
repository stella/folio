/**
 * Font Family Mark Extension
 */

import { panic } from "better-result";

import { editingFontFamily } from "../../../layout-bridge/convert/editingFont";
import { canonicalRunFormatting } from "../../canonicalCommands";
import { createMarkExtension } from "../create";
import type {
  ExtensionContext,
  ExtensionRuntime,
  MarkExtension,
} from "../types";
import {
  canonicalSetMarkFormatting,
  removeMark,
  setMark,
  withRunFormattingLock,
} from "./markUtils";

export const FontFamilyExtension: (
  options?: Partial<Record<string, unknown>> | undefined,
) => MarkExtension = createMarkExtension({
  name: "fontFamily",
  schemaMarkName: "fontFamily",
  markSpec: {
    attrs: {
      ascii: { default: null },
      hAnsi: { default: null },
      eastAsia: { default: null },
      cs: { default: null },
      hint: { default: null },
      asciiTheme: { default: null },
      hAnsiTheme: { default: null },
      eastAsiaTheme: { default: null },
      csTheme: { default: null },
    },
    parseDOM: [
      {
        style: "font-family",
        getAttrs: (value) => {
          // SAFETY: split always returns at least one element
          const firstFont = (value.split(",")[0] ?? "")
            .trim()
            .replace(/['"]/gu, "");
          if (firstFont) {
            return { ascii: firstFont };
          }
          return false;
        },
      },
    ],
    toDOM(mark) {
      // Without an editor's font context; the hidden editor's mark view
      // resolves the same mark with its document's context.
      const fontFamily = editingFontFamily(mark, {});
      return fontFamily === null
        ? ["span", 0]
        : ["span", { style: `font-family: ${fontFamily}` }, 0];
    },
  },
  onSchemaReady(ctx: ExtensionContext): ExtensionRuntime {
    const fontFamilyType = ctx.schema.marks["fontFamily"];
    if (!fontFamilyType) {
      panic("Missing mark type: fontFamily");
    }
    return {
      commands: {
        setFontFamily: (fontName: string) =>
          withRunFormattingLock(
            setMark(fontFamilyType, {
              ascii: fontName,
              hAnsi: fontName,
            }),
            (state) =>
              canonicalSetMarkFormatting({
                state,
                type: fontFamilyType,
                attrs: { ascii: fontName, hAnsi: fontName },
              }),
          ),
        clearFontFamily: () =>
          withRunFormattingLock(removeMark(fontFamilyType), (state) =>
            canonicalRunFormatting(state, { fontFamily: null }),
          ),
      },
    };
  },
});
