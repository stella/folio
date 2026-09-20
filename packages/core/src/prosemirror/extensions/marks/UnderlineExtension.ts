/**
 * Underline Mark Extension
 */

import { panic } from "better-result";

import {
  PLAIN_UNDERLINE,
  underlineDecorationCss,
  underlineStyleFromCssDecoration,
} from "../../../utils/formatToStyle";
import { expectUnderlineMarkAttrs } from "../../attrs";
import type { TextColorAttrs } from "../../schema/marks";
import { createMarkExtension } from "../create";
import type { ExtensionContext, ExtensionRuntime } from "../types";
import { setMark, toggleUnderlineMark } from "./markUtils";

/** `w:u w:val="none"`: the token that cancels an inherited underline. */
const NO_UNDERLINE = "none";

export const UnderlineExtension = createMarkExtension({
  name: "underline",
  schemaMarkName: "underline",
  markSpec: {
    attrs: {
      style: { default: PLAIN_UNDERLINE },
      color: { default: null },
    },
    parseDOM: [
      { tag: "u" },
      {
        // The shorthand carries the line and, where an author wrote one, the
        // style: `text-decoration: underline dotted`. Both are read here
        // rather than from a `text-decoration-style` rule, which sees only its
        // own property's value and would parse a dotted strikethrough as an
        // underline.
        style: "text-decoration",
        getAttrs: (value) => {
          if (value.includes("underline")) {
            return { style: underlineStyleFromCssDecoration(value) };
          }
          return value.includes(NO_UNDERLINE) ? { style: NO_UNDERLINE } : false;
        },
      },
    ],
    toDOM(mark) {
      const { style = PLAIN_UNDERLINE, color } = expectUnderlineMarkAttrs(mark);
      const { decorationStyle, decorationThickness } = underlineDecorationCss(style);
      // Line and style go in the shorthand so a copy out of the editor parses
      // back through the one rule above; the longhands after it survive the
      // reset the shorthand performs.
      const line = decorationStyle === undefined ? NO_UNDERLINE : `underline ${decorationStyle}`;
      const declarations = [`text-decoration: ${line}`];

      if (decorationThickness !== undefined) {
        declarations.push(`text-decoration-thickness: ${decorationThickness}`);
      }
      if (color?.rgb) {
        declarations.push(`text-decoration-color: #${color.rgb}`);
      }

      return ["span", { style: declarations.join("; ") }, 0];
    },
  },
  onSchemaReady(ctx: ExtensionContext): ExtensionRuntime {
    const underlineType = ctx.schema.marks["underline"];
    if (!underlineType) {
      panic("Missing mark type: underline");
    }
    return {
      commands: {
        toggleUnderline: () => toggleUnderlineMark(underlineType),
        setUnderlineStyle: (style: string, color?: TextColorAttrs) =>
          setMark(underlineType, { style, color }),
      },
      keyboardShortcuts: {
        "Mod-u": toggleUnderlineMark(underlineType),
      },
    };
  },
});
