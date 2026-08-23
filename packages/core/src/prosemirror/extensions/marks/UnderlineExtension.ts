/**
 * Underline Mark Extension
 */

import { panic } from "better-result";

import { expectUnderlineMarkAttrs } from "../../attrs";
import type { TextColorAttrs } from "../../schema/marks";
import { createMarkExtension } from "../create";
import type { ExtensionContext, ExtensionRuntime } from "../types";
import { setMark, toggleUnderlineMark } from "./markUtils";

export const UnderlineExtension = createMarkExtension({
  name: "underline",
  schemaMarkName: "underline",
  markSpec: {
    attrs: {
      style: { default: "single" },
      color: { default: null },
    },
    parseDOM: [
      { tag: "u" },
      {
        style: "text-decoration",
        getAttrs: (value) => {
          if (value.includes("underline")) {
            return {};
          }
          return value.includes("none") ? { style: "none" } : false;
        },
      },
    ],
    toDOM(mark) {
      const attrs = expectUnderlineMarkAttrs(mark);
      const style = attrs.style;
      const colorRgb = attrs.color?.rgb;
      const cssStyle: string[] = [`text-decoration: ${style === "none" ? "none" : "underline"}`];

      if (style && style !== "single" && style !== "none") {
        const styleMap: Record<string, string> = {
          double: "double",
          dotted: "dotted",
          dash: "dashed",
          wave: "wavy",
        };
        const cssDecorationStyle = styleMap[style];
        if (cssDecorationStyle) {
          cssStyle.push(`text-decoration-style: ${cssDecorationStyle}`);
        }
      }

      if (colorRgb) {
        cssStyle.push(`text-decoration-color: #${colorRgb}`);
      }

      return ["span", { style: cssStyle.join("; ") }, 0];
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
