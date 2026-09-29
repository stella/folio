/**
 * Text Color Mark Extension
 */

import { panic } from "better-result";
import type { MarkType } from "prosemirror-model";
import type { Command } from "prosemirror-state";

import { textToStyle } from "../../../utils/formatToStyle";
import { expectTextColorMarkAttrs } from "../../attrs";
import type { TextColorAttrs } from "../../schema/marks";
import { createMarkExtension } from "../create";
import type { ExtensionContext, ExtensionRuntime } from "../types";
import { getDocumentStyleResolver } from "../../plugins/documentStyles";
// oxlint-disable-next-line import/no-cycle -- runtime-only: run formatting is re-resolved inside command handlers, not at module load
import { rebaseRunFormattingInRange } from "../../rebaseParagraphRunFormatting";
import { setMark, removeMark } from "./markUtils";

/**
 * Remove the direct text color from the selection. Only the direct color
 * goes: a color the run's styles give it (a hyperlink's character style, a
 * heading's paragraph style) paints again, as it does once the document is
 * reopened.
 */
const clearDirectTextColor =
  (textColorType: MarkType): Command =>
  (state, dispatch) => {
    const { from, to, empty } = state.selection;
    const styleResolver = getDocumentStyleResolver(state);
    return removeMark(textColorType)(
      state,
      dispatch &&
        ((tr) => {
          dispatch(
            empty || !styleResolver ? tr : rebaseRunFormattingInRange(tr, from, to, styleResolver),
          );
        }),
    );
  };

export const TextColorExtension = createMarkExtension({
  name: "textColor",
  schemaMarkName: "textColor",
  markSpec: {
    attrs: {
      rgb: { default: null },
      themeColor: { default: null },
      themeTint: { default: null },
      themeShade: { default: null },
    },
    parseDOM: [
      {
        style: "color",
        getAttrs: (value) => {
          const hexMatch = /#(?<hex>[0-9a-fA-F]{6}|[0-9a-fA-F]{3})/u.exec(value);
          if (hexMatch) {
            // SAFETY: capture group always present when regex matches
            return { rgb: (hexMatch.groups?.["hex"] ?? "").toUpperCase() };
          }
          return false;
        },
      },
    ],
    toDOM(mark) {
      const colorAttrs = expectTextColorMarkAttrs(mark);
      const style = textToStyle({ color: colorAttrs });
      const cssColor: unknown = style.color;
      const cssString = typeof cssColor === "string" && cssColor ? `color: ${cssColor}` : "";
      return ["span", { style: cssString }, 0];
    },
  },
  onSchemaReady(ctx: ExtensionContext): ExtensionRuntime {
    const textColorType = ctx.schema.marks["textColor"];
    if (!textColorType) {
      panic("Missing mark type: textColor");
    }
    return {
      commands: {
        setTextColor: (attrs: TextColorAttrs) => {
          if (!attrs.rgb && !attrs.themeColor) {
            return clearDirectTextColor(textColorType);
          }
          return setMark(textColorType, attrs as Record<string, unknown>);
        },
        clearTextColor: () => clearDirectTextColor(textColorType),
      },
    };
  },
});
