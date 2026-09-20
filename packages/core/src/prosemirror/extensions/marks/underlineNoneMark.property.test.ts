/**
 * `w:u w:val="none"` is ECMA-376's explicit "no underline" token (ST_Underline),
 * not a decoration: on a run or a style it cancels an underline inherited from
 * the style chain. `textFormattingToMarks` is where that is decided; the
 * painters agree (`UNDERLINE_STROKES` gives its row no line,
 * `UNDERLINE_DECORATION_STYLES` no keyword), so an `underline` mark
 * carrying `style: "none"` cannot reach either backend as a line.
 *
 * `textFormattingToMarks` is the only reader that decides whether an underline
 * mark exists. These properties hold over the whole `ST_Underline` enum and over
 * every entry point that restores stored marks from a paragraph's run defaults.
 */

import { describe, expect, test } from "bun:test";
import fc from "fast-check";
import { EditorState, TextSelection } from "prosemirror-state";

import { propertyConfig, propertyTestTimeout } from "../../../../../../test/property-testing";

import { UNDERLINE_STYLE_VALUES } from "../../../types/documentEnumValues";
import { singletonManager, schema } from "../../schema";
import { textFormattingToMarks } from "./markUtils";

const underlineStyleArbitrary = fc.constantFrom(...UNDERLINE_STYLE_VALUES);

const underlineMarkOf = (marks: readonly { type: { name: string } }[] | null) =>
  (marks ?? []).find(({ type }) => type.name === "underline");

/**
 * The stored-mark path: the caret entering an empty paragraph whose run
 * defaults carry the underline. This is the entry point that used to read the
 * token through a second, unguarded copy of `textFormattingToMarks`.
 */
const storedMarksForEmptyParagraph = (defaultTextFormatting: object) => {
  const doc = schema.node("doc", null, [schema.node("paragraph", { defaultTextFormatting })]);
  const state = EditorState.create({ doc, schema, plugins: singletonManager.getPlugins() });

  return state.apply(state.tr.setSelection(TextSelection.create(state.doc, 1))).storedMarks;
};

describe("the underline mark reader honours the `none` token", () => {
  test("an explicit `none` produces no underline mark", () => {
    const marks = textFormattingToMarks({ underline: { style: "none" } }, schema);

    expect(underlineMarkOf(marks)).toBeUndefined();
  });

  test(
    "a mark exists for exactly the styles that paint something",
    () => {
      fc.assert(
        fc.property(underlineStyleArbitrary, (style) => {
          const mark = underlineMarkOf(textFormattingToMarks({ underline: { style } }, schema));

          if (style === "none") {
            expect(mark).toBeUndefined();
            return;
          }
          expect(mark?.attrs["style"]).toBe(style);
        }),
        propertyConfig(),
      );
    },
    propertyTestTimeout(10_000),
  );

  test("typing in a paragraph whose defaults cancel the underline stays un-underlined", () => {
    const stored = storedMarksForEmptyParagraph({ bold: true, underline: { style: "none" } });

    expect(underlineMarkOf(stored)).toBeUndefined();
  });

  test(
    "the stored-mark path agrees with document conversion over the whole enum",
    () => {
      fc.assert(
        fc.property(underlineStyleArbitrary, fc.boolean(), (style, withColor) => {
          const formatting = {
            underline: { style, ...(withColor ? { color: { rgb: "FF0000" } } : {}) },
          };

          const viaConversion = underlineMarkOf(textFormattingToMarks(formatting, schema));
          const viaStoredMarks = underlineMarkOf(storedMarksForEmptyParagraph(formatting));

          expect(viaStoredMarks?.attrs).toEqual(viaConversion?.attrs);
        }),
        propertyConfig(),
      );
    },
    propertyTestTimeout(10_000),
  );
});
