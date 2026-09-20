/**
 * One `ST_Border` member renders one way, whichever path reaches it.
 *
 * Three copies of the OOXML → CSS table had accumulated: the layout bridge's
 * `OOXML_TO_CSS_BORDER`, `formatToStyle`'s `mapBorderStyle` switch, and
 * `TableExtension`'s `BORDER_STYLE_CSS`. They agreed wherever they overlapped
 * and covered different amounts of the enumeration, so the same authored
 * border painted differently depending on where it was drawn: a
 * `thinThickSmallGap` cell edge came out `double` in the editor's DOM and
 * `solid` on the paginated page, and a `dotDash` paragraph rule came out
 * `dashed` through the bridge and `solid` through `borderToStyle`.
 *
 * This is the invariant the single table has to keep, not an example of it:
 * the members come from the committed schema graph, so a refresh that adds one
 * widens the property on its own, and a fourth copy reintroduced anywhere
 * these two entry points reach fails here.
 */

import { readFileSync } from "node:fs";

import { describe, expect, test } from "bun:test";
import fc from "fast-check";

import { BORDER_STYLES, type BorderStyle } from "@stll/docx-core/model";

import { propertyConfig, propertyTestTimeout } from "../../../../test/property-testing";

import { convertBorderSpecToLayout } from "../layout-bridge/convert/toFlowBlocks";
import { CSS_BORDER_STYLES, cssBorderStyle } from "./borderCss";
import { borderToStyle } from "./formatToStyle";

const ST_BORDER_VALUES: readonly string[] = (() => {
  const graph = JSON.parse(
    readFileSync(
      new URL(
        "../../../../specifications/generated/docx-transitional-schema.gen.json",
        import.meta.url,
      ),
      "utf8",
    ),
  ) as { symbols: { kind?: string; name?: string; enumValues?: string[] }[] };
  const border = graph.symbols.find(
    (symbol) => symbol.kind === "simpleType" && symbol.name === "ST_Border",
  );
  if (!border?.enumValues) {
    throw new Error("ST_Border is missing from the schema graph");
  }
  return border.enumValues;
})();

/** `borderToStyle` writes a side-suffixed property bag; read the style out. */
const throughBorderToStyle = (style: BorderStyle): string | undefined => {
  const value = borderToStyle({ style, size: 8 })["borderStyle"];
  return value === undefined ? undefined : String(value);
};

/** The bridge drops a border that paints nothing, so `undefined` means `none`. */
const throughLayoutBridge = (style: BorderStyle): string | undefined =>
  convertBorderSpecToLayout({ style, size: 8 })?.style;

describe("border style rendering", () => {
  test("the table is total over the enumeration", () => {
    expect(Object.keys(CSS_BORDER_STYLES).toSorted()).toEqual([...ST_BORDER_VALUES].toSorted());
    expect([...BORDER_STYLES]).toEqual([...ST_BORDER_VALUES]);
  });

  test(
    "every member renders the same through the bridge and the DOM helpers",
    () => {
      fc.assert(
        fc.property(fc.constantFrom(...BORDER_STYLES), (style) => {
          const painted = cssBorderStyle(style);
          // The bridge returns no border rather than a `none` style, and
          // `borderToStyle` returns an empty bag; both spell "paints nothing".
          const expected = painted === "none" ? undefined : painted;
          expect(throughLayoutBridge(style)).toBe(expected);
          expect(throughBorderToStyle(style)).toBe(expected);
        }),
        propertyConfig({ numRuns: 250 }),
      );
    },
    propertyTestTimeout(15_000),
  );

  test("a token outside the enumeration paints a plain line", () => {
    const unrecognised = { kind: "unrecognised", raw: "notAStyle" } as const;
    expect(cssBorderStyle(unrecognised)).toBe("solid");
    expect(convertBorderSpecToLayout({ style: unrecognised, size: 8 })?.style).toBe("solid");
  });
});
