/**
 * `w:val` on `CT_Border` is `ST_Border`, a 193-member enumeration containing
 * two distinct "no border" tokens: `nil` and `none`. Merging them is not
 * cosmetic — an explicit one cancels a border inherited from the container, and
 * Word round-trips whichever the author wrote.
 *
 * Four `parseBorderSpec` copies had drifted (absent `w:val` became `none` on
 * table and page borders but dropped the border on paragraph and style borders;
 * only two recorded an explicit `w:shadow="0"`; only one read the page-border
 * art relationship ids), and `docx-core`'s build-from-scratch serializer
 * collapsed `none` into `nil`.
 *
 * The member list comes from the committed schema graph, not a hand list, so a
 * schema refresh that adds a member widens this property automatically. It is
 * also what `scripts/generate-border-styles.ts` derives the `BorderStyle` union
 * from, so this file checks the two agree.
 */

import { readFileSync } from "node:fs";

import { describe, expect, test } from "bun:test";
import fc from "fast-check";

import { BORDER_STYLES, borderStyleToken, PARSE_WARNING_CODES } from "@stll/docx-core/model";

import { propertyConfig, propertyTestTimeout } from "../../../../test/property-testing";

import { parseBorderSpec } from "./borderParser";
import { createParseWarningCollector } from "./parseContext";
import { parseParagraphProperties } from "./paragraphProperties";
import { parseSectionProperties } from "./sectionParser";
import { serializeBorder } from "./serializer/borderSerializer";
import { parseStyles } from "./styleParser";
import { parseTableCellProperties, parseTableProperties } from "./tableParser";
import { parseXmlDocument } from "./xmlParser";

/**
 * The enumeration, read from the committed schema graph rather than restated:
 * a schema refresh that adds a member widens these properties on its own.
 */
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

const WORD_NAMESPACE = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

const parseOne = (xml: string) => {
  const element = parseXmlDocument(xml);
  if (!element) {
    throw new Error("fixture did not parse");
  }
  return element;
};

const topBorderXml = (style: string, extra = "") =>
  `<w:top w:val="${style}" w:sz="8" w:space="0" w:color="auto"${extra}/>`;

/**
 * The attribute shapes the four copies disagreed about: an explicit off-value
 * for the toggles (two copies recorded only `true`), and the page-border art
 * relationship ids (one copy read them). A border element with no `w:val` at
 * all is covered separately, because two copies dropped it and two invented
 * `none`.
 */
const BORDER_ATTRIBUTE_VARIANTS = [
  "",
  ' w:shadow="0"',
  ' w:shadow="1"',
  ' w:frame="off"',
  ' w:id="rId7"',
  ' w:topLeft="rId8" w:bottomRight="rId9"',
] as const;

/**
 * Each tier's real entry point, not the shared helper: the property is that the
 * containers agree, which only a per-tier read can show.
 */
const TIERS = {
  paragraph: (border: string) =>
    parseParagraphProperties(
      parseOne(`<w:pPr ${WORD_NAMESPACE}><w:pBdr>${border}</w:pBdr></w:pPr>`),
      null,
    )?.borders?.top,
  style: (border: string) =>
    parseStyles(
      `<w:styles ${WORD_NAMESPACE}>
         <w:style w:type="paragraph" w:styleId="Bordered">
           <w:name w:val="Bordered"/>
           <w:pPr><w:pBdr>${border}</w:pBdr></w:pPr>
         </w:style>
       </w:styles>`,
      null,
    ).get("Bordered")?.pPr?.borders?.top,
  table: (border: string) =>
    parseTableProperties(
      parseOne(`<w:tblPr ${WORD_NAMESPACE}><w:tblBorders>${border}</w:tblBorders></w:tblPr>`),
    )?.borders?.top,
  cell: (border: string) =>
    parseTableCellProperties(
      parseOne(`<w:tcPr ${WORD_NAMESPACE}><w:tcBorders>${border}</w:tcBorders></w:tcPr>`),
    )?.borders?.top,
  page: (border: string) =>
    parseSectionProperties(
      parseOne(`<w:sectPr ${WORD_NAMESPACE}><w:pgBorders>${border}</w:pgBorders></w:sectPr>`),
    ).pageBorders?.top,
} as const;

const TIER_NAMES = Object.keys(TIERS) as (keyof typeof TIERS)[];

/** Parse, run the real serializer, parse the result back. */
const roundTrip = (style: string, tier: keyof typeof TIERS) => {
  const parsed = TIERS[tier](topBorderXml(style));
  if (!parsed) {
    throw new Error(`border did not parse on the ${tier} tier`);
  }
  return TIERS[tier](serializeBorder(parsed, "top"));
};

describe("ST_Border members", () => {
  test("the schema graph supplies the whole enumeration", () => {
    expect(ST_BORDER_VALUES.length).toBeGreaterThan(190);
    expect(ST_BORDER_VALUES).toContain("nil");
    expect(ST_BORDER_VALUES).toContain("none");
  });

  test("the generated union is the enumeration, member for member", () => {
    expect([...BORDER_STYLES]).toEqual([...ST_BORDER_VALUES]);
  });

  test("nil and none stay distinct through a save", () => {
    for (const tier of TIER_NAMES) {
      expect(roundTrip("nil", tier)?.style).toBe("nil");
      expect(roundTrip("none", tier)?.style).toBe("none");
    }
  });

  test(
    "every member survives a save unchanged, in every container",
    () => {
      fc.assert(
        fc.property(
          fc.constantFrom(...ST_BORDER_VALUES),
          fc.constantFrom(...TIER_NAMES),
          (style, tier) => {
            expect(roundTrip(style, tier)?.style).toBe(style);
          },
        ),
        propertyConfig({ numRuns: 250 }),
      );
    },
    propertyTestTimeout(15_000),
  );

  test(
    "a token outside the enumeration survives verbatim, as a token",
    () => {
      fc.assert(
        fc.property(
          fc
            .string({ minLength: 1, maxLength: 24 })
            .filter((value) => /^[A-Za-z][A-Za-z0-9]*$/u.test(value))
            .filter((value) => !ST_BORDER_VALUES.includes(value)),
          fc.constantFrom(...TIER_NAMES),
          (style, tier) => {
            const parsed = roundTrip(style, tier)?.style;
            // The tri-state, not a widened `string`: a consumer that switches
            // on the union cannot mistake an undeclared token for a member.
            expect(parsed).toEqual({ kind: "unrecognised", raw: style });
            expect(borderStyleToken(parsed ?? "nil")).toBe(style);
          },
        ),
        propertyConfig(),
      );
    },
    propertyTestTimeout(15_000),
  );

  test("an undeclared token is reported through the parse context", () => {
    const collector = createParseWarningCollector("word/document.xml");
    const element = parseOne(`<w:top ${WORD_NAMESPACE} w:val="apples"/>`);
    parseBorderSpec(element, collector.context);
    expect(collector.warnings()).toEqual([]);

    parseBorderSpec(parseOne(`<w:top ${WORD_NAMESPACE} w:val="notAStyle"/>`), collector.context);
    expect(collector.warnings()).toEqual([
      {
        code: PARSE_WARNING_CODES.borderStyleOutsideEnum,
        location: { part: "word/document.xml", element: "w:top" },
        value: "notAStyle",
        count: 1,
      },
    ]);
  });

  test("a border element with no w:val reads the same on every tier", () => {
    const specs = TIER_NAMES.map((tier) => TIERS[tier]('<w:top w:sz="8" w:space="0"/>'));
    for (const spec of specs) {
      expect(spec).toEqual(specs[0]);
    }
  });

  test(
    "every tier reads a border the same way",
    () => {
      fc.assert(
        fc.property(
          fc.constantFrom(...ST_BORDER_VALUES),
          fc.constantFrom(...BORDER_ATTRIBUTE_VARIANTS),
          (style, extra) => {
            const specs = TIER_NAMES.map((tier) => TIERS[tier](topBorderXml(style, extra)));
            for (const spec of specs) {
              expect(spec).toEqual(specs[0]);
            }
          },
        ),
        propertyConfig({ numRuns: 250 }),
      );
    },
    propertyTestTimeout(15_000),
  );
});
