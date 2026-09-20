/**
 * `w:themeColor` and `w:themeFill` are `ST_ThemeColor`, spelled `dark1`,
 * `hyperlink`, `background1`. The model's union spelled six of its members the
 * DrawingML way (`dk1`, `hlink`) and omitted six WordprocessingML ones plus
 * `none`, so `narrowEnum` returned `undefined` for every token Word actually
 * writes for a hyperlink colour and the attribute was dropped at parse time —
 * and therefore at save time.
 *
 * The enumeration now comes from the committed schema graph, the same way
 * `ST_Border`'s does, so a schema refresh widens these properties on its own.
 */

import { readFileSync } from "node:fs";

import {
  SCHEME_COLOR_VALUE_BY_THEME_COLOR,
  THEME_COLOR_BY_SCHEME_COLOR_VALUE,
  THEME_COLORS,
  themeColorSlot,
} from "@stll/docx-core/model";
import { describe, expect, test } from "bun:test";
import fc from "fast-check";

import { propertyConfig, propertyTestTimeout } from "../../../../test/property-testing";
import type { Theme } from "../types/document";
import { resolveColor } from "../utils/colorResolver";

import { parseBorderSpec } from "./borderParser";
import { parseRunProperties } from "./runParser";
import { serializeBorder } from "./serializer/borderSerializer";
import { serializeShading, serializeTextFormatting } from "./serializer/textFormattingSerializer";
import { parseShading } from "./shadingParser";
import { parseXmlDocument } from "./xmlParser";

/** The enumeration, read from the committed graph rather than restated. */
const ST_THEME_COLOR_VALUES: readonly string[] = (() => {
  const graph = JSON.parse(
    readFileSync(
      new URL(
        "../../../../specifications/generated/docx-transitional-schema.gen.json",
        import.meta.url,
      ),
      "utf8",
    ),
  ) as { symbols: { kind?: string; name?: string; enumValues?: string[] }[] };
  const themeColor = graph.symbols.find(
    (symbol) => symbol.kind === "simpleType" && symbol.name === "ST_ThemeColor",
  );
  if (!themeColor?.enumValues) {
    throw new Error("ST_ThemeColor is missing from the schema graph");
  }
  return themeColor.enumValues;
})();

const WORD_NAMESPACE = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

const parseOne = (xml: string) => {
  const element = parseXmlDocument(xml);
  if (!element) {
    throw new Error("fixture did not parse");
  }
  return element;
};

/**
 * The three carriers the census recorded as losing the attribute: a border
 * side, `w:shd`'s fill, and `w:color` on a run.
 */
const CARRIERS = {
  "w:top@themeColor": (token: string) => {
    const spec = parseBorderSpec(
      parseOne(`<w:top ${WORD_NAMESPACE} w:val="single" w:sz="8" w:color="0563C1"
                 w:themeColor="${token}" w:themeTint="66"/>`),
    );
    return {
      color: spec?.color,
      saved: spec === undefined ? "" : serializeBorder(spec, "top"),
    };
  },
  "w:shd@themeFill": (token: string) => {
    const shading = parseShading(
      parseOne(`<w:shd ${WORD_NAMESPACE} w:val="clear" w:fill="0563C1"
                 w:themeFill="${token}" w:themeFillTint="66"/>`),
    );
    return { color: shading?.fill, saved: serializeShading(shading) };
  },
  "w:color@themeColor": (token: string) => {
    const formatting = parseRunProperties(
      parseOne(`<w:rPr ${WORD_NAMESPACE}><w:color w:val="0563C1"
                 w:themeColor="${token}" w:themeTint="66"/></w:rPr>`),
      null,
    );
    return { color: formatting?.color, saved: serializeTextFormatting(formatting) };
  },
} as const;

const CARRIER_NAMES = Object.keys(CARRIERS) as (keyof typeof CARRIERS)[];

/** The attribute a carrier writes the token into. */
const SAVED_ATTRIBUTE = {
  "w:top@themeColor": "w:themeColor",
  "w:shd@themeFill": "w:themeFill",
  "w:color@themeColor": "w:themeColor",
} as const satisfies Record<keyof typeof CARRIERS, string>;

const THEME: Theme = {
  name: "Test",
  colorScheme: {
    dk1: "111111",
    lt1: "FFFFFF",
    dk2: "222222",
    lt2: "EEEEEE",
    accent1: "AA0001",
    accent2: "AA0002",
    accent3: "AA0003",
    accent4: "AA0004",
    accent5: "AA0005",
    accent6: "AA0006",
    hlink: "0563C1",
    folHlink: "954F72",
  },
};

describe("ST_ThemeColor members", () => {
  test("the model's enumeration is the schema's, member for member", () => {
    expect([...THEME_COLORS].toSorted()).toEqual([...ST_THEME_COLOR_VALUES].toSorted());
  });

  test("a hyperlink theme colour survives a save", () => {
    for (const carrier of CARRIER_NAMES) {
      const { saved } = CARRIERS[carrier]("hyperlink");
      expect(saved).toContain(`${SAVED_ATTRIBUTE[carrier]}="hyperlink"`);
    }
  });

  test("a hyperlink theme colour resolves to the theme's hlink slot", () => {
    const { color } = CARRIERS["w:color@themeColor"]("hyperlink");
    expect(themeColorSlot("hyperlink")).toBe("hlink");
    expect(resolveColor({ themeColor: "hyperlink" }, THEME)).toBe("#0563C1");
    expect(color?.themeColor).toBe("hyperlink");
  });

  test("the mapped members follow w:clrSchemeMapping's default pairs", () => {
    expect(themeColorSlot("text1")).toBe("dk1");
    expect(themeColorSlot("background1")).toBe("lt1");
    expect(themeColorSlot("text2")).toBe("dk2");
    expect(themeColorSlot("background2")).toBe("lt2");
    expect(resolveColor({ themeColor: "text1" }, THEME)).toBe("#111111");
  });

  test("none names no slot, so it paints the element's own colour", () => {
    expect(themeColorSlot("none")).toBeUndefined();
    expect(resolveColor({ themeColor: "none", rgb: "0563C1" }, THEME)).toBe("#0563C1");
  });

  test("the two DrawingML spelling maps are mutual inverses", () => {
    for (const themeColor of THEME_COLORS) {
      const schemeValue = SCHEME_COLOR_VALUE_BY_THEME_COLOR[themeColor];
      if (schemeValue === null) {
        continue;
      }
      expect(THEME_COLOR_BY_SCHEME_COLOR_VALUE[schemeValue]).toBe(themeColor);
    }
  });

  test(
    "every member survives a save unchanged, on every carrier",
    () => {
      fc.assert(
        fc.property(
          fc.constantFrom(...ST_THEME_COLOR_VALUES),
          fc.constantFrom(...CARRIER_NAMES),
          (token, carrier) => {
            const { color, saved } = CARRIERS[carrier](token);
            expect(color?.themeColor).toBe(token);
            expect(saved).toContain(`${SAVED_ATTRIBUTE[carrier]}="${token}"`);
          },
        ),
        propertyConfig({ numRuns: 200 }),
      );
    },
    propertyTestTimeout(15_000),
  );

  test(
    "a token outside the enumeration survives verbatim, unpainted",
    () => {
      fc.assert(
        fc.property(
          fc
            .string({ minLength: 1, maxLength: 24 })
            .filter((value) => /^[A-Za-z][A-Za-z0-9]*$/u.test(value))
            .filter((value) => !ST_THEME_COLOR_VALUES.includes(value)),
          fc.constantFrom(...CARRIER_NAMES),
          (token, carrier) => {
            const { color, saved } = CARRIERS[carrier](token);
            expect(color?.themeColor).toEqual({ kind: "unrecognised", raw: token });
            expect(saved).toContain(`${SAVED_ATTRIBUTE[carrier]}="${token}"`);
          },
        ),
        propertyConfig(),
      );
    },
    propertyTestTimeout(15_000),
  );
});
