/**
 * A property stated twice in one `w:rPr` resolves to the last statement, for
 * every owner of a run property set.
 *
 * `EG_RPrBase` is an `xsd:choice` referenced `maxOccurs="unbounded"`, so the
 * repeat is valid markup and producers write it; the schema says nothing about
 * which statement wins, and folio's two readers used to answer differently —
 * `runParser.ts` took the first, `styleParser.ts` the last. One reader now
 * serves all of them and the rule is stated once, in
 * `docs/reserved-values.md` and above the reader.
 *
 * Three things are pinned here, over the whole class of repeated children
 * rather than over one example:
 *
 *   1. the later statement is the one that reaches the model;
 *   2. the statements it beat are not written back, so the saved element
 *      states the property once and a first-wins consumer and a last-wins
 *      consumer read the same value from it;
 *   3. saving twice changes nothing more.
 *
 * A run, the paragraph mark and a style are all exercised, because one reader
 * with one rule is exactly the claim that they cannot answer differently.
 */

import { describe, expect, test } from "bun:test";
import fc from "fast-check";

import { propertyConfig } from "../../../../test/property-testing";

import { parseParagraph } from "./paragraphParser";
import { parseRun, parseRunProperties, RUN_PROPERTY_OWNERS } from "./runParser";
import { serializeParagraphFormatting } from "./serializer/paragraphSerializer";
import { serializeRun } from "./serializer/runSerializer";
import { parseStyles } from "./styleParser";
import { parseXmlDocument, type XmlElement } from "./xmlParser";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

/**
 * A property stated two ways, with the value each statement carries.
 *
 * Every entry disagrees with itself: a repeat whose halves agree cannot tell
 * first-wins and last-wins apart, and 3131 of the 3132 repeats in the public
 * corpus are of that harmless kind.
 */
const DISAGREEING_REPEATS = [
  { name: "b", field: "bold", first: "<w:b/>", last: '<w:b w:val="0"/>', value: false },
  { name: "b", field: "bold", first: '<w:b w:val="0"/>', last: "<w:b/>", value: true },
  { name: "i", field: "italic", first: "<w:i/>", last: '<w:i w:val="false"/>', value: false },
  {
    name: "sz",
    field: "fontSize",
    first: '<w:sz w:val="24"/>',
    last: '<w:sz w:val="26"/>',
    value: 26,
  },
  {
    name: "szCs",
    field: "fontSizeCs",
    first: '<w:szCs w:val="26"/>',
    last: '<w:szCs w:val="24"/>',
    value: 24,
  },
  {
    name: "caps",
    field: "allCaps",
    first: "<w:caps/>",
    last: '<w:caps w:val="0"/>',
    value: false,
  },
  {
    name: "rStyle",
    field: "styleId",
    first: '<w:rStyle w:val="Early"/>',
    last: '<w:rStyle w:val="Late"/>',
    value: "Late",
  },
  {
    name: "highlight",
    field: "highlight",
    first: '<w:highlight w:val="cyan"/>',
    last: '<w:highlight w:val="yellow"/>',
    value: "yellow",
  },
  {
    name: "position",
    field: "position",
    first: '<w:position w:val="4"/>',
    last: '<w:position w:val="8"/>',
    value: 8,
  },
] as const;

const parseOne = (xml: string): XmlElement => {
  const root = parseXmlDocument(xml) as XmlElement | null;
  if (!root) {
    throw new Error("the fixture did not parse");
  }
  return root;
};

/** A run's own `w:rPr`, parsed and written back from the model alone. */
const rebuildRun = (properties: string): string =>
  serializeRun(
    parseRun(
      parseOne(`<w:r xmlns:w="${W}"><w:rPr>${properties}</w:rPr><w:t>x</w:t></w:r>`),
      null,
      null,
    ),
  );

/** The paragraph mark's `w:rPr`, parsed and written back from the model alone. */
const rebuildParagraphMark = (properties: string): string => {
  const paragraph = parseParagraph(
    parseOne(`<w:p xmlns:w="${W}"><w:pPr><w:rPr>${properties}</w:rPr></w:pPr></w:p>`),
    null,
    null,
    null,
  );
  return serializeParagraphFormatting(
    paragraph.formatting,
    paragraph.propertyChanges,
    paragraph.pPrMark,
  );
};

/** How many times the saved markup states this child. */
const occurrences = (saved: string, name: string): number =>
  saved.match(new RegExp(`<w:${name}[ />]`, "gu"))?.length ?? 0;

/** The outer `w:rPr`'s own children: a `w:rPrChange` nests a second one. */
const propertiesOf = (saved: string): string =>
  saved.slice(saved.indexOf("<w:rPr>") + "<w:rPr>".length, saved.lastIndexOf("</w:rPr>"));

const repeat = fc.constantFrom(...DISAGREEING_REPEATS);

describe("a property stated twice in one w:rPr", () => {
  test("resolves to the last statement, on a run and on a style alike", () => {
    fc.assert(
      fc.property(repeat, ({ field, first, last, value }) => {
        const properties = `${first}${last}`;

        const run = parseRunProperties(
          parseOne(`<w:rPr xmlns:w="${W}">${properties}</w:rPr>`),
          null,
          RUN_PROPERTY_OWNERS.standalone,
        );
        const style = parseStyles(
          `<w:styles xmlns:w="${W}"><w:style w:type="character" w:styleId="S">` +
            `<w:rPr>${properties}</w:rPr></w:style></w:styles>`,
          null,
        ).get("S")?.rPr;

        expect(run?.[field]).toBe(value);
        expect(style?.[field]).toBe(value);
      }),
      propertyConfig({ numRuns: 100 }),
    );
  });

  test("comes back stated once, carrying the value that won", () => {
    fc.assert(
      fc.property(
        repeat,
        fc.constantFrom(rebuildRun, rebuildParagraphMark),
        ({ name, field, first, last, value }, rebuild) => {
          const saved = rebuild(`${first}${last}`);
          const savedProperties = propertiesOf(saved);

          expect(occurrences(saved, name)).toBe(1);
          expect(
            parseRunProperties(
              parseOne(`<w:rPr xmlns:w="${W}">${savedProperties}</w:rPr>`),
              null,
              RUN_PROPERTY_OWNERS.standalone,
            )?.[field],
          ).toBe(value);
          expect(rebuild(savedProperties)).toBe(saved);
        },
      ),
      propertyConfig({ numRuns: 100 }),
    );
  });

  test("a repeat beside a captured child leaves the capture where it stood", () => {
    // The sink records a schema ordinal, and the winner takes the slot its
    // name owns; a capture between the two statements keeps its own.
    const saved = rebuildRun('<w:b/><w:bdr w:val="single" w:sz="4"/><w:b w:val="0"/>');

    expect(saved).toContain('<w:b w:val="0"/><w:bdr w:val="single" w:sz="4"/>');
    expect(occurrences(saved, "b")).toBe(1);
  });

  test("the paragraph mark projects only the last run-in marker", () => {
    const parseMark = (properties: string) =>
      parseParagraph(
        parseOne(`<w:p xmlns:w="${W}"><w:pPr><w:rPr>${properties}</w:rPr></w:pPr></w:p>`),
        null,
        null,
        null,
      );

    const offWins = parseMark('<w:specVanish/><w:specVanish w:val="0"/>');
    expect(offWins.formatting?.runInWithNext).toBeUndefined();
    const savedOff = serializeParagraphFormatting(
      offWins.formatting,
      offWins.propertyChanges,
      offWins.pPrMark,
    );
    expect(occurrences(savedOff, "specVanish")).toBe(1);
    expect(savedOff).toContain('<w:specVanish w:val="0"/>');

    const onWins = parseMark('<w:specVanish w:val="0"/><w:specVanish/>');
    expect(onWins.formatting?.runInWithNext).toBe(true);
    const savedOn = serializeParagraphFormatting(
      onWins.formatting,
      onWins.propertyChanges,
      onWins.pPrMark,
    );
    expect(occurrences(savedOn, "specVanish")).toBe(1);
    expect(savedOn).toContain("<w:specVanish/>");
  });
});
