/**
 * `ST_OnOff` has three spellings per polarity — `1`/`true`/`on` and
 * `0`/`false`/`off` — and Word writes all of them. A hand-rolled
 * `=== "1" || === "true"` reads `w:beforeAutospacing="on"` as *false*, and the
 * save path then writes `w:beforeAutospacing="0"`: the document comes back
 * saying the opposite of what its author said.
 *
 * `parseOnOffAttribute` is the one reader. These properties hold over the whole
 * lexical space, over the two tiers that carry the same slots (a paragraph's
 * `w:pPr` and a style's `w:pPr`), and across the save path.
 */

import { describe, expect, test } from "bun:test";
import fc from "fast-check";

import { propertyConfig, propertyTestTimeout } from "../../../../test/property-testing";

import { serializeParagraphFormatting } from "./serializer/paragraphSerializer";
import { parseParagraphProperties } from "./paragraphParser";
import { parseStyles } from "./styleParser";
import { parseXmlDocument } from "./xmlParser";

const WORD_NAMESPACE = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

/** Every `ST_OnOff` lexical value, plus the attribute being absent (`null`). */
const ON_OFF_SPELLINGS = ["1", "true", "on", "0", "false", "off", null] as const;

type OnOffSpelling = (typeof ON_OFF_SPELLINGS)[number];

const expectedValue = (spelling: OnOffSpelling): boolean | undefined => {
  if (spelling === null) {
    return undefined;
  }
  return spelling === "1" || spelling === "true" || spelling === "on";
};

const attribute = (name: string, spelling: OnOffSpelling): string =>
  spelling === null ? "" : ` w:${name}="${spelling}"`;

const paragraphPropertiesXml = (spelling: OnOffSpelling): string => `
  <w:pPr ${WORD_NAMESPACE}>
    <w:spacing w:before="120" w:after="120"${attribute("beforeAutospacing", spelling)}/>
    <w:pBdr><w:top w:val="single" w:sz="4"${attribute("shadow", spelling)}/></w:pBdr>
  </w:pPr>
`;

const parseParagraphTier = (spelling: OnOffSpelling) => {
  const pPr = parseXmlDocument(paragraphPropertiesXml(spelling));
  if (!pPr) {
    throw new Error("fixture did not parse");
  }
  return parseParagraphProperties(pPr, null);
};

const parseStyleTier = (spelling: OnOffSpelling) => {
  const styles = parseStyles(
    `<w:styles ${WORD_NAMESPACE}>
       <w:style w:type="paragraph" w:styleId="Spaced">
         <w:name w:val="Spaced"/>
         ${paragraphPropertiesXml(spelling).replace(` ${WORD_NAMESPACE}`, "")}
       </w:style>
     </w:styles>`,
    null,
  );
  return styles.get("Spaced")?.pPr;
};

describe("ST_OnOff attributes", () => {
  test('w:beforeAutospacing="on" survives a save instead of inverting', () => {
    const formatting = parseParagraphTier("on");

    expect(formatting?.beforeAutospacing).toBe(true);
    expect(serializeParagraphFormatting(formatting ?? {})).toContain('w:beforeAutospacing="1"');
  });

  test(
    "every spelling reads the same on a paragraph and on a style",
    () => {
      fc.assert(
        fc.property(fc.constantFrom(...ON_OFF_SPELLINGS), (spelling) => {
          const expected = expectedValue(spelling);
          const paragraph = parseParagraphTier(spelling);
          const style = parseStyleTier(spelling);

          expect(paragraph?.beforeAutospacing).toBe(expected);
          expect(style?.beforeAutospacing).toBe(expected);
          expect(paragraph?.borders?.top?.shadow).toBe(expected);
          expect(style?.borders?.top?.shadow).toBe(expected);
        }),
        propertyConfig(),
      );
    },
    propertyTestTimeout(10_000),
  );

  test(
    "a saved paragraph re-reads as the value it was parsed from",
    () => {
      fc.assert(
        fc.property(fc.constantFrom(...ON_OFF_SPELLINGS), (spelling) => {
          const parsed = parseParagraphTier(spelling);
          const saved = parseXmlDocument(
            serializeParagraphFormatting(parsed ?? {}).replace(
              "<w:pPr>",
              `<w:pPr ${WORD_NAMESPACE}>`,
            ),
          );
          if (!saved) {
            throw new Error("serialized properties did not parse");
          }

          expect(parseParagraphProperties(saved, null)?.beforeAutospacing).toBe(
            expectedValue(spelling),
          );
        }),
        propertyConfig(),
      );
    },
    propertyTestTimeout(10_000),
  );
});
