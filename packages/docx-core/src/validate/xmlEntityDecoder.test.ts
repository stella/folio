import { expect, test } from "bun:test";
import fc from "fast-check";
import { propertyConfig, propertyTestTimeout } from "../../../../test/property-testing";
import { createXmlEntityDecoder } from "./xmlEntityDecoder";

test(
  "XML characters preserve their value across literal, decimal, and hex encodings",
  () => {
    fc.assert(
      fc.property(
        fc.oneof(
          fc.constantFrom(0x09, 0x0a, 0x0d),
          fc.integer({ min: 0x20, max: 0xd7ff }),
          fc.integer({ min: 0xe000, max: 0xfffd }),
          fc.integer({ min: 0x10000, max: 0x10ffff }),
        ),
        (codePoint) => {
          const decoder = createXmlEntityDecoder();
          const character = String.fromCodePoint(codePoint);
          expect(decoder.decode(`&#${codePoint};`)).toBe(character);
          expect(decoder.decode(`&#x${codePoint.toString(16)};`)).toBe(character);
          expect(decoder.decode(character)).toBe(character);
        },
      ),
      propertyConfig({ numRuns: 100 }),
    );
  },
  propertyTestTimeout(5_000),
);

test("XML names decode once while HTML-only names remain unexpanded", () => {
  const decoder = createXmlEntityDecoder();
  expect(decoder.decode("&lt;&gt;&amp;&quot;&apos;")).toBe("<>&\"'");
  expect(decoder.decode("&amp;#49;&#38;#49;")).toBe("&#49;&#49;");
  expect(decoder.decode("&nbsp;&copy;&eacute;")).toBe("&nbsp;&copy;&eacute;");
});

test("the decoder rejects document-defined entities", () => {
  const decoder = createXmlEntityDecoder();
  expect(() => decoder.addInputEntities({ custom: "replacement" })).toThrow();
});
