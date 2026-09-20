/**
 * `word/fontTable.xml` survives a rebuild from the model alone.
 *
 * A repack copies the part across byte for byte, so nothing here runs on an
 * ordinary save and the reader's gaps were invisible: the survival census
 * reported every pair in the part unmeasured rather than lost. The census now
 * forces the part serializer, and this is the same assertion in the package
 * that owns the reader, so a regression fails `bun test packages/core` rather
 * than waiting for a full sweep.
 *
 * The property is over `CT_Font`'s declared children rather than over
 * examples: the generated `CONTAINER_CHILDREN` set is the schema's, so a child
 * a schema refresh adds arrives in the test without anybody remembering to add
 * it, and the order assertion below fails until somebody places it.
 */

import { describe, expect, test } from "bun:test";
import fc from "fast-check";

import { propertyConfig } from "../../../../test/property-testing";

import { CONTAINER_CHILDREN } from "./containerChildren.gen";
import { parseFontTable } from "./fontTableParser";
import { serializeFontTableXml } from "./serializer/fontTableSerializer";
import { getChildElements, getLocalName, parseXmlDocument } from "./xmlParser";

const FONT_KEY = "{001B70DC-AA60-4AD5-90EC-18A0948E1EAE}";

/**
 * One instance of each child `CT_Font` declares, in the order it declares them.
 *
 * Written out rather than generated, because each needs the attributes its own
 * type requires; bound to the generated set by the first test, so it cannot
 * quietly fall behind the schema.
 */
const FONT_CHILDREN: ReadonlyArray<readonly [name: string, xml: string]> = [
  ["altName", '<w:altName w:val="Arial"/>'],
  ["panose1", '<w:panose1 w:val="020B0604020202020204"/>'],
  ["charset", '<w:charset w:val="00" w:characterSet="ANSI_CHARSET"/>'],
  ["family", '<w:family w:val="swiss"/>'],
  ["notTrueType", "<w:notTrueType/>"],
  ["pitch", '<w:pitch w:val="variable"/>'],
  [
    "sig",
    '<w:sig w:usb0="E0002EFF" w:usb1="C000785B" w:usb2="00000009" w:usb3="00000000"' +
      ' w:csb0="000001FF" w:csb1="00000000"/>',
  ],
  ["embedRegular", `<w:embedRegular r:id="rId1" w:fontKey="${FONT_KEY}" w:subsetted="true"/>`],
  ["embedBold", `<w:embedBold r:id="rId2" w:fontKey="${FONT_KEY}"/>`],
  ["embedItalic", '<w:embedItalic r:id="rId3"/>'],
  ["embedBoldItalic", `<w:embedBoldItalic r:id="rId4" w:subsetted="false"/>`],
];

const partWith = (fonts: string, rootAttributes = ""): string =>
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  `<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"` +
  ` xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"` +
  ` xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"` +
  ` xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"${rootAttributes}>` +
  `${fonts}</w:fonts>`;

const fontWith = (children: string, attributes = ""): string =>
  partWith(`<w:font w:name="Arial"${attributes}>${children}</w:font>`);

const rebuild = (xml: string): string => {
  const table = parseFontTable(xml);
  if (table === undefined) {
    throw new Error(`the reader found no font table in ${xml}`);
  }
  return serializeFontTableXml(table);
};

/** The child element names one `w:font` carries, in document order. */
const childNamesOf = (xml: string): string[] => {
  const root = parseXmlDocument(xml);
  if (!root) {
    throw new Error(`the rebuilt part did not parse: ${xml}`);
  }
  return getChildElements(root)
    .flatMap((font) => getChildElements(font))
    .map((node) => getLocalName(node.name));
};

describe("the font table's reader covers what the schema declares", () => {
  test("every declared child of CT_Font has a case here", () => {
    expect(FONT_CHILDREN.map(([name]) => name).toSorted()).toEqual([
      ...CONTAINER_CHILDREN["w:font"],
    ]);
  });
});

describe("a font survives a rebuild from the model", () => {
  test("any subset of CT_Font's children comes back, in place", () => {
    fc.assert(
      fc.property(fc.subarray([...FONT_CHILDREN], { minLength: 1 }), (children) => {
        const source = fontWith(children.map(([, xml]) => xml).join(""));
        const rebuilt = rebuild(source);
        expect(childNamesOf(rebuilt)).toEqual(children.map(([name]) => name));
      }),
      propertyConfig({ numRuns: 60 }),
    );
  });

  test("a rebuilt part is a fixed point of the reader", () => {
    fc.assert(
      fc.property(fc.subarray([...FONT_CHILDREN], { minLength: 1 }), (children) => {
        const source = fontWith(children.map(([, xml]) => xml).join(""));
        const once = parseFontTable(rebuild(source));
        const twice = parseFontTable(rebuild(rebuild(source)));
        expect(twice).toEqual(once);
      }),
      propertyConfig({ numRuns: 60 }),
    );
  });
});

/**
 * The four cases the part-rebuild law found, each of which used to be a drop.
 *
 * They are named apart from the property because each is a different reason:
 * a child the reader had no case for, a child it read as nothing, an attribute
 * its record had no field for, and an element the model held and the writer
 * never emitted.
 */
describe("what a rebuild used to lose", () => {
  test("an element the model has no field for keeps its bytes", () => {
    expect(rebuild(fontWith("<w:notTrueType/>"))).toContain("<w:notTrueType/>");
  });

  test("a bare w:charset is not an absent one", () => {
    expect(rebuild(fontWith("<w:charset/>"))).toContain("<w:charset/>");
  });

  test("w:charset keeps the character set it names as well as the one it numbers", () => {
    const rebuilt = rebuild(fontWith('<w:charset w:val="00" w:characterSet="ANSI_CHARSET"/>'));
    expect(rebuilt).toContain('w:val="00"');
    expect(rebuilt).toContain('w:characterSet="ANSI_CHARSET"');
  });

  test("an embedded face keeps its relationship, its key and its subsetting", () => {
    const rebuilt = rebuild(
      fontWith(`<w:embedRegular r:id="rId1" w:fontKey="${FONT_KEY}" w:subsetted="true"/>`),
    );
    expect(rebuilt).toContain('r:id="rId1"');
    expect(rebuilt).toContain(`w:fontKey="${FONT_KEY}"`);
    expect(rebuilt).toContain('w:subsetted="1"');
  });

  test("an unmodelled attribute of w:font rides the element's remainder", () => {
    expect(rebuild(fontWith("", ' w14:unknown="7"'))).toContain('w14:unknown="7"');
  });

  test("an unmodelled child of w:fonts keeps its place among the fonts", () => {
    const rebuilt = rebuild(
      partWith('<w:font w:name="A"/><w:someLaterElement/><w:font w:name="B"/>'),
    );
    expect(rebuilt.indexOf('w:name="A"')).toBeLessThan(rebuilt.indexOf("<w:someLaterElement/>"));
    expect(rebuilt.indexOf("<w:someLaterElement/>")).toBeLessThan(rebuilt.indexOf('w:name="B"'));
  });
});

/**
 * `mc:Ignorable` is stated by the rebuilt root, never replayed from the source.
 *
 * It lists which of the root's own `xmlns:*` bindings a consumer may skip, and
 * a rebuilt root binds the prefixes its body uses. Replaying the source's
 * would name prefixes nothing binds and would be written twice, because the
 * part writer appends its own.
 */
describe("a part root states its ignorable prefixes rather than replaying them", () => {
  test("a source mc:Ignorable does not reach the rebuilt root", () => {
    const rebuilt = rebuild(partWith('<w:font w:name="Arial"/>', ' mc:Ignorable="w14 w15"'));
    expect(rebuilt).not.toContain('mc:Ignorable="w14 w15"');
  });

  test("another unmodelled root attribute does", () => {
    const rebuilt = rebuild(partWith('<w:font w:name="Arial"/>', ' w:unknownRootFlag="1"'));
    expect(rebuilt).toContain('w:unknownRootFlag="1"');
  });
});
