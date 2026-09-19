/**
 * `wp:anchor/@behindDoc` decides whether an anchored object paints behind the
 * body text. It is xsd:boolean, so `1`, `0`, `true` and `false` are all legal;
 * Word writes `1`, other producers write `true`. A watermark text box written
 * with `behindDoc="true"` used to paint in front of the document, because the
 * shape and text-box tier read the attribute with its own `=== "1"` test while
 * the image tier used the shared on/off reader.
 *
 * `parseAnchorBehindDoc` is now the only reader. These properties hold over
 * every legal spelling and over both tiers.
 */

import { describe, expect, test } from "bun:test";
import fc from "fast-check";

import { propertyConfig, propertyTestTimeout } from "../../../../test/property-testing";

import { parseDrawing } from "./imageParser";
import { parseTextBox } from "./textBoxParser";
import { parseXmlDocument } from "./xmlParser";
import type { XmlElement } from "./xmlParser";

const NAMESPACES = [
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"',
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"',
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"',
  'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"',
  'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"',
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',
].join(" ");

/** `behindDoc` omitted entirely is spelled as `null` here. */
const BEHIND_DOC_SPELLINGS = ["1", "0", "true", "false", null] as const;

const parse = (xml: string): XmlElement => {
  const element = parseXmlDocument(xml);
  if (!element) {
    throw new Error("fixture did not parse");
  }
  return element;
};

const anchorAttributes = (behindDoc: string | null): string =>
  `distT="0" distB="0" distL="0" distR="0"${behindDoc === null ? "" : ` behindDoc="${behindDoc}"`}`;

const textBoxDrawing = (behindDoc: string | null): XmlElement =>
  parse(`
    <w:drawing ${NAMESPACES}>
      <wp:anchor ${anchorAttributes(behindDoc)}>
        <wp:extent cx="1000000" cy="500000"/>
        <wp:docPr id="1" name="TextBox 1"/>
        <a:graphic>
          <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
            <wps:wsp>
              <wps:spPr/>
              <wps:txbx><w:txbxContent><w:p/></w:txbxContent></wps:txbx>
              <wps:bodyPr/>
            </wps:wsp>
          </a:graphicData>
        </a:graphic>
      </wp:anchor>
    </w:drawing>
  `);

const imageDrawing = (behindDoc: string | null): XmlElement =>
  parse(`
    <w:drawing ${NAMESPACES}>
      <wp:anchor ${anchorAttributes(behindDoc)}>
        <wp:extent cx="1000000" cy="500000"/>
        <wp:docPr id="2" name="Picture 1"/>
        <a:graphic>
          <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
            <pic:pic>
              <pic:blipFill><a:blip r:embed="rId1"/></pic:blipFill>
              <pic:spPr/>
            </pic:pic>
          </a:graphicData>
        </a:graphic>
      </wp:anchor>
    </w:drawing>
  `);

const textBoxWrapType = (behindDoc: string | null): string | undefined =>
  parseTextBox(textBoxDrawing(behindDoc))?.wrap?.type;

const imageWrapType = (behindDoc: string | null): string | undefined =>
  parseDrawing(imageDrawing(behindDoc))?.wrap?.type;

describe("the wp:anchor behindDoc reader", () => {
  test('a text box anchored with behindDoc="true" paints behind the text', () => {
    expect(textBoxWrapType("true")).toBe("behind");
  });

  test(
    "both tiers agree on every legal spelling",
    () => {
      fc.assert(
        fc.property(fc.constantFrom(...BEHIND_DOC_SPELLINGS), (behindDoc) => {
          const expected = behindDoc === "1" || behindDoc === "true" ? "behind" : "inFront";

          expect(textBoxWrapType(behindDoc)).toBe(expected);
          expect(imageWrapType(behindDoc)).toBe(expected);
        }),
        propertyConfig(),
      );
    },
    propertyTestTimeout(10_000),
  );
});
