import { describe, expect, test } from "bun:test";

import { parseEndnoteProperties, parseFootnoteProperties } from "./notePropertiesParser";
import { serializeSectionProperties } from "./serializer/sectionPropertiesSerializer";
import { findChild, parseXmlDocument } from "./xmlParser";

describe("note properties number formats", () => {
  test("keeps specialized OOXML number formats", () => {
    const root = parseXmlDocument(
      `<w:sectPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:footnotePr>
          <w:numFmt w:val="hindiCounting" w:format="authored-format"/>
        </w:footnotePr>
      </w:sectPr>`,
    );

    const properties = parseFootnoteProperties(findChild(root, "w", "footnotePr"));

    expect(properties.numFmt).toBe("hindiCounting");
    expect(properties.numFmtFormat).toBe("authored-format");
    expect(serializeSectionProperties({ footnotePr: properties })).toContain(
      '<w:numFmt w:val="hindiCounting" w:format="authored-format"/>',
    );
  });

  test("keeps endnote format metadata and escapes it on save", () => {
    const root = parseXmlDocument(
      `<w:sectPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:endnotePr>
          <w:numFmt w:val="lowerRoman" w:format="chapter&amp;&quot;number"/>
        </w:endnotePr>
      </w:sectPr>`,
    );

    const properties = parseEndnoteProperties(findChild(root, "w", "endnotePr"));

    expect(properties.numFmtFormat).toBe('chapter&"number');
    expect(serializeSectionProperties({ endnotePr: properties })).toContain(
      '<w:numFmt w:val="lowerRoman" w:format="chapter&amp;&quot;number"/>',
    );
  });
});
