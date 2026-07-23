import { describe, expect, test } from "bun:test";
import fc from "fast-check";
import JSZip from "jszip";
import { readFileSync } from "node:fs";
import { resolve } from "node:path";

import { parseStreamingXml } from "./streamingXmlParser";
import { parseXml } from "./xmlParser";

const REPO_ROOT = resolve(import.meta.dir, "../../../..");
const DOCUMENT_FIXTURE_GLOBS = [
  "tests/visual/fixtures/*.docx",
  "packages/core/src/docx/__fixtures__/*.docx",
  "packages/core/src/docx/__tests__/__fixtures__/**/*.docx",
] as const;

const XML_CASES = [
  `<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="urn:w">
  <!-- ignored -->
  <w:body>
    <w:p data-one="&quot;&amp;&#65;&#x42;" data-two='apostrophe'>
      <w:r><w:t xml:space="preserve"> text &lt; value </w:t></w:r>
    </w:p>
  </w:body>
</w:document>`,
  "<document>\r\n  <body><p><![CDATA[raw <xml>\r\n&\rtext]]></p></body>\r</document>",
  "<document><text><![CDATA[first]]><!-- ignored -->second</text></document>",
  "<x:document><x:body><x:p/><x:tbl><x:tr><x:tc/></x:tr></x:tbl></x:body></x:document>",
] as const;

describe("parseStreamingXml", () => {
  test.each(XML_CASES)("matches the compatibility parser", (xml) => {
    expect(parseStreamingXml(xml)).toEqual({
      status: "parsed",
      value: parseXml(xml),
    });
  });

  test("falls back for declarations and malformed nesting outside its contract", () => {
    const deeplyNested = `<document>${"<x>".repeat(101)}value${"</x>".repeat(101)}</document>`;

    expect(parseStreamingXml("<!DOCTYPE document><document/>").status).toBe("unsupported");
    expect(parseStreamingXml("<document><body></document>").status).toBe("unsupported");
    expect(parseStreamingXml("<document>&custom;</document>").status).toBe("unsupported");
    expect(parseStreamingXml('<document __proto__="unsafe"/>').status).toBe("unsupported");
    expect(parseStreamingXml(deeplyNested).status).toBe("unsupported");
  });

  test("matches entity and line-ending behavior for generated values", () => {
    const encodedToken = fc.constantFrom(
      "text",
      " ",
      "\t",
      "\r",
      "\r\n",
      "žluťoučký",
      "日本語",
      "&amp;",
      "&apos;",
      "&gt;",
      "&lt;",
      "&quot;",
      "&#65;",
      "&#x1F642;",
    );

    fc.assert(
      fc.property(
        fc.array(encodedToken, { maxLength: 40 }),
        fc.array(encodedToken, { maxLength: 20 }),
        (textTokens, attributeTokens) => {
          const xml = `<document value="${attributeTokens.join("")}"><text>${textTokens.join(
            "",
          )}</text><empty/></document>`;
          expect(parseStreamingXml(xml)).toEqual({
            status: "parsed",
            value: parseXml(xml),
          });
        },
      ),
      { numRuns: 250 },
    );
  });

  test("matches the compatibility parser for every repository document body", async () => {
    const paths: string[] = [];
    for (const pattern of DOCUMENT_FIXTURE_GLOBS) {
      paths.push(...new Bun.Glob(pattern).scanSync({ cwd: REPO_ROOT }));
    }
    expect(paths.length).toBeGreaterThan(30);

    for (const path of paths) {
      const zip = await JSZip.loadAsync(readFileSync(resolve(REPO_ROOT, path)));
      const documentFile = zip.file("word/document.xml");
      if (!documentFile) {
        throw new Error(`Missing word/document.xml in ${path}`);
      }
      const xml = await documentFile.async("string");
      expect(parseStreamingXml(xml), path).toEqual({
        status: "parsed",
        value: parseXml(xml),
      });
    }
  });
});
