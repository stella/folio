// `w14:textId` is a Word 2010 extension, and a prefix is only an alias: the
// URI it is bound to decides whose attribute it is. Reading it by prefix fell
// through to an any-prefix local-name match, so a foreign `vendor:textId`
// entered the model as Word's paragraph identity and was written back as one.

import { describe, expect, test } from "bun:test";

import { parseParagraph } from "./paragraphParser";
import type { XmlElement } from "./xmlParser";
import { parseXmlDocument } from "./xmlParser";

const W_URI = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W14_URI = "http://schemas.microsoft.com/office/word/2010/wordml";

const parseParagraphXml = (attributes: string) => {
  const root = parseXmlDocument(
    `<w:p xmlns:w="${W_URI}" ${attributes}><w:r><w:t>x</w:t></w:r></w:p>`,
  );
  if (!root) {
    throw new Error("Failed to parse paragraph XML fixture");
  }
  return parseParagraph(root as XmlElement, null, null, null, null, null);
};

describe("paragraph identity attributes resolve by namespace URI", () => {
  test("w14:textId under the Word 2010 URI is read", () => {
    expect(parseParagraphXml(`xmlns:w14="${W14_URI}" w14:textId="1A2B3C4D"`).textId).toBe(
      "1A2B3C4D",
    );
  });

  test("an alternate prefix bound to the same URI is read", () => {
    expect(parseParagraphXml(`xmlns:ns0="${W14_URI}" ns0:textId="1A2B3C4D"`).textId).toBe(
      "1A2B3C4D",
    );
  });

  test("a textId under a foreign URI is not Word's", () => {
    expect(
      parseParagraphXml('xmlns:vendor="urn:example:vendor" vendor:textId="1A2B3C4D"').textId,
    ).toBeUndefined();
  });
});
