import { describe, expect, test } from "bun:test";

import { collectXmlnsDeclarations, elementToXml, parseXmlDocument } from "./xmlParser";
import type { XmlElement } from "./xmlParser";

describe("OOXML parsing", () => {
  test("preserves ordered nested elements and attributes", () => {
    const document = parseXmlDocument(
      '<w:p xmlns:w="urn" w:rsidR="1234"><w:r w:rsidRPr="5678"><w:t xml:space="preserve"> first </w:t></w:r><w:r><w:t>second</w:t></w:r></w:p>',
    );

    expect(document).toEqual({
      type: "element",
      name: "w:p",
      attributes: { "xmlns:w": "urn", "w:rsidR": "1234" },
      elements: [
        {
          type: "element",
          name: "w:r",
          attributes: { "w:rsidRPr": "5678" },
          elements: [
            {
              type: "element",
              name: "w:t",
              attributes: { "xml:space": "preserve" },
              elements: [{ type: "text", text: " first " }],
            },
          ],
        },
        {
          type: "element",
          name: "w:r",
          elements: [
            {
              type: "element",
              name: "w:t",
              elements: [{ type: "text", text: "second" }],
            },
          ],
        },
      ],
    });
  });

  test("preserves legacy inline binary payloads", () => {
    const document = parseXmlDocument(
      '<w:p xmlns:w="urn"><w:binData w:name="image1">QUJDRA==</w:binData><w:r><w:t>after</w:t></w:r></w:p>',
    );

    expect(document?.elements?.at(0)?.name).toBe("w:binData");
    expect(document?.elements?.at(0)?.elements?.at(0)?.text).toBe("QUJDRA==");
    expect(document?.elements?.at(1)?.name).toBe("w:r");
    expect(document ? elementToXml(document) : "").toContain(
      '<w:binData w:name="image1">QUJDRA==</w:binData>',
    );
  });
});

describe("collectXmlnsDeclarations", () => {
  test("caps a hostile number of distinct xmlns declarations on one element", () => {
    const attributes: Record<string, string> = {};
    for (let i = 0; i < 10_000; i++) {
      attributes[`xmlns:ns${i}`] = `urn:example:${i}`;
    }
    const element: XmlElement = { type: "element", name: "w:pict", attributes };

    expect(Object.keys(collectXmlnsDeclarations(element))).toHaveLength(64);
  });
});
