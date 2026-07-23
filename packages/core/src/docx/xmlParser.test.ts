import { describe, expect, test } from "bun:test";

import {
  cloneWithXmlnsDeclarations,
  collectXmlnsDeclarations,
  elementToXml,
  NAMESPACES,
  parseXmlDocument,
} from "./xmlParser";
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

  test("keeps character-referenced carriage returns stable when serialized", () => {
    const document = parseXmlDocument("<v:shape>&#xD;&#xA;<v:path/></v:shape>");
    const serialized = document ? elementToXml(document) : "";
    const reopened = parseXmlDocument(serialized);

    expect(serialized).not.toContain("\r");
    expect(reopened ? elementToXml(reopened) : "").toBe(serialized);
  });

  test("supplies canonical bindings for known unbound OOXML prefixes", () => {
    const document = parseXmlDocument(
      '<w:pict><v:shape><v:imagedata r:id="image"/></v:shape></w:pict>',
    );
    const selfContained = document ? cloneWithXmlnsDeclarations(document, {}) : null;
    const serialized = selfContained ? elementToXml(selfContained) : "";

    expect(serialized).toContain(`xmlns:w="${NAMESPACES.w}"`);
    expect(serialized).toContain(`xmlns:v="${NAMESPACES.v}"`);
    expect(serialized).toContain(`xmlns:r="${NAMESPACES.r}"`);
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
