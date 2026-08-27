import { describe, expect, test } from "bun:test";

import {
  cloneWithXmlnsDeclarations,
  collectXmlnsDeclarations,
  elementToXml,
  mergeXmlnsDeclarations,
  NAMESPACES,
  parseOnOffValue,
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

describe("ST_OnOff values", () => {
  test("recognizes every on and off spelling", () => {
    for (const value of ["1", "true", "on"]) {
      expect(parseOnOffValue(value)).toBe(true);
    }
    for (const value of ["0", "false", "off"]) {
      expect(parseOnOffValue(value)).toBe(false);
    }
  });

  test("leaves absent and invalid values unresolved", () => {
    expect(parseOnOffValue(null)).toBeUndefined();
    expect(parseOnOffValue(undefined)).toBeUndefined();
    expect(parseOnOffValue("yes")).toBeUndefined();
    expect(parseOnOffValue(" TRUE ")).toBeUndefined();
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

  test("drops a declaration whose value is longer than a namespace URI", () => {
    const element: XmlElement = {
      type: "element",
      name: "w:pict",
      attributes: {
        "xmlns:a": "urn:example:a",
        "xmlns:b": `urn:${"x".repeat(600)}`,
      },
    };

    expect(collectXmlnsDeclarations(element)).toEqual({ "xmlns:a": "urn:example:a" });
  });

  test("stops once the collected set reaches the chain budget", () => {
    // The merge step compares the accumulated set against the same budget, so
    // bounding the collected set here keeps every set this module hands out
    // within it — including the root set a capture inherits unchanged.
    const attributes: Record<string, string> = {};
    for (let i = 0; i < 60; i++) {
      attributes[`xmlns:ns${String(i)}`] = `urn:${"y".repeat(400)}:${String(i)}`;
    }
    const element: XmlElement = { type: "element", name: "w:pict", attributes };

    const collected = collectXmlnsDeclarations(element);
    const chars = Object.entries(collected).reduce(
      (total, [name, value]) => total + name.length + value.length,
      0,
    );

    expect(chars).toBeLessThanOrEqual(8192);
    expect(Object.keys(collected).length).toBeGreaterThan(0);
  });
});

describe("mergeXmlnsDeclarations", () => {
  test("merges own declarations over the inherited set", () => {
    const element: XmlElement = {
      type: "element",
      name: "w:p",
      attributes: { "xmlns:w": "urn:w:new", "xmlns:v": "urn:v" },
    };

    expect(mergeXmlnsDeclarations({ "xmlns:w": "urn:w:old" }, element)).toEqual({
      "xmlns:w": "urn:w:new",
      "xmlns:v": "urn:v",
    });
  });

  test("keeps the inherited set when merging would exceed the chain budget", () => {
    // Each ancestor contributes to the set a captured subtree replays, so the
    // accumulated size is bounded as well as the per-element count.
    const inherited: Record<string, string> = {};
    for (let i = 0; i < 40; i++) {
      inherited[`xmlns:in${String(i)}`] = `urn:${"y".repeat(200)}:${String(i)}`;
    }
    const attributes: Record<string, string> = {};
    for (let i = 0; i < 40; i++) {
      attributes[`xmlns:own${String(i)}`] = `urn:${"z".repeat(200)}:${String(i)}`;
    }
    const element: XmlElement = { type: "element", name: "w:p", attributes };

    expect(mergeXmlnsDeclarations(inherited, element)).toBe(inherited);
  });
});
