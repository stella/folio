import { describe, expect, test } from "bun:test";

import {
  attributeRemainder,
  InvalidPreservedAttributeError,
  NO_MODELLED_ATTRIBUTES,
  serializePreservedAttributes,
} from "./attributeRemainder";

import { parseXmlDocument } from "./xmlParser";
import { OOXML_NAMESPACES } from "./serializer/partNamespaces";

test("namespace binding projections stay outside every element's attribute remainder", () => {
  for (const name of ["p", "r", "tr", "sectPr"]) {
    for (const prefix of ["mc", "alias"]) {
      const node = parseXmlDocument(
        `<w:${name} xmlns:w="${OOXML_NAMESPACES.w.uri}" xmlns:${prefix}="${OOXML_NAMESPACES.mc.uri}" ${prefix}:Ignorable="w14" w:rsidR="00112233"/>`,
      );
      if (!node) throw new Error("Expected element");
      expect(attributeRemainder({ element: node, modelled: NO_MODELLED_ATTRIBUTES })).toEqual([
        { namespace: OOXML_NAMESPACES.w.uri, name: "rsidR", value: "00112233" },
      ]);
    }
  }
});

describe("preserved attribute serialization", () => {
  test("escapes a valid resolved attribute", () => {
    expect(serializePreservedAttributes([], [{ name: "vendorFlag", value: 'a&"b' }])).toEqual([
      'vendorFlag="a&amp;&quot;b"',
    ]);
  });

  test.each([
    { name: 'x="1"/><w:sectPr', value: "1" },
    { name: "xmlns", value: "urn:hostile" },
    { namespace: "urn:unbound", name: "flag", value: "1" },
  ])("rejects an attribute whose name cannot be emitted: %o", (attribute) => {
    expect(() => serializePreservedAttributes([], [attribute])).toThrow(
      InvalidPreservedAttributeError,
    );
  });
});
