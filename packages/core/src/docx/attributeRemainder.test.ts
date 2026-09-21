import { describe, expect, test } from "bun:test";

import { InvalidPreservedAttributeError, serializePreservedAttributes } from "./attributeRemainder";

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
