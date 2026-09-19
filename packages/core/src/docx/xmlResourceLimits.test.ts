import { describe, expect, test } from "bun:test";

import {
  assertXmlResourceLimits,
  createXmlPackageBudget,
  FOLIO_XML_RESOURCE_LIMITS,
  XmlResourceLimitError,
} from "./xmlResourceLimits";

const limits = {
  ...FOLIO_XML_RESOURCE_LIMITS,
  maxBytes: 1024,
  maxDepth: 3,
  maxElementsPerPart: 4,
  maxAttributesPerPart: 4,
  maxElementsPerPackage: 6,
  maxAttributesPerPackage: 6,
};

describe("assertXmlResourceLimits", () => {
  test("counts UTF-8 bytes rather than JavaScript code units", () => {
    expect(() => assertXmlResourceLimits({ xml: "é", limits: { ...limits, maxBytes: 1 } })).toThrow(
      XmlResourceLimitError,
    );
  });

  test("rejects excessive element count before parsing", () => {
    expect(() => assertXmlResourceLimits({ xml: "<root><a/><b/><c/><d/></root>", limits })).toThrow(
      XmlResourceLimitError,
    );
  });

  test("rejects excessive nesting before recursive tree conversion", () => {
    expect(() => assertXmlResourceLimits({ xml: "<a><b><c><d/></c></b></a>", limits })).toThrow(
      XmlResourceLimitError,
    );
  });

  test("ignores markup-like content in comments, CDATA, and quoted attributes", () => {
    expect(() =>
      assertXmlResourceLimits({
        xml: '<root value=">"><!-- <fake/> --><![CDATA[<fake/>]]></root>',
        limits,
      }),
    ).not.toThrow();
  });

  test("counts attributes, which no element bound catches", () => {
    // One element, many attributes: an element budget alone leaves this
    // unbounded, and each attribute retains roughly 18 bytes of tree.
    expect(() =>
      assertXmlResourceLimits({ xml: `<root a="1" b="2" c="3" d="4" e="5"/>`, limits }),
    ).toThrow(XmlResourceLimitError);
  });

  test("does not count an `=` inside a quoted attribute value", () => {
    expect(() => assertXmlResourceLimits({ xml: `<root a="x=y=z=w=v=u"/>`, limits })).not.toThrow();
  });

  test("reports the part path, the count reached, and the bound", () => {
    try {
      assertXmlResourceLimits({
        xml: "<root><a/><b/><c/><d/></root>",
        limits,
        partPath: "word/document.xml",
      });
      expect.unreachable();
    } catch (error) {
      expect(error).toBeInstanceOf(XmlResourceLimitError);
      const refusal = error as XmlResourceLimitError;
      expect(refusal.limit).toBe("elements");
      expect(refusal.partPath).toBe("word/document.xml");
      expect(refusal.observed).toBe(5);
      expect(refusal.allowed).toBe(4);
      expect(refusal.message).toContain("word/document.xml");
    }
  });

  test("bounds a package of parts that are each within the per-part bound", () => {
    const budget = createXmlPackageBudget();
    const part = "<root><a/><b/></root>";
    // Three elements a part, six allowed a package: the third part is refused
    // although no part on its own comes near the per-part bound of four.
    expect(() => assertXmlResourceLimits({ xml: part, limits, budget })).not.toThrow();
    expect(() => assertXmlResourceLimits({ xml: part, limits, budget })).not.toThrow();
    expect(budget.elements).toBe(6);
    try {
      assertXmlResourceLimits({ xml: part, limits, budget, partPath: "word/header3.xml" });
      expect.unreachable();
    } catch (error) {
      const refusal = error as XmlResourceLimitError;
      expect(refusal.limit).toBe("package-elements");
      expect(refusal.observed).toBe(7);
      expect(refusal.allowed).toBe(6);
    }
  });

  test("charges the budget only for input it accepted", () => {
    const budget = createXmlPackageBudget();
    expect(() =>
      assertXmlResourceLimits({ xml: "<root><a/><b/><c/><d/></root>", limits, budget }),
    ).toThrow(XmlResourceLimitError);
    expect(budget.elements).toBe(0);
    expect(budget.attributes).toBe(0);
  });

  test("returns what it counted", () => {
    expect(assertXmlResourceLimits({ xml: `<root x="1"><a y="2"/></root>`, limits })).toEqual({
      elements: 2,
      attributes: 2,
      maxDepth: 2,
    });
  });
});
