import { describe, expect, test } from "bun:test";

import { assertXmlResourceLimits, XmlResourceLimitError } from "./xmlResourceLimits";

const limits = {
  maxBytes: 1024,
  maxDepth: 3,
  maxNodes: 4,
};

describe("assertXmlResourceLimits", () => {
  test("counts UTF-8 bytes rather than JavaScript code units", () => {
    expect(() =>
      assertXmlResourceLimits("é", {
        ...limits,
        maxBytes: 1,
      }),
    ).toThrow(XmlResourceLimitError);
  });

  test("rejects excessive element count before parsing", () => {
    expect(() =>
      assertXmlResourceLimits("<root><a/><b/><c/><d/></root>", {
        ...limits,
        maxNodes: 4,
      }),
    ).toThrow(XmlResourceLimitError);
  });

  test("rejects excessive nesting before recursive tree conversion", () => {
    expect(() =>
      assertXmlResourceLimits("<a><b><c><d/></c></b></a>", {
        ...limits,
        maxDepth: 3,
      }),
    ).toThrow(XmlResourceLimitError);
  });

  test("ignores markup-like content in comments, CDATA, and quoted attributes", () => {
    expect(() =>
      assertXmlResourceLimits('<root value=">"><!-- <fake/> --><![CDATA[<fake/>]]></root>', limits),
    ).not.toThrow();
  });
});
