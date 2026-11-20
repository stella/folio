import { expect, spyOn, test } from "bun:test";

import * as reader from "@stll/docx-core/xml";
import { FOLIO_XML_RESOURCE_LIMITS } from "@stll/docx-core/xml/limits";

import { isSafeCapturedXmlDocument } from "./verbatimCapture";
import * as parser from "./xmlParser";

test("standalone, entity-free paragraph fragments use the syntax tree for namespace validation", () => {
  const syntax = spyOn(reader, "wellFormedXmlIssue");
  const tree = spyOn(reader, "readXml");
  const envelope = spyOn(parser, "parseXml");
  try {
    for (let i = 0; i < 32; i += 1) {
      const fragment = `<w:p xmlns:w="urn:p"><w:r><w:t>${i}</w:t></w:r></w:p>`;
      expect(isSafeCapturedXmlDocument(fragment)).toBe(true);
    }
    expect(syntax).toHaveBeenCalledTimes(0);
    expect(envelope).toHaveBeenCalledTimes(0);
    expect(tree).toHaveBeenCalledTimes(32);

    for (const fragment of ["<p>A &amp; B</p>", '<?xml version="1.0"?><p/>']) {
      expect(isSafeCapturedXmlDocument(fragment)).toBe(true);
    }
    expect(syntax).toHaveBeenCalledTimes(2);
    expect(envelope).toHaveBeenCalledTimes(2);
  } finally {
    syntax.mockRestore();
    tree.mockRestore();
    envelope.mockRestore();
  }
});

test("standalone fragments keep the replay envelope's one level of nesting", () => {
  const nested = (depth: number) => "<p>".repeat(depth) + "</p>".repeat(depth);
  const { maxDepth } = FOLIO_XML_RESOURCE_LIMITS;
  expect(isSafeCapturedXmlDocument(nested(maxDepth - 1))).toBe(true);
  expect(isSafeCapturedXmlDocument(nested(maxDepth))).toBe(false);
});
