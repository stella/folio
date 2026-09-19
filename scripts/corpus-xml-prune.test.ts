import { describe, expect, test } from "bun:test";

import { prepareXmlPruning } from "./lib/corpus-xml-prune";

const DOCUMENT = [
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
  '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">',
  "<w:body>",
  "<w:p><w:r><w:t>first</w:t></w:r></w:p>",
  "<w:p><w:r><w:t>second</w:t></w:r></w:p>",
  "</w:body>",
  "</w:document>",
].join("");

describe("prepareXmlPruning", () => {
  test("addresses every element once, in pre-order", () => {
    // document, body, p, r, t, p, r, t
    expect(prepareXmlPruning(DOCUMENT).addresses).toHaveLength(8);
  });

  test("keeping everything reproduces the document", () => {
    const prunable = prepareXmlPruning(DOCUMENT);
    const rendered = prunable.render(new Set(prunable.addresses));
    expect(rendered).toContain("first");
    expect(rendered).toContain("second");
    expect(rendered).toContain(
      'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"',
    );
  });

  test("dropping an element drops its subtree and nothing else", () => {
    const prunable = prepareXmlPruning(DOCUMENT);
    const firstParagraph = 2;
    const rendered = prunable.render(
      new Set(prunable.addresses.filter((address) => address !== firstParagraph)),
    );
    expect(rendered).not.toContain("first");
    expect(rendered).toContain("second");
  });

  test("an address means the same element however many others were dropped", () => {
    const prunable = prepareXmlPruning(DOCUMENT);
    const secondParagraph = 5;
    const withoutFirstRun = prunable.render(
      new Set(prunable.addresses.filter((address) => address !== 3)),
    );
    const withoutBoth = prunable.render(
      new Set(prunable.addresses.filter((address) => address !== 3 && address !== secondParagraph)),
    );
    expect(withoutFirstRun).toContain("second");
    expect(withoutBoth).not.toContain("second");
  });

  test("the XML declaration survives pruning", () => {
    const prunable = prepareXmlPruning(DOCUMENT);
    expect(prunable.render(new Set()).startsWith("<?xml")).toBe(true);
  });
});
