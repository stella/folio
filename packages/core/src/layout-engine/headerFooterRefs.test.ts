import { describe, expect, test } from "bun:test";

import { createEmptyDocument } from "../utils/createDocument";
import { resolveSectionHeaderFooterRefs } from "./headerFooterRefs";

describe("resolveSectionHeaderFooterRefs", () => {
  test("inherits omitted slots independently across sections", () => {
    const document = createEmptyDocument();
    document.package.document.sections = [
      {
        content: [],
        properties: {
          headerReferences: [
            { type: "default", rId: "header-default" },
            { type: "first", rId: "header-first" },
          ],
          footerReferences: [{ type: "default", rId: "footer-default" }],
        },
      },
      {
        content: [],
        properties: {
          titlePg: true,
          headerReferences: [{ type: "first", rId: "next-header-first" }],
        },
      },
      {
        content: [],
        properties: {
          footerReferences: [{ type: "default", rId: "next-footer-default" }],
        },
      },
    ];

    expect(resolveSectionHeaderFooterRefs(document)).toEqual([
      {
        evenAndOddHeaders: false,
        headerDefault: "header-default",
        headerFirst: "header-first",
        footerDefault: "footer-default",
      },
      {
        evenAndOddHeaders: false,
        titlePg: true,
        headerDefault: "header-default",
        headerFirst: "next-header-first",
        footerDefault: "footer-default",
      },
      {
        evenAndOddHeaders: false,
        headerDefault: "header-default",
        headerFirst: "next-header-first",
        footerDefault: "next-footer-default",
      },
    ]);
  });

  test("applies the document odd-even setting without inheriting title-page mode", () => {
    const document = createEmptyDocument();
    document.package.settings = { evenAndOddHeaders: true };
    document.package.document.sections = [
      {
        content: [],
        properties: {
          titlePg: true,
          headerReferences: [{ type: "even", rId: "even-header" }],
        },
      },
      { content: [], properties: {} },
    ];

    expect(resolveSectionHeaderFooterRefs(document)).toEqual([
      { evenAndOddHeaders: true, titlePg: true, headerEven: "even-header" },
      { evenAndOddHeaders: true, headerEven: "even-header" },
    ]);
  });

  test("uses final section properties when the section list is absent", () => {
    const document = createEmptyDocument();
    document.package.document.finalSectionProperties = {
      footerReferences: [{ type: "default", rId: "footer" }],
    };

    expect(resolveSectionHeaderFooterRefs(document)).toEqual([
      { evenAndOddHeaders: false, footerDefault: "footer" },
    ]);
  });

  test("returns no references without a document model", () => {
    expect(resolveSectionHeaderFooterRefs(null)).toBeUndefined();
    expect(resolveSectionHeaderFooterRefs(undefined)).toBeUndefined();
  });
});
