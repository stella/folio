import { describe, expect, test } from "bun:test";

import type {
  Document,
  DocumentBody,
  HeaderFooter,
  Paragraph,
  SectionProperties,
} from "../types/document";
import {
  createEmptyHeaderFooter,
  removeHeaderFooter,
  resolveEffectiveSectionProperties,
  resolveHeaderFooterContent,
  saveHeaderFooterContent,
} from "./headerFooter";

const paragraph: Paragraph = {
  type: "paragraph",
  content: [],
};

const docFromPackage = (
  body: DocumentBody,
  headers?: Map<string, HeaderFooter>,
  footers?: Map<string, HeaderFooter>,
): Document => ({
  package: {
    document: body,
    ...(headers ? { headers } : {}),
    ...(footers ? { footers } : {}),
  },
});

describe("resolveEffectiveSectionProperties", () => {
  test("uses the first content section instead of an empty final section", () => {
    const firstSection: SectionProperties = {
      marginTop: 1296,
      headerDistance: 288,
      titlePg: true,
    };
    const emptyFinalSection: SectionProperties = {
      marginTop: 1728,
      headerDistance: 1872,
      titlePg: true,
    };
    const body: DocumentBody = {
      content: [paragraph],
      sections: [
        { properties: firstSection, content: [paragraph] },
        { properties: emptyFinalSection, content: [] },
      ],
      finalSectionProperties: emptyFinalSection,
    };

    expect(resolveEffectiveSectionProperties(body, true)).toBe(firstSection);
  });

  test("preserves title-page mode when inherited from header references", () => {
    const firstSection: SectionProperties = { headerDistance: 288 };
    const body: DocumentBody = {
      content: [paragraph],
      sections: [{ properties: firstSection, content: [paragraph] }],
    };

    expect(resolveEffectiveSectionProperties(body, true)).toEqual({
      headerDistance: 288,
      titlePg: true,
    });
  });
});

describe("resolveHeaderFooterContent", () => {
  test("returns the empty resolution for a missing package", () => {
    expect(resolveHeaderFooterContent(undefined)).toMatchObject({
      headerContent: null,
      footerContent: null,
      hasTitlePg: false,
      activeHeaderRId: null,
    });
  });

  test("resolves the default header from finalSectionProperties", () => {
    const header: HeaderFooter = { type: "header", content: [paragraph] };
    const body: DocumentBody = {
      content: [paragraph],
      finalSectionProperties: {
        headerReferences: [{ type: "default", rId: "rIdH" }],
      },
    };
    const resolution = resolveHeaderFooterContent(
      docFromPackage(body, new Map([["rIdH", header]])).package,
    );

    expect(resolution.headerContent).toBe(header);
    expect(resolution.activeHeaderRId).toBe("rIdH");
    expect(resolution.hasTitlePg).toBe(false);
  });

  test("resolves the first-page footer when a section enables titlePg", () => {
    const firstFooter: HeaderFooter = { type: "footer", hdrFtrType: "first", content: [paragraph] };
    const body: DocumentBody = {
      content: [paragraph],
      sections: [
        {
          properties: {
            titlePg: true,
            footerReferences: [{ type: "first", rId: "rIdF1" }],
          },
          content: [paragraph],
        },
      ],
    };
    const resolution = resolveHeaderFooterContent(
      docFromPackage(body, undefined, new Map([["rIdF1", firstFooter]])).package,
    );

    expect(resolution.hasTitlePg).toBe(true);
    expect(resolution.firstPageFooterContent).toBe(firstFooter);
    expect(resolution.activeFirstFooterRId).toBe("rIdF1");
  });
});

describe("createEmptyHeaderFooter", () => {
  test("returns null when there are no final section properties", () => {
    const body: DocumentBody = { content: [paragraph] };
    expect(createEmptyHeaderFooter(docFromPackage(body), "header", false)).toBeNull();
  });

  test("adds an empty header plus a reference into finalSectionProperties", () => {
    const body: DocumentBody = {
      content: [paragraph],
      finalSectionProperties: { marginTop: 1440 },
    };
    const next = createEmptyHeaderFooter(docFromPackage(body), "header", false);

    expect(next).not.toBeNull();
    const rId = "rId_new_header_default";
    expect(next?.package.headers?.get(rId)).toEqual({
      type: "header",
      hdrFtrType: "default",
      content: [{ type: "paragraph", content: [] }],
    });
    expect(next?.package.document.finalSectionProperties?.headerReferences).toEqual([
      { type: "default", rId },
    ]);
    expect(next?.package.relationships?.get(rId)).toEqual({
      id: rId,
      type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header",
      target: "header1.xml",
    });
  });

  test("allocates collision-free relationship ids and part targets", () => {
    const body: DocumentBody = {
      content: [paragraph],
      finalSectionProperties: { marginTop: 1440 },
    };
    const document = docFromPackage(
      body,
      new Map([
        ["rId_new_header_default", { type: "header", hdrFtrType: "default", content: [paragraph] }],
      ]),
    );
    document.package.relationships = new Map([
      [
        "rIdExisting",
        {
          id: "rIdExisting",
          type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header",
          target: "header1.xml",
        },
      ],
    ]);

    const next = createEmptyHeaderFooter(document, "header", false);

    expect(next?.package.headers?.has("rId_new_header_default_2")).toBe(true);
    expect(next?.package.relationships?.get("rId_new_header_default_2")?.target).toBe(
      "header2.xml",
    );
  });
});

describe("saveHeaderFooterContent", () => {
  test("returns null when the target relationship id is missing", () => {
    const body: DocumentBody = { content: [paragraph] };
    const result = saveHeaderFooterContent({
      document: docFromPackage(body, new Map()),
      position: "header",
      isFirstPage: false,
      activeRId: "missing",
      blocks: [paragraph],
    });
    expect(result).toBeNull();
  });

  test("writes new blocks without mutating the prior document", () => {
    const existing: HeaderFooter = { type: "header", content: [] };
    const headers = new Map([["rIdH", existing]]);
    const body: DocumentBody = { content: [paragraph] };
    const original = docFromPackage(body, headers);

    const next = saveHeaderFooterContent({
      document: original,
      position: "header",
      isFirstPage: false,
      activeRId: "rIdH",
      blocks: [paragraph],
    });

    expect(next?.package.headers?.get("rIdH")?.content).toEqual([paragraph]);
    // Prior document's map + entry stay untouched so undo can restore it.
    expect(original.package.headers).toBe(headers);
    expect(existing.content).toEqual([]);
  });
});

describe("removeHeaderFooter", () => {
  test("drops the header and its references from every section", () => {
    const header: HeaderFooter = { type: "header", content: [paragraph] };
    const body: DocumentBody = {
      content: [paragraph],
      finalSectionProperties: {
        headerReferences: [{ type: "default", rId: "rIdH" }],
      },
    };
    const next = removeHeaderFooter({
      document: docFromPackage(body, new Map([["rIdH", header]])),
      position: "header",
      activeRId: "rIdH",
    });

    expect(next.package.headers?.has("rIdH")).toBe(false);
    expect(next.package.document.finalSectionProperties?.headerReferences).toEqual([]);
  });
});
