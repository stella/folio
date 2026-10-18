import { describe, expect, test } from "bun:test";
import { setDocumentBodyComments } from "@stll/docx-core/model";
import type { BlockContent } from "@stll/docx-core/model";
import fc from "fast-check";

import type { Document } from "../types/document";
import { validateFolioDocumentModel } from "./modelValidation";

const documentWith = (body: Document["package"]["document"]): Document => ({
  package: { document: body },
});

// Deliberately malformed input: typed only at the validation boundary.
const malformedDocumentWith = (body: unknown): Document =>
  ({ package: { document: body } }) as unknown as Document;

describe("canonical Folio document model validation", () => {
  test("rejects an empty comments array", () => {
    const result = validateFolioDocumentModel(documentWith({ content: [], comments: [] }));

    expect(result.valid).toBe(false);
    expect(result.issues).toContainEqual({
      path: "package.document.comments",
      message: "Empty comment lists must be absent.",
      severity: "error",
    });
  });

  test("rejects empty paragraph formatting", () => {
    const result = validateFolioDocumentModel(
      documentWith({ content: [{ type: "paragraph", content: [], formatting: {} }] }),
    );

    expect(result.valid).toBe(false);
    expect(result.issues).toContainEqual({
      path: "package.document.content[0].formatting",
      message: "Empty paragraph formatting must be absent.",
      severity: "error",
    });
  });

  test("rejects paragraph formatting that contains only source replay metadata", () => {
    const result = validateFolioDocumentModel(
      documentWith({
        content: [{ type: "paragraph", content: [], formatting: { sourceXml: "<w:pPr/>" } }],
      }),
    );

    expect(result.valid).toBe(false);
    expect(result.issues).toContainEqual({
      path: "package.document.content[0].formatting",
      message: "Empty paragraph formatting must be absent.",
      severity: "error",
    });
  });

  test.each([
    {
      name: "header",
      build: (): Document => ({
        package: {
          document: { content: [] },
          headers: new Map([
            [
              "rId1",
              {
                type: "header",
                hdrFtrType: "default",
                content: [{ type: "paragraph", content: [], formatting: {} }],
              },
            ],
          ]),
        },
      }),
      path: 'package.headers["rId1"].content[0].formatting',
    },
    {
      name: "footnote",
      build: (): Document => ({
        package: {
          document: { content: [] },
          footnotes: [
            {
              type: "footnote",
              id: 1,
              content: [{ type: "paragraph", content: [], formatting: {} }],
            },
          ],
        },
      }),
      path: "package.footnotes[0].content[0].formatting",
    },
    {
      name: "comment",
      build: (): Document =>
        documentWith({
          content: [],
          comments: [
            {
              id: 1,
              author: "Reviewer",
              content: [{ type: "paragraph", content: [], formatting: {} }],
            },
          ],
        }),
      path: "package.document.comments[0].content[0].formatting",
    },
    {
      name: "block custom XML",
      build: (): Document =>
        documentWith({
          content: [
            {
              type: "blockCustomXml",
              openingXml: "<w:customXml>",
              closingXml: "</w:customXml>",
              content: [{ type: "paragraph", content: [], formatting: {} }],
            },
          ],
        }),
      path: "package.document.content[0].content[0].formatting",
    },
  ])("rejects empty paragraph formatting in a $name story", ({ build, path }) => {
    expect(validateFolioDocumentModel(build()).issues).toContainEqual({
      path,
      message: "Empty paragraph formatting must be absent.",
      severity: "error",
    });
  });

  test("visits every recursive block container kind", () => {
    const traversal = {
      paragraph: "leaf",
      table: "container",
      blockSdt: "container",
      blockCustomXml: "container",
      preservedBlock: "leaf",
      bookmarkStart: "leaf",
      bookmarkEnd: "leaf",
    } as const satisfies Record<BlockContent["type"], "container" | "leaf">;
    const emptyFormattingParagraph = {
      type: "paragraph",
      content: [],
      formatting: {},
    } as const;
    const blocks = {
      paragraph: { type: "paragraph", content: [] },
      table: {
        type: "table",
        rows: [{ cells: [{ content: [emptyFormattingParagraph] }] }],
      },
      blockSdt: { type: "blockSdt", properties: {}, content: [emptyFormattingParagraph] },
      blockCustomXml: {
        type: "blockCustomXml",
        openingXml: "<w:customXml>",
        closingXml: "</w:customXml>",
        content: [emptyFormattingParagraph],
      },
      preservedBlock: { type: "preservedBlock", xml: "<w:altChunk/>" },
      bookmarkStart: { type: "bookmarkStart", id: 1, name: "bookmark" },
      bookmarkEnd: { type: "bookmarkEnd", id: 1 },
    } as const satisfies Record<BlockContent["type"], BlockContent>;

    for (const block of Object.values(blocks)) {
      if (traversal[block.type] === "leaf") continue;
      expect(validateFolioDocumentModel(documentWith({ content: [block] })).valid).toBe(false);
    }
  });

  test("rejects null paragraph formatting without throwing", () => {
    const document = malformedDocumentWith({
      content: [{ type: "paragraph", content: [], formatting: null }],
    });

    expect(() => validateFolioDocumentModel(document)).not.toThrow();
    expect(validateFolioDocumentModel(document)).toMatchObject({ valid: false });
  });

  test.each([
    {
      name: "a non-object row",
      rows: [null],
      path: "package.document.content[0].rows[0]",
    },
    {
      name: "missing cells",
      rows: [{}],
      path: "package.document.content[0].rows[0].cells",
    },
    {
      name: "a null cell",
      rows: [{ cells: [null] }],
      path: "package.document.content[0].rows[0].cells[0]",
    },
  ])("rejects malformed tables with $name without throwing", ({ rows, path }) => {
    const document = malformedDocumentWith({ content: [{ type: "table", rows }] });

    expect(() => validateFolioDocumentModel(document)).not.toThrow();
    expect(validateFolioDocumentModel(document).issues).toContainEqual({
      path,
      message: expect.any(String),
      severity: "error",
    });
  });

  test("never throws when arbitrary JSON replaces a document position", () => {
    fc.assert(
      fc.property(fc.integer({ min: 0, max: 6 }), fc.jsonValue(), (position, replacement) => {
        const document = documentWith({
          content: [
            {
              type: "table",
              rows: [
                { cells: [{ content: [{ type: "paragraph", content: [], formatting: {} }] }] },
              ],
            },
          ],
        });
        const table = document.package.document.content[0];
        if (table?.type !== "table") throw new Error("Invalid test fixture");

        switch (position) {
          case 0:
            Object.assign(document, { package: replacement });
            break;
          case 1:
            Object.assign(document.package, { document: replacement });
            break;
          case 2:
            Object.assign(document.package.document, { content: replacement });
            break;
          case 3:
            Object.assign(table, { rows: replacement });
            break;
          case 4:
            Object.assign(table.rows.at(0) ?? {}, { cells: replacement });
            break;
          case 5:
            Object.assign(table.rows.at(0)?.cells.at(0) ?? {}, { content: replacement });
            break;
          case 6:
            Object.assign(table.rows.at(0)?.cells.at(0)?.content.at(0) ?? {}, {
              formatting: replacement,
            });
        }

        expect(() => validateFolioDocumentModel(document)).not.toThrow();
      }),
    );
  });

  test("the comments owner stores undefined for an empty list", () => {
    const body: Document["package"]["document"] = { content: [] };

    setDocumentBodyComments(body, []);

    expect(body.comments).toBeUndefined();
    expect(validateFolioDocumentModel(documentWith(body)).valid).toBe(true);
  });
});
