import { describe, expect, test } from "bun:test";
import type { Node as PMNode } from "prosemirror-model";

import type { Document, Paragraph } from "../../types/document";
import {
  assignParagraphPropertySource,
  getParagraphPropertySource,
} from "../../docx/paragraphPropertySource";
import { canonicalJson } from "../../utils/canonicalJson";
import { fromProseDoc } from "./fromProseDoc";
import { toProseDoc } from "./toProseDoc";

function childTypeNames(pmDoc: ReturnType<typeof toProseDoc>): string[] {
  const names: string[] = [];
  for (let i = 0; i < pmDoc.childCount; i++) {
    names.push(pmDoc.child(i).type.name);
  }
  return names;
}

function descendantsOfType(pmDoc: ReturnType<typeof toProseDoc>, typeName: string) {
  const nodes: PMNode[] = [];
  pmDoc.descendants((node) => {
    if (node.type.name === typeName) {
      nodes.push(node);
    }
  });
  return nodes;
}

describe('toProseDoc — hard page break (`<w:br w:type="page"/>`)', () => {
  test.each([
    ["before", [{ type: "break", breakType: "page" }] as const],
    [
      "after",
      [
        { type: "text", text: "Before" },
        { type: "break", breakType: "page" },
      ] as const,
    ],
  ])("a %s carrier rebuild preserves its paragraph-property owner", (_position, content) => {
    const paragraph: Paragraph = {
      type: "paragraph",
      content: [{ type: "run", content: [...content] }],
    };
    assignParagraphPropertySource(paragraph, {
      formattingJson: canonicalJson({}),
      xml:
        '<w:pPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
        'xmlns:x="urn:folio:test"><x:property/></w:pPr>',
    });
    const document: Document = { package: { document: { content: [paragraph] } } };

    const restored = fromProseDoc(toProseDoc(document), document);
    const restoredParagraph = restored.package.document.content.find(
      (block): block is Paragraph => block.type === "paragraph",
    );
    expect(getParagraphPropertySource(restoredParagraph)).toEqual(
      getParagraphPropertySource(paragraph),
    );
  });

  test("keeps a break-only run inside its paragraph", () => {
    // <w:p><w:r><w:br w:type="page"/></w:r></w:p>
    const document: Document = {
      package: {
        document: {
          content: [
            {
              type: "paragraph",
              content: [
                {
                  type: "run",
                  content: [{ type: "text", text: "Before" }],
                },
              ],
            },
            {
              type: "paragraph",
              content: [
                {
                  type: "run",
                  content: [{ type: "break", breakType: "page" }],
                },
              ],
            },
            {
              type: "paragraph",
              content: [
                {
                  type: "run",
                  content: [{ type: "text", text: "After" }],
                },
              ],
            },
          ],
        },
      },
    };

    const pmDoc = toProseDoc(document);
    expect(childTypeNames(pmDoc)).toEqual(["paragraph", "paragraph", "paragraph"]);
    expect(pmDoc.child(1).child(0).type.name).toBe("pageBreakRun");
  });

  test("keeps an enabled split break-only paragraph as authored page content", () => {
    const document: Document = {
      package: {
        document: {
          content: [
            {
              type: "paragraph",
              content: [
                {
                  type: "run",
                  content: [{ type: "break", breakType: "page" }],
                },
              ],
            },
            { type: "paragraph", content: [] },
          ],
        },
        settings: {
          defaultTabStop: 720,
          splitPageBreakAndParagraphMark: true,
        },
      },
    };

    const pmDoc = toProseDoc(document);

    expect(childTypeNames(pmDoc)).toEqual(["paragraph", "paragraph"]);
    expect(pmDoc.child(0).child(0).type.name).toBe("pageBreakRun");
    expect(pmDoc.child(0).attrs["_pageBreakCarrier"]).toBeNull();
    expect(pmDoc.child(1).attrs["_pageBreakCarrier"]).toBeNull();
  });

  test("does not tag a separate authored empty paragraph as a legacy break carrier", () => {
    const document: Document = {
      package: {
        document: {
          content: [
            {
              type: "paragraph",
              content: [
                {
                  type: "run",
                  content: [{ type: "break", breakType: "page" }],
                },
              ],
            },
            { type: "paragraph", content: [] },
          ],
        },
      },
    };

    const pmDoc = toProseDoc(document);

    expect(pmDoc.child(0).child(0).type.name).toBe("pageBreakRun");
    expect(pmDoc.child(0).attrs["_pageBreakCarrier"]).toBeNull();
    expect(pmDoc.child(1).attrs["_pageBreakCarrier"]).toBeNull();
  });

  test("keeps a trailing page break after text in the same paragraph", () => {
    const document: Document = {
      package: {
        document: {
          content: [
            {
              type: "paragraph",
              content: [
                {
                  type: "run",
                  content: [
                    { type: "text", text: "Before" },
                    { type: "break", breakType: "page" },
                  ],
                },
              ],
            },
            {
              type: "paragraph",
              content: [
                {
                  type: "run",
                  content: [{ type: "text", text: "After" }],
                },
              ],
            },
          ],
        },
      },
    };

    const pmDoc = toProseDoc(document);
    expect(childTypeNames(pmDoc)).toEqual(["paragraph", "paragraph"]);
    expect(pmDoc.child(0).child(0).text).toBe("Before");
    expect(pmDoc.child(0).child(1).type.name).toBe("pageBreakRun");
    expect(pmDoc.child(0).attrs["_trailingPageBreak"]).toBeNull();
  });

  test("keeps the inline carrier inside a hyperlink wrapper", () => {
    // <w:p><w:hyperlink><w:r><w:br w:type="page"/></w:r></w:hyperlink></w:p>
    const document: Document = {
      package: {
        document: {
          content: [
            {
              type: "paragraph",
              content: [
                {
                  type: "run",
                  content: [{ type: "text", text: "Before" }],
                },
              ],
            },
            {
              type: "paragraph",
              content: [
                {
                  type: "hyperlink",
                  url: "https://example.com",
                  children: [
                    {
                      type: "run",
                      content: [{ type: "break", breakType: "page" }],
                    },
                  ],
                },
              ],
            },
            {
              type: "paragraph",
              content: [
                {
                  type: "run",
                  content: [{ type: "text", text: "After" }],
                },
              ],
            },
          ],
        },
      },
    };

    const pmDoc = toProseDoc(document);
    const pageBreak = descendantsOfType(pmDoc, "pageBreakRun").at(0);
    expect(pageBreak).toBeDefined();
    expect(pageBreak?.marks.some((mark) => mark.type.name === "hyperlink")).toBe(true);
  });

  test("keeps the inline carrier inside an inlineSdt wrapper", () => {
    // <w:p><w:sdt><w:sdtContent><w:r><w:br w:type="page"/></w:r></w:sdtContent></w:sdt></w:p>
    const document: Document = {
      package: {
        document: {
          content: [
            {
              type: "paragraph",
              content: [
                {
                  type: "run",
                  content: [{ type: "text", text: "Before" }],
                },
              ],
            },
            {
              type: "paragraph",
              content: [
                {
                  type: "inlineSdt",
                  properties: { sdtType: "richText" },
                  content: [
                    {
                      type: "run",
                      content: [{ type: "break", breakType: "page" }],
                    },
                  ],
                },
              ],
            },
            {
              type: "paragraph",
              content: [
                {
                  type: "run",
                  content: [{ type: "text", text: "After" }],
                },
              ],
            },
          ],
        },
      },
    };

    const pmDoc = toProseDoc(document);
    expect(descendantsOfType(pmDoc, "pageBreakRun")).toHaveLength(1);
    expect(descendantsOfType(pmDoc, "sdt")).toHaveLength(1);
  });

  test("fails closed when a page break shares its source run with a softHyphen", () => {
    // <w:p><w:r><w:softHyphen/><w:br w:type="page"/></w:r></w:p>
    const document: Document = {
      package: {
        document: {
          content: [
            {
              type: "paragraph",
              content: [
                {
                  type: "run",
                  content: [{ type: "softHyphen" }, { type: "break", breakType: "page" }],
                },
              ],
            },
            {
              type: "paragraph",
              content: [
                {
                  type: "run",
                  content: [{ type: "text", text: "After" }],
                },
              ],
            },
          ],
        },
      },
    };

    expect(() => toProseDoc(document)).toThrow(
      "A page-break-bearing run containing softHyphen cannot be represented in the editor model",
    );
  });

  test("keeps a break after a mathEquation in source order", () => {
    const document: Document = {
      package: {
        document: {
          content: [
            {
              type: "paragraph",
              content: [
                {
                  type: "mathEquation",
                  display: "inline",
                  ommlXml: "<m:oMath/>",
                },
                {
                  type: "run",
                  content: [{ type: "break", breakType: "page" }],
                },
              ],
            },
            {
              type: "paragraph",
              content: [
                {
                  type: "run",
                  content: [{ type: "text", text: "After" }],
                },
              ],
            },
          ],
        },
      },
    };

    const pmDoc = toProseDoc(document);
    expect(childTypeNames(pmDoc)).toEqual(["paragraph", "paragraph"]);
    expect(descendantsOfType(pmDoc, "pageBreakRun")).toHaveLength(1);
  });

  test("keeps a break after an empty hyperlink before following text", () => {
    // <w:p><w:hyperlink/><w:r><w:br w:type="page"/></w:r><w:r><w:t>x</w:t></w:r></w:p>
    // Empty hyperlinks carry no inline node; the break still retains its exact
    // position before the following text.
    const document: Document = {
      package: {
        document: {
          content: [
            {
              type: "paragraph",
              content: [
                {
                  type: "run",
                  content: [{ type: "text", text: "Before" }],
                },
              ],
            },
            {
              type: "paragraph",
              content: [
                {
                  type: "hyperlink",
                  url: "https://example.com",
                  children: [],
                },
                {
                  type: "run",
                  content: [{ type: "break", breakType: "page" }],
                },
                {
                  type: "run",
                  content: [{ type: "text", text: "AfterBreak" }],
                },
              ],
            },
          ],
        },
      },
    };

    const pmDoc = toProseDoc(document);
    expect(childTypeNames(pmDoc)).toEqual(["paragraph", "paragraph"]);
    const paragraph = pmDoc.child(1);
    expect(paragraph.child(0).type.name).toBe("pageBreakRun");
    expect(paragraph.child(1).text).toBe("AfterBreak");
  });

  test("keeps the inline carrier inside a tracked-change wrapper", () => {
    // <w:p><w:ins><w:r><w:br w:type="page"/></w:r></w:ins></w:p>
    const document: Document = {
      package: {
        document: {
          content: [
            {
              type: "paragraph",
              content: [
                {
                  type: "run",
                  content: [{ type: "text", text: "Before" }],
                },
              ],
            },
            {
              type: "paragraph",
              content: [
                {
                  type: "insertion",
                  info: {
                    id: 1,
                    author: "Author",
                    date: "2026-01-01T00:00:00Z",
                  },
                  content: [
                    {
                      type: "run",
                      content: [{ type: "break", breakType: "page" }],
                    },
                  ],
                },
              ],
            },
            {
              type: "paragraph",
              content: [
                {
                  type: "run",
                  content: [{ type: "text", text: "After" }],
                },
              ],
            },
          ],
        },
      },
    };

    const pmDoc = toProseDoc(document);
    const pageBreak = descendantsOfType(pmDoc, "pageBreakRun").at(0);
    expect(pageBreak).toBeDefined();
    expect(pageBreak?.marks.some((mark) => mark.type.name === "insertion")).toBe(true);
  });
});
