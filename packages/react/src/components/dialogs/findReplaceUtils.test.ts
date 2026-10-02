import { GlobalRegistrator } from "@happy-dom/global-registrator";

GlobalRegistrator.register();

import { afterAll, describe, expect, spyOn, test } from "bun:test";

import type { Document } from "@stll/folio-core/types/document";
import { replaceTextInDocument } from "@stll/folio-core/utils/replaceText";
import { createDefaultFindOptions, findInDocument, scrollToMatch } from "./findReplaceUtils";

afterAll(() => GlobalRegistrator.unregister());

const createTableDocument = (): Document => ({
  package: {
    document: {
      content: [
        {
          type: "paragraph",
          content: [
            {
              type: "run",
              content: [{ type: "text", text: "Outside text" }],
            },
          ],
        },
        {
          type: "table",
          rows: [
            {
              cells: [
                {
                  content: [
                    {
                      type: "paragraph",
                      content: [
                        {
                          type: "run",
                          content: [{ type: "text", text: "Inside table" }],
                        },
                      ],
                    },
                  ],
                },
              ],
            },
          ],
        },
      ],
    },
  },
});

describe("Folio find and replace", () => {
  test("finds text inside table cells", () => {
    const matches = findInDocument(createTableDocument(), "Inside", createDefaultFindOptions());

    expect(matches).toHaveLength(1);
    expect(matches[0]?.paragraphIndex).toBe(1);
    expect(matches[0]?.startOffset).toBe(0);
  });

  test("reports match offsets relative to the whole paragraph", () => {
    const document: Document = {
      package: {
        document: {
          content: [
            {
              type: "paragraph",
              content: [
                {
                  type: "run",
                  content: [{ type: "text", text: "Series " }],
                },
                {
                  type: "run",
                  content: [{ type: "text", text: "Stock" }],
                },
              ],
            },
          ],
        },
      },
    };

    const matches = findInDocument(document, "stock", createDefaultFindOptions());

    expect(matches).toHaveLength(1);
    expect(matches[0]?.contentIndex).toBe(1);
    expect(matches[0]?.startOffset).toBe(7);
    expect(matches[0]?.endOffset).toBe(12);
  });

  test("replaces matches inside table cells", () => {
    const document = createTableDocument();
    const match = findInDocument(document, "Inside", createDefaultFindOptions())[0];
    if (!match) {
      throw new Error("Expected table-cell match");
    }

    const replaced = replaceTextInDocument(
      document,
      {
        start: {
          paragraphIndex: match.paragraphIndex,
          offset: match.startOffset,
        },
        end: {
          paragraphIndex: match.paragraphIndex,
          offset: match.endOffset,
        },
      },
      "Within",
    );

    expect(findInDocument(replaced, "Within table", createDefaultFindOptions())).toHaveLength(1);
  });

  test("finds run text wrapped inside an inline SDT", () => {
    const document: Document = {
      package: {
        document: {
          content: [
            {
              type: "paragraph",
              content: [
                {
                  type: "run",
                  content: [{ type: "text", text: "Before " }],
                },
                {
                  type: "inlineSdt",
                  properties: {},
                  content: [
                    {
                      type: "run",
                      content: [{ type: "text", text: "tagged" }],
                    },
                  ],
                },
                {
                  type: "run",
                  content: [{ type: "text", text: " after" }],
                },
              ],
            },
          ],
        },
      },
    };

    const matches = findInDocument(document, "tagged", createDefaultFindOptions());

    expect(matches).toHaveLength(1);
    expect(matches[0]?.startOffset).toBe(7);
    expect(matches[0]?.endOffset).toBe(13);
    // The inline SDT is the top-level container at index 1.
    expect(matches[0]?.contentIndex).toBe(1);
  });

  test("finds visible text inside a simple field wrapped in an inline SDT", () => {
    const document: Document = {
      package: {
        document: {
          content: [
            {
              type: "paragraph",
              content: [
                {
                  type: "inlineSdt",
                  properties: {},
                  content: [
                    {
                      type: "simpleField",
                      instruction: "TITLE",
                      fieldType: "TITLE",
                      content: [
                        {
                          type: "run",
                          content: [{ type: "text", text: "Doc Title" }],
                        },
                      ],
                    },
                  ],
                },
              ],
            },
          ],
        },
      },
    };

    const matches = findInDocument(document, "Title", createDefaultFindOptions());

    expect(matches).toHaveLength(1);
    expect(matches[0]?.startOffset).toBe(4);
    expect(matches[0]?.endOffset).toBe(9);
  });

  for (const lookup of ["generated-block-id", "paragraph-order"] as const) {
    test(`find scrolling resolves the editor root through ${lookup}`, () => {
      const host = document.createElement("div");
      const container = document.createElement("div");
      container.setAttribute("data-folio-scroll", "");
      Object.defineProperty(container, "clientHeight", { value: 300 });
      container.getBoundingClientRect = () => new DOMRect(0, 100, 800, 300);
      const paragraphs = [document.createElement("p"), document.createElement("p")];
      for (const paragraph of paragraphs) paragraph.className = "layout-paragraph";
      const second = paragraphs[1];
      if (lookup === "generated-block-id") second.dataset["blockId"] = "block-2";
      second.getBoundingClientRect = () => new DOMRect(0, 700, 400, 40);
      container.append(...paragraphs);
      host.append(container);
      host.scrollTop = 77;
      const scroll = spyOn(container, "scrollTo").mockImplementation(() => {});
      scrollToMatch(container, {
        paragraphIndex: 1,
        contentIndex: 0,
        startOffset: 0,
        endOffset: 6,
        text: "Inside",
      });
      expect(scroll).toHaveBeenCalledWith(expect.objectContaining({ top: 470 }));
      scroll.mockRestore();
      expect(host.scrollTop).toBe(77);
    });
  }

  test("missing rendered find matches do not throw or move the root", () => {
    const container = document.createElement("div");
    container.setAttribute("data-folio-scroll", "");
    scrollToMatch(container, {
      paragraphIndex: 1,
      contentIndex: 0,
      startOffset: 0,
      endOffset: 6,
      text: "Inside",
    });
    expect(container.scrollTop).toBe(0);
  });
});
