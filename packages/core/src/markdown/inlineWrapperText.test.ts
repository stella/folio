/**
 * Markdown reads through a transparent inline wrapper.
 *
 * `w:bdo` and `w:dir` say how their text is laid out and a `w:sdt` says what
 * it is bound to; markdown carries none of that, but it does carry the text.
 * The inline renderers narrowed by a switch whose default contributed nothing,
 * so a paragraph whose text sat inside a bidirectional wrapper exported as an
 * empty line: the wrapper is the common way to write a right-to-left run, so
 * the loss falls entirely on right-to-left documents.
 */

import { describe, expect, test } from "bun:test";

import type { Document, Paragraph, Table } from "../types/document";
import { createEmptyDocument } from "../utils/createDocument";
import { toMarkdown } from "./index";

const RUN = { type: "run", content: [{ type: "text", text: "shalom" }] } as const;

const documentWith = (content: (Paragraph | Table)[]): Document => {
  const template = createEmptyDocument();
  return {
    ...template,
    package: {
      ...template.package,
      document: { ...template.package.document, content },
    },
  };
};

describe("markdown export of a transparent inline wrapper", () => {
  test("a bidirectional override keeps its text", () => {
    const markdown = toMarkdown(
      documentWith([
        {
          type: "paragraph",
          content: [{ type: "bidiWrapper", control: "override", direction: "rtl", content: [RUN] }],
        },
      ]),
    );
    expect(markdown).toContain("shalom");
  });

  test("a bidirectional embedding inside an HTML table cell keeps its text", () => {
    // `gridSpan` is what picks the HTML cell renderer over the GFM one; the
    // two narrow inline content separately, so both have to be asserted.
    const markdown = toMarkdown(
      documentWith([
        {
          type: "table",
          rows: [
            {
              cells: [
                {
                  formatting: { gridSpan: 2 },
                  content: [
                    {
                      type: "paragraph",
                      content: [
                        {
                          type: "bidiWrapper",
                          control: "embedding",
                          direction: "rtl",
                          content: [RUN],
                        },
                      ],
                    },
                  ],
                },
              ],
            },
          ],
        },
      ]),
    );
    expect(markdown).toContain("shalom");
  });
});
