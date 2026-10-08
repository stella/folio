import assert from "node:assert/strict";
import { test } from "node:test";
import { fromMarkdown } from "@stll/folio-core/markdown";
import { docxToMarkdown } from "@stll/folio-core/server";

import { packDocument, toArrayBuffer } from "../support/documents.ts";
import { MARKDOWN_READ_OPTIONS, markdownViews } from "../support/readers.ts";

test("inline breaks preserve Markdown block identity across paragraph kinds", async () => {
  const starts = {
    paragraph: { markdown: "First", view: { kind: "paragraph" } },
    heading: { markdown: "## First", view: { kind: "heading", headingLevel: 2 } },
    listItem: { markdown: "- First", view: { kind: "listItem", number: "(bullet)" } },
    tableCell: { markdown: "| First |\n| --- |", view: { kind: "paragraph" } },
  } as const;
  const separators = { textWrapping: "\n", column: "\n", tab: "    ", softHyphen: "" } as const;
  for (const [kind, start] of Object.entries(starts)) {
    for (const inline of ["textWrapping", "column", "tab", "softHyphen"] as const) {
      for (const formatting of [
        {},
        { bold: true, italic: true },
        { fontFamily: { ascii: "Consolas" } },
      ]) {
        const document = fromMarkdown(start.markdown);
        const first = document.package.document.content.at(0);
        assert.ok(first);
        const paragraph =
          first.type === "table" ? first.rows.at(0)?.cells.at(0)?.content.at(0) : first;
        assert.ok(paragraph?.type === "paragraph");
        paragraph.content = [
          {
            type: "run",
            formatting,
            content: [
              { type: "text", text: "First" },
              inline === "tab" || inline === "softHyphen"
                ? { type: inline }
                : { type: "break", breakType: inline },
              { type: "text", text: "Second" },
            ],
          },
        ];
        const markdown = await docxToMarkdown(
          toArrayBuffer(await packDocument(document)),
          MARKDOWN_READ_OPTIONS,
        );
        const text = `First${separators[inline]}Second`;
        const expected = { text, ...start.view };
        assert.deepEqual(
          markdownViews(markdown, [expected]),
          [expected],
          `${kind} ${inline} ${JSON.stringify(formatting)}`,
        );
      }
    }
  }
});
