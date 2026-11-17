import assert from "node:assert/strict";
import { test } from "node:test";

import { fromMarkdown } from "@stll/folio-core/markdown";
import { docxToMarkdown } from "@stll/folio-core/server";

import { packDocument, toArrayBuffer } from "../support/documents.ts";
import { assertReadersAgree } from "../support/invariants.ts";
import {
  MARKDOWN_READ_OPTIONS,
  markdownViews,
  readAll,
} from "../support/readers.ts";

for (const rendering of ["gfm", "html"] as const)
  for (const { source, expected } of [
    { source: "Alpha\n\nBeta\n\nGamma", expected: "Alpha\nBeta\nGamma" },
    {
      source: "Alpha  \nContinued\n\nBeta",
      expected: "Alpha\nContinued\nBeta",
    },
    { source: "- Alpha\n\n- Beta", expected: "Alpha\nBeta" },
    { source: "\\- Alpha\n\n\\- Beta", expected: "- Alpha\n- Beta" },
  ])
    test(`${rendering} keeps every paragraph and hard break within one cell: ${source}`, async () => {
      const document = fromMarkdown(
        `${source}\n\n| Heading | Other |\n| --- | --- |\n| Alpha | Beta |\n\nOutside.`,
      );
      const blocks = document.package.document.content;
      const tableIndex = blocks.findIndex(({ type }) => type === "table");
      const table = blocks.at(tableIndex);
      assert.equal(table?.type, "table");
      if (table?.type !== "table") assert.fail("The fixture has a table");
      const row = table.rows.at(0);
      const cell = row?.cells.at(0);
      assert.ok(cell && row);
      cell.content = blocks.splice(0, tableIndex);
      if (rendering === "html") {
        cell.formatting = { gridSpan: 2 };
        row.cells.splice(1, 1);
      }
      const bytes = await packDocument(document);
      const views = await readAll(bytes);
      assert.deepEqual(views.markdown, views.getContentAsMarkdown);
      assert.equal(
        views.getContent.length - views.markdown.length,
        cell.content.length - 1,
      );
      assert.equal(views.markdown.at(0)?.text, expected);
      await assertReadersAgree(
        bytes,
        "multiple paragraphs in one Markdown cell",
      );

      const markdown = await docxToMarkdown(
        toArrayBuffer(bytes),
        MARKDOWN_READ_OPTIONS,
      );
      const expectedMarkdown = views.getContentAsMarkdown.map((block, index) =>
        index === 0
          ? {
              ...block,
              cellParagraphs: views.getContent.slice(0, cell.content.length),
            }
          : block,
      );
      for (const replacement of ["Changed", "", "Alpha"]) {
        const changed = markdown.replace("Beta", replacement);
        assert.notEqual(changed, markdown);
        assert.notDeepEqual(
          markdownViews(changed, expectedMarkdown),
          views.getContentAsMarkdown,
        );
      }
    });

for (const rendering of ["gfm", "html"] as const)
  test(`${rendering} cannot move a paragraph between adjacent cells while preserving word order`, async () => {
    const document = fromMarkdown(
      "Beta\n\nDelta\n\n| Alpha | Gamma |\n| --- | --- |",
    );
    const [beta, delta] = document.package.document.content.splice(0, 2);
    const table = document.package.document.content.find(
      ({ type }) => type === "table",
    );
    assert.ok(beta && delta && table?.type === "table");
    const row = table.rows.at(0);
    const first = row?.cells.at(0);
    const second = row?.cells.at(1);
    assert.ok(first && second);
    first.content.push(beta);
    second.content.push(delta);
    if (rendering === "html") {
      first.formatting = { gridSpan: 2 };
      table.columnWidths = [1000, 1000, 1000];
    }
    const views = await readAll(await packDocument(document));
    assert.deepEqual(views.getContentAsMarkdown, [
      { kind: "paragraph", text: "Alpha\nBeta" },
      { kind: "paragraph", text: "Gamma\nDelta" },
    ]);
    assert.deepEqual(views.markdown, views.getContentAsMarkdown);
    const moved =
      rendering === "gfm"
        ? "| Alpha | Beta<br>Gamma<br>Delta |\n| --- | --- |"
        : '<table><tr><th colspan="2">Alpha</th><th>Beta<br>Gamma<br>Delta</th></tr></table>';
    assert.notDeepEqual(
      markdownViews(moved, views.getContentAsMarkdown),
      views.getContentAsMarkdown,
    );
  });
