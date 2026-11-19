/**
 * A tracked batch on the canonical session takes revision ids past every
 * one the package holds, a table's grid change included, which the editor
 * projection does not show.
 */
import { expect, test } from "bun:test";

import { FolioDocxReviewer } from "../ai-edits/headless";
import { FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION } from "../document-operations";
import { createDocx } from "../docx/rezip";
import { paragraph, run, table } from "../docx/server/build";
import { fromMarkdown } from "../markdown/fromMarkdown";
import type { Table } from "../types/document";

test("a tracked edit takes a revision id past a table's pending grid change", async () => {
  const document = fromMarkdown("Opening words.");
  const grid = table({
    header: ["Item", "Price"],
    rows: [["Widget", "10"]],
  }) as Table;
  // The grid change holds the id right after the last one the projection
  // shows, so a seed read from the projection alone takes it.
  grid.formatting = {
    ...grid.formatting,
    gridChange: { id: 50, columnWidths: [null, null] },
  } as Table["formatting"];
  document.package.document.content = [
    paragraph([
      run("Opening "),
      {
        type: "insertion",
        info: { id: 49, author: "Reviewer B", date: "2026-01-01T00:00:00Z" },
        content: [run("added ")],
      },
      run("words."),
    ]),
    grid,
    paragraph([run("Closing words.")]),
  ];
  const reviewer = await FolioDocxReviewer.fromBuffer(
    await createDocx(document),
    { author: "Editor", session: "canonical" },
  );
  const closing = reviewer
    .snapshot()
    .blocks.find(({ text }) => text === "Closing words.");
  if (closing === undefined) throw new Error("the fixture has no closing");
  const result = reviewer.applyDocumentOperations(
    {
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode: "tracked-changes",
      operations: [
        {
          id: "after",
          type: "insertAfterBlock",
          blockId: closing.id,
          text: "Added words.",
        },
      ],
    },
    { undefinedReferences: "refuse" },
  );
  expect(result.skipped).toEqual([]);
  expect(result.applied.map(({ id }) => id)).toEqual(["after"]);
});
