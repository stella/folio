/**
 * Every synthetic document an integrator can build opens, saves unchanged,
 * reopens, and reads alike through every reader.
 */

import assert from "node:assert/strict";
import { describe, test } from "node:test";

import { fromMarkdown, toMarkdown } from "@stll/folio-core/markdown";
import { docxToMarkdown, parseDocx } from "@stll/folio-core/server";

import {
  FIXTURE_NAMES,
  FIXTURES,
  openReviewer,
  packDocument,
  styleNumberedDocument,
  toArrayBuffer,
} from "../support/documents.ts";
import { assertHealthy, assertReadersAgree } from "../support/invariants.ts";
import { expectedFailure } from "../support/known-issues.ts";
import { MARKDOWN_READ_OPTIONS } from "../support/readers.ts";

describe("synthetic documents", () => {
  for (const name of FIXTURE_NAMES) {
    test(`${name}: opens, saves, reopens and reads alike`, async () => {
      const bytes = await FIXTURES[name]();
      await assertReadersAgree(bytes, `${name} as built`);
      const reviewer = await openReviewer(bytes);
      const { bytes: saved } = await assertHealthy(reviewer, `${name} saved untouched`);
      // A second save of the reopened package is a fixed point.
      await assertHealthy(await openReviewer(saved), `${name} saved twice`);
    });
  }

  test("fixtures carry what their names promise", async () => {
    const lists = await openReviewer(await FIXTURES.lists());
    assert.deepEqual(
      lists
        .getContent()
        .filter((block) => block.kind === "listItem")
        .map((block) => block.displayLabel),
      ["•", "•", "1.", "2.", "3."],
    );
    const tables = await openReviewer(await FIXTURES.tables());
    assert.equal(tables.getContent().filter((block) => block.table !== undefined).length, 12);
    const notes = await openReviewer(await FIXTURES.notes());
    assert.match(notes.getNotesAsText(), /As defined in the order form\./u);
    assert.match(notes.getNotesAsText(), /See the warranty schedule\./u);
    const comments = await openReviewer(await FIXTURES.comments());
    assert.deepEqual(
      comments
        .getComments()
        .map((comment) => [comment.text, comment.done, comment.replies.length])
        .sort(),
      [
        ["Define good order.", false, 1],
        ["Thirty days is long.", true, 0],
      ],
    );
    const tracked = await openReviewer(await FIXTURES.trackedChanges());
    assert.deepEqual([...new Set(tracked.getChanges().map((change) => change.type))].sort(), [
      "deletion",
      "insertion",
      "paragraphMarkDeleted",
      "paragraphMarkInserted",
    ]);
    const numbered = await openReviewer(await styleNumberedDocument());
    assert.deepEqual(
      numbered
        .getContent()
        .filter((block) => block.headingLevel !== undefined)
        .map((block) => [block.text, block.headingLevel, block.displayLabel]),
      [
        ["Agreement", 1, "Heading1"],
        ["Scope", 2, "1."],
        ["Definitions", 3, "1.1."],
        ["Payment", 2, "2."],
      ],
    );
  });

  expectedFailure(
    1094,
    "a style-numbered heading is a heading to every reader, and read rows carry its number",
    /read_document|docxToMarkdown vs getContent|snapshot vs getContent/u,
    async () => {
      await assertReadersAgree(await styleNumberedDocument(), "style-numbered headings", {
        strict: true,
      });
    },
  );
});

describe("markdown round trips", () => {
  const STABLE = [
    "# Title\n\nA paragraph.\n\n## Section\n\nAnother paragraph.",
    "Intro.\n\n- one\n- two\n- three\n\nOutro.",
    "Intro.\n\n1. first\n2. second\n3. third\n\nOutro.",
    "| A | B |\n| --- | --- |\n| 1 | 2 |",
    "Some **bold** and *italic* text.",
  ];

  for (const markdown of STABLE) {
    test(`fromMarkdown → docx → docxToMarkdown is stable: ${JSON.stringify(markdown.slice(0, 32))}`, async () => {
      const bytes = await packDocument(fromMarkdown(markdown));
      const first = await docxToMarkdown(toArrayBuffer(bytes), MARKDOWN_READ_OPTIONS);
      const again = await docxToMarkdown(
        toArrayBuffer(await packDocument(fromMarkdown(first))),
        MARKDOWN_READ_OPTIONS,
      );
      assert.equal(again, first);
      // The in-memory model and the parsed package render alike.
      assert.equal(toMarkdown(await parseDocx(toArrayBuffer(bytes)), MARKDOWN_READ_OPTIONS), first);
    });
  }

  // Markdown has no note syntax fromMarkdown reads back, so a note fixture's
  // Markdown is not meant to survive a second pass.
  for (const name of FIXTURE_NAMES.filter((fixture) => fixture !== "notes")) {
    test(`${name}: its Markdown reads back to the same Markdown`, async () => {
      const first = await docxToMarkdown(
        toArrayBuffer(await FIXTURES[name]()),
        MARKDOWN_READ_OPTIONS,
      );
      const again = await docxToMarkdown(
        toArrayBuffer(await packDocument(fromMarkdown(first))),
        MARKDOWN_READ_OPTIONS,
      );
      assert.equal(again, first);
    });
  }
});
