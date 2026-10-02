import assert from "node:assert/strict";
import { test } from "node:test";

import { paragraphNumberingFromSlots } from "@stll/folio-core/docx";
import { fromMarkdown } from "@stll/folio-core/markdown";
import { FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION } from "@stll/folio-core/server";

import { assertReadersAgree } from "../support/invariants.ts";
import {
  directNumberedDocument,
  listDocument,
  notesDocument,
  openReviewer,
  packDocument,
} from "../support/documents.ts";
import { contentView, markdownViews, readAll } from "../support/readers.ts";

type Document = ReturnType<typeof fromMarkdown>;
type Numbering = NonNullable<Document["package"]["numbering"]>;
type NumberingLevel = Numbering["abstractNums"][number]["levels"][number];
type Paragraph = Extract<Document["package"]["document"]["content"][number], { type: "paragraph" }>;
type Table = Extract<Document["package"]["document"]["content"][number], { type: "table" }>;

const paragraphs = (texts: readonly string[]) =>
  texts.map((text) => ({ text, kind: "paragraph" as const }));

const listGlyphDocument = async (): Promise<Uint8Array> => {
  const document = fromMarkdown("Bullet glyph\n\nLetter item");
  const bulletId = 91;
  const letterId = 92;
  const level = (numFmt: NumberingLevel["numFmt"], lvlText: string): NumberingLevel => ({
    ilvl: 0,
    start: 1,
    numFmt,
    lvlText,
    suffix: "space",
    pPr: { indentLeft: 0, indentFirstLine: 0 },
  });
  const numbering = document.package.numbering ?? { abstractNums: [], nums: [] };
  numbering.abstractNums.push(
    { abstractNumId: bulletId, multiLevelType: "singleLevel", levels: [level("bullet", "o")] },
    {
      abstractNumId: letterId,
      multiLevelType: "singleLevel",
      levels: [level("lowerLetter", "%1.")],
    },
  );
  numbering.nums.push(
    { numId: bulletId, abstractNumId: bulletId },
    { numId: letterId, abstractNumId: letterId },
  );
  document.package.numbering = numbering;
  const paragraphBlocks = document.package.document.content.filter(
    (
      block,
    ): block is Extract<
      Document["package"]["document"]["content"][number],
      { type: "paragraph" }
    > => block.type === "paragraph",
  );
  const bullet = paragraphBlocks.find((paragraph) =>
    paragraph.content.some(
      (item) =>
        item.type === "run" &&
        item.content.some((run) => run.type === "text" && run.text === "Bullet glyph"),
    ),
  );
  const letter = paragraphBlocks.find((paragraph) =>
    paragraph.content.some(
      (item) =>
        item.type === "run" &&
        item.content.some((run) => run.type === "text" && run.text === "Letter item"),
    ),
  );
  if (!bullet || !letter) throw new Error("Numbering fixture paragraphs are missing");
  bullet.formatting = {
    ...bullet.formatting,
    numPr: paragraphNumberingFromSlots({ numId: bulletId, ilvl: 0 }),
  };
  letter.formatting = {
    ...letter.formatting,
    numPr: paragraphNumberingFromSlots({ numId: letterId, ilvl: 0 }),
  };
  return packDocument(document);
};

const numberedTableDocument = async (
  format: "decimal" | "bullet" = "decimal",
  rendering: "gfm" | "html" = "gfm",
): Promise<Uint8Array> => {
  const document = fromMarkdown(
    "## Terms\n\n| Service | Value |\n| --- | --- |\n| Term | Included |",
  );
  const numId = 93;
  const numbering = document.package.numbering ?? { abstractNums: [], nums: [] };
  numbering.abstractNums.push({
    abstractNumId: numId,
    multiLevelType: "singleLevel",
    levels: [
      {
        ilvl: 0,
        start: 3,
        numFmt: format,
        lvlText: format === "bullet" ? "o" : "%1.",
        suffix: "space",
        pPr: { indentLeft: 0, indentFirstLine: 0 },
      },
    ],
  });
  numbering.nums.push({ numId, abstractNumId: numId });
  document.package.numbering = numbering;
  const table = document.package.document.content.find(
    (block): block is Table => block.type === "table",
  );
  const cell = table?.rows.at(1)?.cells.at(0);
  const paragraph = cell?.content.find((block): block is Paragraph => block.type === "paragraph");
  if (!paragraph || !cell || !table) throw new Error("Numbered table cell fixture is missing");
  if (rendering === "html") {
    cell.formatting = { ...cell.formatting, gridSpan: 2 };
    table.rows.at(1)?.cells.splice(1, 1);
  }
  paragraph.formatting = {
    ...paragraph.formatting,
    numPr: paragraphNumberingFromSlots({ numId, ilvl: 0 }),
  };
  return packDocument(document);
};

test("uses package numFmt to distinguish an alphabetic bullet glyph from a letter list", async () => {
  const bytes = await listGlyphDocument();
  const views = await readAll(bytes);
  assert.deepEqual(
    views.getContent.map(({ text, number }) => ({ text, number })),
    [
      { text: "Bullet glyph", number: "(bullet)" },
      { text: "Letter item", number: "a." },
    ],
  );
  assert.deepEqual(views.markdown, views.getContentAsMarkdown);
  await assertReadersAgree(bytes, "alphabetic bullet glyph reader fixture");
});

test("parses Markdown block boundaries while retaining source hard breaks", () => {
  const cases = [
    {
      markdown: "First line  \nSecond line\\\nThird line\n\nNext paragraph",
      texts: ["First line\nSecond line\nThird line", "Next paragraph"],
      first: "First line\nSecond line\nThird line",
    },
    {
      markdown: "First line  \nSecond line\\\nThird line  \nFourth line\n\nNext paragraph",
      texts: ["First line\nSecond line\nThird line\nFourth line", "Next paragraph"],
      first: "First line\nSecond line\nThird line\nFourth line",
    },
  ];
  for (const { markdown, texts, first } of cases) {
    assert.deepEqual(markdownViews(markdown, paragraphs(texts)), [
      { text: first, kind: "paragraph" },
      { text: "Next paragraph", kind: "paragraph" },
    ]);
  }
});

test("ignores opaque images and compares visible link text", () => {
  const markdown =
    "Text ![diagram](data:image/svg+xml,%3Csvg%3E(a(b))%3C/svg%3E) [visible text](https://example.test/a_(b))";
  assert.deepEqual(markdownViews(markdown, paragraphs(["Text  visible text"])), [
    { text: "Text  visible text", kind: "paragraph" },
  ]);
});

test("keeps a numbered table cell label in the Markdown comparison", async () => {
  const bytes = await numberedTableDocument();
  const views = await readAll(bytes);
  assert.ok(views.markdown.some(({ text }) => text === "3. Term"));
  assert.deepEqual(views.markdown, views.getContentAsMarkdown);

  const mutated = views.markdown.map((block) =>
    block.text === "3. Term" ? { ...block, text: "4. Term" } : block,
  );
  assert.notDeepEqual(mutated, views.getContentAsMarkdown);
  await assertReadersAgree(bytes, "numbered table cell reader fixture");
});

test("normalizes bullet table cells in GFM and HTML without hiding changed text", async () => {
  for (const rendering of ["gfm", "html"] as const) {
    const bytes = await numberedTableDocument("bullet", rendering);
    const views = await readAll(bytes);
    assert.ok(
      views.getContent.some(({ text, number }) => text === "Term" && number === "(bullet)"),
    );
    assert.ok(views.markdown.some(({ text }) => text === "Term"));
    assert.deepEqual(views.markdown, views.getContentAsMarkdown);
    assert.notDeepEqual(
      views.markdown.map((block) =>
        block.text === "Term" ? { ...block, text: "Changed" } : block,
      ),
      views.getContentAsMarkdown,
    );
    await assertReadersAgree(bytes, `${rendering} bullet table cell reader fixture`);
  }
  const expected = [{ text: "Term", kind: "listItem" as const, number: "(bullet)" }];
  assert.deepEqual(markdownViews("| - Term |\n| --- |", expected), paragraphs(["Term"]));
  assert.notDeepEqual(markdownViews("| - Changed |\n| --- |", expected), paragraphs(["Term"]));
  assert.deepEqual(
    markdownViews("| - Literal prose |\n| --- |", paragraphs(["- Literal prose"])),
    paragraphs(["- Literal prose"]),
  );
});

test("keeps ordinary list reader agreement", async () => {
  await assertReadersAgree(await listDocument(), "ordinary list reader fixture");
});

test("retains every rendered ordered-list label, including restarts", () => {
  const expected = [
    { text: "First", kind: "listItem", number: "1." },
    { text: "Restarted", kind: "listItem", number: "1." },
  ] as const;
  assert.deepEqual(markdownViews("1. First\n1. Restarted", expected), expected);
  assert.notDeepEqual(markdownViews("1. First\n2. Restarted", expected), expected);
});

test("an undefined numbering level without a displayed marker reads as prose", () => {
  const formats = new Map([["7:0", "decimal"]]);
  const block = {
    id: "4207D525",
    text: "No marker",
    kind: "paragraph",
    listReference: { numId: 7, level: 8 },
  };
  assert.deepEqual(contentView(block, formats), { text: "No marker", kind: "paragraph" });
  assert.throws(
    () => contentView({ ...block, displayLabel: "1." }, formats),
    /Missing numbering format for displayed label/,
  );
});

test("custom markers retain boundaries, hard breaks and actual rendered labels", () => {
  const cases = [
    { markers: ["(1)", "(2)", "(1)"] },
    { markers: ["a.", "b.", "a."] },
    { markers: ["i)", "ii)", "i)"] },
    { markers: ["1.", "(2)", "a."] },
  ];
  for (const { markers } of cases) {
    // The first item contains a marker-shaped hard-break line. It is one
    // source block; only a matching complete source prefix grants a split.
    const texts = ["First\n(a) stays inside the first item", "Second", "Restarted"];
    const expected = texts.map((text, index) => {
      const number = markers.at(index);
      if (number === undefined) throw new Error("List-marker fixture is missing a label");
      return { text, kind: "listItem" as const, number };
    });
    const lines = expected.map(({ text, number }) => `${number} ${text}`);
    const markdown = lines.join("\n");
    assert.deepEqual(markdownViews(markdown, expected), expected);
    assert.throws(
      () => markdownViews(markdown.replace("Second", "Changed"), expected),
      /fewer text blocks/,
    );
    const secondMarker = markers.at(1);
    if (secondMarker === undefined)
      throw new Error("List-marker fixture is missing its second label");
    assert.notDeepEqual(
      markdownViews(markdown.replace(secondMarker, "wrong)"), expected),
      expected,
    );
    assert.throws(() => markdownViews(lines.slice(0, 2).join("\n"), expected), /fewer text blocks/);
  }
});

test("removing the prose between custom-numbered blocks keeps every saved reader block", async () => {
  for (const mode of ["direct", "tracked-changes"] as const) {
    const reviewer = await openReviewer(await directNumberedDocument());
    const separator = reviewer.getContent().find(({ text }) => text === "Unnumbered body text.");
    assert.ok(separator);
    const result = reviewer.applyDocumentOperations({
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode,
      operations: [{ id: "delete-separator", type: "deleteBlock", blockId: separator.id }],
    });
    assert.equal(result.applied.length, 1);
    reviewer.acceptAll();
    await assertReadersAgree(
      new Uint8Array(await reviewer.toBuffer()),
      `adjacent custom labels ${mode}`,
    );
  }
});

// The lexer treats exported footnote definitions as paragraph text; the body
// reader excludes note stories. Keep their trailer out of the body comparison.
test("separates appended exported note definitions from body paragraphs", () => {
  const expected = paragraphs(["Body1", "Endnote†"]);
  for (const definitions of [
    "[^1]: Footnote text\n[^e1]: Endnote text",
    "[^1]: **Footnote** text\n[^e1]: *Endnote* text",
    "[^1]: \n[^e1]: Endnote text",
  ]) {
    assert.deepEqual(
      markdownViews(`Body[^1]\n\nEndnote[^e1]\n\n${definitions}`, expected),
      expected,
    );
  }
});

test("note trailers cannot hide an extra body paragraph", () => {
  const expected = paragraphs(["Body1"]);
  for (const markdown of [
    "Body[^1]\n\nUnexpected body\n\n[^1]: Footnote text",
    "Body[^1]\n\n[^1]: Footnote text\n\nUnexpected body",
    "Body[^1]\n\n[^1]: Footnote text\nUnexpected body",
    "Body[^1]\n\nUnexpected body",
  ]) {
    assert.throws(
      () => markdownViews(markdown, expected),
      /Markdown contains (?:more text blocks than the source reader|malformed or unreferenced note definitions)/,
    );
  }
});

test("extracts visible nested inline text without losing leaf-token fallback text", () => {
  const cases = [
    { markdown: "Plain leaf text", text: "Plain leaf text" },
    { markdown: "**bold *nested* tail**", text: "bold nested tail" },
    { markdown: "~~deleted **bold *nested*** tail~~", text: "deleted bold nested tail" },
    {
      markdown: "[**linked *nested* text**](https://example.test/target)",
      text: "linked nested text",
    },
    { markdown: "Text `literal *stars*` tail", text: "Text literal *stars* tail" },
    { markdown: "Text \\*escaped\\* tail", text: "Text *escaped* tail" },
  ];
  for (const { markdown, text } of cases) {
    const expected = paragraphs([text]);
    assert.deepEqual(markdownViews(markdown, expected), expected);
  }
});

test("keeps packed footnote and endnote reader agreement", async () => {
  const bytes = await notesDocument();
  const views = await readAll(bytes);
  assert.deepEqual(views.markdown, views.getContentAsMarkdown);
  assert.equal(views.markdown.length, views.getContent.length);
  await assertReadersAgree(bytes, "footnote and endnote definition trailers");
});

// Definition recognition is narrower than a paragraph starting with [^...].
test("rejects malformed, duplicate and unreferenced note definitions", () => {
  const expected = paragraphs(["Body1"]);
  for (const definitions of [
    "[^1]: First\n[^1]: Duplicate",
    "[^2]: Unreferenced",
    "[^1]: First\nNot a definition",
  ]) {
    assert.throws(
      () => markdownViews(`Body[^1]\n\n${definitions}`, expected),
      /Markdown contains malformed or unreferenced note definitions/,
    );
  }
});
