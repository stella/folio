import assert from "node:assert/strict";
import { test } from "node:test";

import { paragraphNumberingFromSlots } from "@stll/folio-core/docx";
import { fromMarkdown } from "@stll/folio-core/markdown";

import { assertReadersAgree } from "../support/invariants.ts";
import { listDocument, packDocument } from "../support/documents.ts";
import { markdownViews, readAll } from "../support/readers.ts";

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

const numberedTableDocument = async (): Promise<Uint8Array> => {
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
        numFmt: "decimal",
        lvlText: "%1.",
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
  if (!paragraph) throw new Error("Numbered table cell fixture is missing");
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
