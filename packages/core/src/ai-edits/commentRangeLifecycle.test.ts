/**
 * A comment's thread lives exactly as long as its anchor, and the reviewer
 * shows the scope the saved package will reopen with.
 *
 * Word removes a comment with the content it is anchored to: the
 * `w:commentReference` run sits inside that content (ECMA-376 Part 1
 * §17.13.4.5), so rejecting the insertion that holds it, or deleting the
 * paragraph it covers, takes the reference and with it the only place the
 * comment is shown. A `comments.xml` definition nothing references
 * (§17.13.4.2) is a thread about nothing. The reviewer used to keep listing
 * it and wrote it back on save.
 *
 * A comment range is two points in document order (§17.13.4.3, §17.13.4.4),
 * so whatever lands between them is covered, whether the range crosses
 * paragraphs or a paragraph and a table. The live reading used to miss a
 * paragraph inserted before a table the range ended in, which the saved
 * package then covered.
 */

import { describe, expect, test } from "bun:test";
import JSZip from "jszip";

import { createDocx } from "../docx/rezip";
import type { Comment, Document, Paragraph, ParagraphContent, Table } from "../types/document";
import { FolioDocxReviewer, type FolioReviewComment } from "./headless";

const run = (text: string): ParagraphContent => ({
  type: "run",
  content: [{ type: "text", text }],
});

const comment = (id: number, text: string): Comment => ({
  id,
  author: "Dana Lindqvist",
  initials: "DL",
  date: "2024-01-01T00:00:00Z",
  content: [{ type: "paragraph", content: [run(text)] }],
});

const paragraph = (paraId: string, content: ParagraphContent[]): Paragraph => ({
  type: "paragraph",
  paraId,
  content,
});

const table = (cells: Paragraph[][]): Table => ({
  type: "table",
  rows: [
    {
      type: "tableRow",
      cells: cells.map((content) => ({ type: "tableCell", content })),
    },
  ],
});

const reopen = async (reviewer: FolioDocxReviewer): Promise<FolioDocxReviewer> =>
  FolioDocxReviewer.fromBuffer(await reviewer.toBuffer(), { author: "Editor" });

/** The thread fields a save must round-trip. */
const threads = (comments: readonly FolioReviewComment[]) =>
  comments.map(({ id, text, anchoredText, replies }) => ({
    id,
    text,
    anchoredText,
    replies: replies.map((reply) => reply.text),
  }));

const packageCounts = async (buffer: ArrayBuffer) => {
  const zip = await JSZip.loadAsync(buffer);
  const documentXml = (await zip.file("word/document.xml")?.async("text")) ?? "";
  const commentsXml = (await zip.file("word/comments.xml")?.async("text")) ?? "";
  return {
    rangeStarts: (documentXml.match(/<w:commentRangeStart\b/gu) ?? []).length,
    references: (documentXml.match(/<w:commentReference\b/gu) ?? []).length,
    definitions: (commentsXml.match(/<w:comment\b(?!s)/gu) ?? []).length,
  };
};

const twoParagraphs = async (): Promise<FolioDocxReviewer> => {
  const document: Document = {
    package: {
      document: {
        content: [
          paragraph("30000001", [run("Original.")]),
          paragraph("30000002", [
            { type: "commentRangeStart", id: 1 },
            run("Keep."),
            { type: "commentRangeEnd", id: 1 },
            { type: "commentReference", id: 1 },
          ]),
        ],
        comments: [comment(1, "On surviving text.")],
      },
    },
  };
  return FolioDocxReviewer.fromBuffer(await createDocx(document), { author: "Editor" });
};

/** Insert tracked text after the first block and comment on all of it. */
const commentOnTrackedInsertion = async (): Promise<FolioDocxReviewer> => {
  const reviewer = await twoParagraphs();
  const [original] = reviewer.snapshot().blocks;
  const inserted = reviewer.applyOperations(
    [{ id: "insert", type: "insertAfterBlock", blockId: original!.id, text: "Temporary." }],
    { mode: "tracked-changes" },
  );
  expect(inserted.applied).toHaveLength(1);
  const target = reviewer.snapshot().blocks.find((block) => block.text === "Temporary.");
  const commented = reviewer.applyOperations(
    [
      {
        id: "comment",
        type: "commentOnBlock",
        blockId: target!.id,
        comment: { text: "On temporary text." },
      },
    ],
    { mode: "direct" },
  );
  expect(commented.applied).toHaveLength(1);
  expect(threads(reviewer.getComments()).map(({ anchoredText }) => anchoredText)).toEqual([
    "Keep.",
    "Temporary.",
  ]);
  return reviewer;
};

const SURVIVOR = { id: 1, text: "On surviving text.", anchoredText: "Keep.", replies: [] };

describe("a comment whose anchored content is removed goes with it", () => {
  const cases: [string, (reviewer: FolioDocxReviewer) => void][] = [
    ["reject all", (reviewer) => reviewer.rejectAll()],
    [
      // The inserted text and the paragraph mark that opens it are two
      // revisions; rejecting both is rejecting the insertion.
      "reject that change",
      (reviewer) => {
        const changes = reviewer.getChanges().filter((entry) => entry.text === "Temporary.");
        expect(changes.map(({ type }) => type).toSorted()).toEqual([
          "insertion",
          "paragraphMarkInserted",
        ]);
        for (const change of changes) {
          expect(reviewer.rejectChange(change)).toBe(true);
        }
      },
    ],
  ];

  for (const [name, reject] of cases) {
    test(`${name} before the save`, async () => {
      const reviewer = await commentOnTrackedInsertion();
      const threadId = reviewer.getComments().find((entry) => entry.text === "On temporary text.");
      reviewer.replyTo(threadId!, { text: "A reply." });
      reject(reviewer);

      expect(threads(reviewer.getComments())).toEqual([SURVIVOR]);
      const saved = await reviewer.toBuffer();
      expect(await packageCounts(saved)).toEqual({ rangeStarts: 1, references: 1, definitions: 1 });
      const reopened = await FolioDocxReviewer.fromBuffer(saved);
      expect(threads(reopened.getComments())).toEqual([SURVIVOR]);
    });

    test(`${name} after a save and reopen`, async () => {
      const reopened = await reopen(await commentOnTrackedInsertion());
      reject(reopened);

      expect(threads(reopened.getComments())).toEqual([SURVIVOR]);
      const final = await reopen(reopened);
      expect(threads(final.getComments())).toEqual([SURVIVOR]);
      expect(await packageCounts(await final.toBuffer())).toEqual({
        rangeStarts: 1,
        references: 1,
        definitions: 1,
      });
    });
  }

  test("accepting the insertion keeps the comment", async () => {
    const reviewer = await reopen(await commentOnTrackedInsertion());
    reviewer.acceptAll();
    const expected = [
      SURVIVOR,
      {
        id: expect.any(Number),
        text: "On temporary text.",
        anchoredText: "Temporary.",
        replies: [],
      },
    ];
    expect(threads(reviewer.getComments())).toEqual(expected);
    expect(threads((await reopen(reviewer)).getComments())).toEqual(expected);
  });

  test("deleting the only paragraph a comment covers removes the comment", async () => {
    const reviewer = await twoParagraphs();
    const kept = reviewer.snapshot().blocks.find((block) => block.text === "Keep.");
    reviewer.applyOperations([{ id: "delete", type: "deleteBlock", blockId: kept!.id }], {
      mode: "direct",
    });
    expect(reviewer.getComments()).toEqual([]);
    const saved = await reviewer.toBuffer();
    expect((await packageCounts(saved)).definitions).toBe(0);
    expect((await FolioDocxReviewer.fromBuffer(saved)).getComments()).toEqual([]);
  });

  test("a tracked deletion keeps the comment until it is accepted", async () => {
    const reviewer = await twoParagraphs();
    const kept = reviewer.snapshot().blocks.find((block) => block.text === "Keep.");
    reviewer.applyOperations([{ id: "delete", type: "deleteBlock", blockId: kept!.id }], {
      mode: "tracked-changes",
    });
    expect(threads(reviewer.getComments())).toEqual([SURVIVOR]);
    expect(threads((await reopen(reviewer)).getComments())).toEqual([SURVIVOR]);

    reviewer.acceptAll();
    expect(reviewer.getComments()).toEqual([]);
    expect((await reopen(reviewer)).getComments()).toEqual([]);
  });

  test("a comment the source package already left unanchored is kept", async () => {
    const document: Document = {
      package: {
        document: {
          content: [paragraph("30000001", [run("Original.")])],
          comments: [comment(4, "Anchored nowhere.")],
        },
      },
    };
    const reviewer = await FolioDocxReviewer.fromBuffer(await createDocx(document));
    const [first] = reviewer.snapshot().blocks;
    reviewer.applyOperations(
      [{ id: "insert", type: "insertAfterBlock", blockId: first!.id, text: "More." }],
      { mode: "direct" },
    );
    const expected = [{ id: 4, text: "Anchored nowhere.", anchoredText: "", replies: [] }];
    expect(threads(reviewer.getComments())).toEqual(expected);
    expect(threads((await reopen(reviewer)).getComments())).toEqual(expected);
  });
});

describe("a comment spanning a paragraph and a table", () => {
  const spanningDocument = (): Document => ({
    package: {
      document: {
        content: [
          paragraph("31000001", [{ type: "commentRangeStart", id: 7 }, run("Before.")]),
          table([
            [
              paragraph("31000002", [
                run("Cell A."),
                { type: "commentRangeEnd", id: 7 },
                { type: "commentReference", id: 7 },
              ]),
            ],
            [paragraph("31000003", [run("Cell B.")])],
          ]),
          paragraph("31000004", [run("After.")]),
        ],
        comments: [comment(7, "Cross-table comment.")],
      },
    },
  });

  const cases: [string, string, string][] = [
    ["between the paragraph and the table", "Before.", "Before.Inserted.Cell A."],
    ["after the range's last cell paragraph", "Cell A.", "Before.Cell A."],
    ["after the paragraph that follows the table", "After.", "Before.Cell A."],
  ];

  for (const mode of ["direct", "tracked-changes"] as const) {
    for (const [where, after, expected] of cases) {
      test(`inserting ${where} (${mode}) reads the same live and reopened`, async () => {
        const reviewer = await FolioDocxReviewer.fromBuffer(await createDocx(spanningDocument()), {
          author: "Editor",
        });
        expect(reviewer.getComments()[0]?.anchoredText).toBe("Before.Cell A.");
        const target = reviewer.snapshot().blocks.find((block) => block.text === after);
        const receipt = reviewer.applyOperations(
          [{ id: "insert", type: "insertAfterBlock", blockId: target!.id, text: "Inserted." }],
          { mode },
        );
        expect(receipt.applied).toHaveLength(1);

        const live = threads(reviewer.getComments());
        expect(live.map(({ anchoredText }) => anchoredText)).toEqual([expected]);
        expect(threads((await reopen(reviewer)).getComments())).toEqual(live);
      });
    }
  }
});
