import { describe, expect, test } from "bun:test";
import { panic } from "better-result";
import JSZip from "jszip";

import { FolioDocxReviewer } from "../ai-edits/headless";
import type { FolioAIParagraphSpacing } from "../ai-edits/types";
import { createDocx } from "../docx/rezip";
import type { Paragraph, ParagraphAlignment, Table } from "../types/document";
import { createEmptyDocument } from "../utils/createDocument";
import { compareDocx } from "./compare";

const OPTIONS = { author: "compare", timestamp: "2026-09-11T00:00:00.000Z" } as const;

type DirectParagraphFormatting = {
  styleId: string;
  alignment: ParagraphAlignment;
  spacing: FolioAIParagraphSpacing;
};

const JOINED_FORMATTING = {
  styleId: "JoinedBody",
  alignment: "left",
  spacing: { spaceBefore: 120, spaceAfter: 60 },
} as const satisfies DirectParagraphFormatting;

const FIRST_SPLIT_FORMATTING = {
  styleId: "OpeningBody",
  alignment: "center",
  spacing: { spaceBefore: 240, lineSpacing: 360, lineSpacingRule: "exact" },
} as const satisfies DirectParagraphFormatting;

const SECOND_SPLIT_FORMATTING = {
  styleId: "ClosingBody",
  alignment: "right",
  spacing: { spaceAfter: 360, afterAutospacing: true },
} as const satisfies DirectParagraphFormatting;

const JOINED_BLOCKS = [
  { paraId: "11111111", text: "Alpha Beta", formatting: JOINED_FORMATTING },
] as const;

const SPLIT_BLOCKS = [
  { paraId: "22222222", text: "Alpha", formatting: FIRST_SPLIT_FORMATTING },
  { paraId: "33333333", text: "Beta", formatting: SECOND_SPLIT_FORMATTING },
] as const;

const paragraph = (
  paraId: string,
  text: string,
  { styleId, alignment, spacing }: DirectParagraphFormatting,
): Paragraph => ({
  type: "paragraph",
  paraId,
  textId: paraId,
  formatting: { styleId, alignment, ...spacing },
  content: [{ type: "run", content: [{ type: "text", text }] }],
});

const documentWith = async (
  blocks: readonly {
    paraId: string;
    text: string;
    formatting: DirectParagraphFormatting;
  }[],
): Promise<ArrayBuffer> => {
  const document = createEmptyDocument();
  document.package.styles = {
    styles: [
      { type: "paragraph", styleId: "Normal", name: "Normal", default: true },
      { type: "paragraph", styleId: JOINED_FORMATTING.styleId, name: "Joined Body" },
      { type: "paragraph", styleId: FIRST_SPLIT_FORMATTING.styleId, name: "Opening Body" },
      { type: "paragraph", styleId: SECOND_SPLIT_FORMATTING.styleId, name: "Closing Body" },
    ],
  };
  document.package.document.content = blocks.map(({ paraId, text, formatting }) =>
    paragraph(paraId, text, formatting),
  );
  return await createDocx(document);
};

const tableDocumentWith = async (
  blocks: readonly {
    paraId: string;
    text: string;
    formatting: DirectParagraphFormatting;
  }[],
): Promise<ArrayBuffer> => {
  const document = createEmptyDocument();
  document.package.styles = {
    styles: [
      { type: "paragraph", styleId: "Normal", name: "Normal", default: true },
      { type: "paragraph", styleId: JOINED_FORMATTING.styleId, name: "Joined Body" },
      { type: "paragraph", styleId: FIRST_SPLIT_FORMATTING.styleId, name: "Opening Body" },
      { type: "paragraph", styleId: SECOND_SPLIT_FORMATTING.styleId, name: "Closing Body" },
    ],
  };
  const table: Table = {
    type: "table",
    rows: [
      {
        type: "tableRow",
        cells: [
          {
            type: "tableCell",
            content: blocks.map(({ paraId, text, formatting }) =>
              paragraph(paraId, text, formatting),
            ),
          },
        ],
      },
    ],
  };
  document.package.document.content = [
    table,
    { type: "paragraph", paraId: "44444444", textId: "44444444", content: [] },
  ];
  return await createDocx(document);
};

const joinedDocument = (): Promise<ArrayBuffer> => documentWith(JOINED_BLOCKS);

const splitDocument = (): Promise<ArrayBuffer> => documentWith(SPLIT_BLOCKS);

const projectedBlocks = (reviewer: FolioDocxReviewer) =>
  reviewer.snapshot().blocks.map(({ text, styleId, directAlignment, directSpacing }) => ({
    text,
    styleId,
    directAlignment,
    directSpacing,
  }));

const JOINED_PROJECTION = [
  {
    text: "Alpha Beta",
    styleId: JOINED_FORMATTING.styleId,
    directAlignment: JOINED_FORMATTING.alignment,
    directSpacing: JOINED_FORMATTING.spacing,
  },
];

const SPLIT_PROJECTION = [
  {
    text: "Alpha",
    styleId: FIRST_SPLIT_FORMATTING.styleId,
    directAlignment: FIRST_SPLIT_FORMATTING.alignment,
    directSpacing: FIRST_SPLIT_FORMATTING.spacing,
  },
  {
    text: "Beta",
    styleId: SECOND_SPLIT_FORMATTING.styleId,
    directAlignment: SECOND_SPLIT_FORMATTING.alignment,
    directSpacing: SECOND_SPLIT_FORMATTING.spacing,
  },
];

const EMPTY_CARRIER_PROJECTION = {
  text: "",
  styleId: undefined,
  directAlignment: undefined,
  directSpacing: undefined,
};

const mainDocumentXml = async (buffer: ArrayBuffer): Promise<string> => {
  const part = (await JSZip.loadAsync(buffer)).file("word/document.xml");
  if (!part) {
    return panic("expected word/document.xml");
  }
  return await part.async("string");
};

const expectReviewedProjection = async (
  buffer: ArrayBuffer,
  action: "accept" | "reject",
  expected: readonly ReturnType<typeof projectedBlocks>[number][],
): Promise<void> => {
  const reviewer = await FolioDocxReviewer.fromBuffer(buffer);
  const revisionCount = action === "accept" ? reviewer.acceptAll() : reviewer.rejectAll();
  expect(revisionCount).toBeGreaterThan(0);
  expect(reviewer.getChanges()).toEqual([]);
  expect(projectedBlocks(reviewer)).toEqual(expected);

  const saved = await reviewer.toBuffer();
  const savedXml = await mainDocumentXml(saved);
  expect(savedXml).not.toContain("<w:pPrChange");
  expect(savedXml).not.toContain("<w:ins");
  expect(savedXml).not.toContain("<w:del");
  const reopened = await FolioDocxReviewer.fromBuffer(saved);
  expect(reopened.getChanges()).toEqual([]);
  expect(projectedBlocks(reopened)).toEqual(expected);
};

describe("paragraph split and merge formatting", () => {
  test("direct split and merge operations apply their complete paragraph property payloads", async () => {
    const splitting = await FolioDocxReviewer.fromBuffer(await joinedDocument());
    const joinedBlockId = splitting.snapshot().blocks.at(0)?.id;
    if (!joinedBlockId) {
      panic("expected a joined paragraph block id");
    }
    const split = splitting.applyDocumentOperations({
      version: 1,
      mode: "direct",
      operations: [
        {
          id: "split",
          type: "splitBlock",
          blockId: joinedBlockId,
          offset: 5,
          separator: " ",
          firstParagraphProperties: {
            styleId: FIRST_SPLIT_FORMATTING.styleId,
            alignment: FIRST_SPLIT_FORMATTING.alignment,
            spacing: FIRST_SPLIT_FORMATTING.spacing,
          },
          secondParagraphProperties: {
            styleId: SECOND_SPLIT_FORMATTING.styleId,
            alignment: SECOND_SPLIT_FORMATTING.alignment,
            spacing: SECOND_SPLIT_FORMATTING.spacing,
          },
        },
      ],
    });
    expect(split.skipped).toEqual([]);
    expect(projectedBlocks(splitting)).toEqual(SPLIT_PROJECTION);

    const merging = await FolioDocxReviewer.fromBuffer(await splitDocument());
    const firstSplitBlockId = merging.snapshot().blocks.at(0)?.id;
    if (!firstSplitBlockId) {
      panic("expected a split paragraph block id");
    }
    const merge = merging.applyDocumentOperations({
      version: 1,
      mode: "direct",
      operations: [
        {
          id: "merge",
          type: "mergeBlockWithNext",
          blockId: firstSplitBlockId,
          separator: " ",
          mergedParagraphProperties: {
            styleId: JOINED_FORMATTING.styleId,
            alignment: JOINED_FORMATTING.alignment,
            spacing: JOINED_FORMATTING.spacing,
          },
        },
      ],
    });
    expect(merge.skipped).toEqual([]);
    expect(projectedBlocks(merging)).toEqual(JOINED_PROJECTION);
  });

  test("a split carries each resulting paragraph's style, alignment, and spacing", async () => {
    const result = await compareDocx(await joinedDocument(), await splitDocument(), OPTIONS);
    if (result.isErr()) {
      throw result.error;
    }

    expect(result.value.verification).toEqual({ status: "verified" });
    expect(result.value.unsupported).toEqual([]);
    expect(result.value.changes).toEqual([
      expect.objectContaining({ kind: "split" }),
      expect.objectContaining({
        kind: "paragraph-format",
        targetBlockId: "22222222",
        properties: {
          styleId: FIRST_SPLIT_FORMATTING.styleId,
          alignment: FIRST_SPLIT_FORMATTING.alignment,
          spacing: FIRST_SPLIT_FORMATTING.spacing,
        },
      }),
      expect.objectContaining({
        kind: "paragraph-format",
        targetBlockId: "33333333",
        properties: {
          styleId: SECOND_SPLIT_FORMATTING.styleId,
          alignment: SECOND_SPLIT_FORMATTING.alignment,
          spacing: SECOND_SPLIT_FORMATTING.spacing,
        },
      }),
    ]);
    expect((await mainDocumentXml(result.value.buffer)).match(/<w:pPrChange\b/gu)).toHaveLength(2);
    await expectReviewedProjection(result.value.buffer, "accept", SPLIT_PROJECTION);
    await expectReviewedProjection(result.value.buffer, "reject", JOINED_PROJECTION);
  });

  test("the reverse merge carries the joined paragraph's style, alignment, and spacing", async () => {
    const result = await compareDocx(await splitDocument(), await joinedDocument(), OPTIONS);
    if (result.isErr()) {
      throw result.error;
    }

    expect(result.value.verification).toEqual({ status: "verified" });
    expect(result.value.unsupported).toEqual([]);
    expect(result.value.changes).toEqual([
      expect.objectContaining({ kind: "merge" }),
      expect.objectContaining({
        kind: "paragraph-format",
        targetBlockId: "11111111",
        properties: {
          styleId: JOINED_FORMATTING.styleId,
          alignment: JOINED_FORMATTING.alignment,
          spacing: JOINED_FORMATTING.spacing,
        },
      }),
    ]);
    expect((await mainDocumentXml(result.value.buffer)).match(/<w:pPrChange\b/gu)).toHaveLength(1);
    await expectReviewedProjection(result.value.buffer, "accept", JOINED_PROJECTION);
    await expectReviewedProjection(result.value.buffer, "reject", SPLIT_PROJECTION);
  });

  test.each([
    {
      label: "split",
      baseBlocks: JOINED_BLOCKS,
      revisedBlocks: SPLIT_BLOCKS,
      accepted: [...SPLIT_PROJECTION, EMPTY_CARRIER_PROJECTION],
      rejected: [...JOINED_PROJECTION, EMPTY_CARRIER_PROJECTION],
    },
    {
      label: "merge",
      baseBlocks: SPLIT_BLOCKS,
      revisedBlocks: JOINED_BLOCKS,
      accepted: [...JOINED_PROJECTION, EMPTY_CARRIER_PROJECTION],
      rejected: [...SPLIT_PROJECTION, EMPTY_CARRIER_PROJECTION],
    },
  ] as const)(
    "round-trips a formatting $label inside a table-cell container",
    async ({ baseBlocks, revisedBlocks, accepted, rejected }) => {
      const result = await compareDocx(
        await tableDocumentWith(baseBlocks),
        await tableDocumentWith(revisedBlocks),
        OPTIONS,
      );
      if (result.isErr()) {
        throw result.error;
      }

      expect(result.value.verification).toEqual({ status: "verified" });
      expect(result.value.unsupported).toEqual([]);
      await expectReviewedProjection(result.value.buffer, "accept", accepted);
      await expectReviewedProjection(result.value.buffer, "reject", rejected);
    },
  );

  test("refuses a formatting split when the source already has a pending pPrChange", async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(await joinedDocument());
    const blockId = reviewer.snapshot().blocks.at(0)?.id;
    if (!blockId) {
      panic("expected a joined paragraph block id");
    }
    expect(
      reviewer.applyDocumentOperations({
        version: 1,
        mode: "tracked-changes",
        operations: [
          {
            id: "pending",
            type: "setBlockParagraphProperties",
            blockId,
            properties: { alignment: "both" },
          },
        ],
      }).skipped,
    ).toEqual([]);

    const outcome = reviewer.applyDocumentOperations({
      version: 1,
      mode: "tracked-changes",
      atomic: true,
      operations: [
        {
          id: "split",
          type: "splitBlock",
          blockId,
          offset: 5,
          separator: " ",
          firstParagraphProperties: { styleId: FIRST_SPLIT_FORMATTING.styleId },
          secondParagraphProperties: { styleId: SECOND_SPLIT_FORMATTING.styleId },
        },
      ],
    });

    expect(outcome.status).toBe("rejected");
    expect(outcome.applied).toEqual([]);
    expect(outcome.skipped).toEqual([{ id: "split", reason: "pendingParagraphPropertyChange" }]);
    expect(reviewer.snapshot().blocks.map(({ text }) => text)).toEqual(["Alpha Beta"]);
  });

  test("refuses a formatting merge when the first paragraph has a pending pPrChange", async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(await splitDocument());
    const blockId = reviewer.snapshot().blocks.at(0)?.id;
    if (!blockId) {
      panic("expected a split paragraph block id");
    }
    expect(
      reviewer.applyDocumentOperations({
        version: 1,
        mode: "tracked-changes",
        operations: [
          {
            id: "pending",
            type: "setBlockParagraphProperties",
            blockId,
            properties: { alignment: "both" },
          },
        ],
      }).skipped,
    ).toEqual([]);

    const outcome = reviewer.applyDocumentOperations({
      version: 1,
      mode: "tracked-changes",
      atomic: true,
      operations: [
        {
          id: "merge",
          type: "mergeBlockWithNext",
          blockId,
          separator: " ",
          mergedParagraphProperties: { styleId: JOINED_FORMATTING.styleId },
        },
      ],
    });

    expect(outcome.status).toBe("rejected");
    expect(outcome.applied).toEqual([]);
    expect(outcome.skipped).toEqual([{ id: "merge", reason: "pendingParagraphPropertyChange" }]);
    expect(reviewer.snapshot().blocks.map(({ text }) => text)).toEqual(["Alpha", "Beta"]);
  });
});
