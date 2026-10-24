/**
 * Whether an inserted or merged paragraph becomes a list item, and in which
 * list, must not depend on how the request is applied: directly, tracked and
 * then accepted, or one operation at a time.
 */

import { describe, expect, test } from "bun:test";

import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperation,
} from "../document-operations";
import { ensureParaIds } from "../docx/ensureParaIds";
import { createDocx } from "../docx/rezip";
import { fromMarkdown } from "../markdown";
import type { Footnote, Paragraph } from "../types/document";
import { createEmptyDocument } from "../utils/createDocument";
import {
  FolioDocxReviewer,
  type FolioEditableDocumentStoryHandle,
} from "./headless";
import { createFolioAITextRangeHandle } from "./snapshot";

const AUTHOR = "Reviewer";

const LIST_MARKDOWN = [
  "# Delivery Terms",
  "The following apply to every order.",
  "- Goods are packed securely\n- Goods are insured in transit",
  "Payment happens in stages.",
  "1. Deposit on signature\n2. Balance on delivery\n3. Retention after inspection",
  "Closing remarks.",
].join("\n\n");

const listDocx = async (): Promise<ArrayBuffer> => {
  const bytes = (
    await ensureParaIds(
      new Uint8Array(await createDocx(fromMarkdown(LIST_MARKDOWN))),
    )
  ).docx;
  return bytes.buffer.slice(
    bytes.byteOffset,
    bytes.byteOffset + bytes.byteLength,
  ) as ArrayBuffer;
};

const open = (buffer: ArrayBuffer) =>
  FolioDocxReviewer.fromBuffer(buffer, { author: AUTHOR });

type Reviewer = Awaited<ReturnType<typeof open>>;
type Mode = "direct" | "tracked-changes" | "suggested";

const idOf = (reviewer: Reviewer, text: string): string => {
  const block = reviewer.getContent().find((row) => row.text === text);
  if (!block) throw new Error(`No block reads "${text}"`);
  return block.id;
};

const apply = (
  reviewer: Reviewer,
  mode: Mode,
  operations: FolioDocumentOperation[],
) =>
  reviewer.applyDocumentOperations(
    {
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode,
      operations,
    },
    { undefinedReferences: "refuse" },
  );

/** What a reader sees of each list paragraph, with list ids made relative. */
const lists = (reviewer: Reviewer) => {
  const ids = new Map<number, number>();
  return reviewer
    .getContent()
    .filter((row) => row.kind !== "diagnostic")
    .map((row) => {
      const numId =
        row.statedNumbering?.kind === "reference"
          ? row.statedNumbering.numId
          : undefined;
      if (numId !== undefined && !ids.has(numId)) ids.set(numId, ids.size);
      return {
        text: row.text,
        kind: row.kind,
        label: row.displayLabel,
        list: numId === undefined ? undefined : ids.get(numId),
      };
    });
};

/** The fixture with a whole list paragraph pending deletion. */
const withPendingDeletion = async (): Promise<ArrayBuffer> => {
  const reviewer = await open(await listDocx());
  const result = apply(reviewer, "tracked-changes", [
    {
      id: "pending",
      type: "deleteBlock",
      blockId: idOf(reviewer, "Goods are insured in transit"),
    },
  ]);
  expect(result.skipped).toEqual([]);
  return reviewer.toBuffer();
};

const anchorInheritanceDocx = async (): Promise<ArrayBuffer> => {
  type ParagraphOptions = { text: string; paraId: string };
  const paragraph = ({ text, paraId }: ParagraphOptions) =>
    ({
      type: "paragraph",
      paraId,
      content: [{ type: "run", content: [{ type: "text", text }] }],
    }) satisfies Paragraph;
  const document = createEmptyDocument();
  document.package.document.content = [
    paragraph({ text: "Anchor sentence.", paraId: "61000001" }),
    paragraph({ text: "Unrelated body.", paraId: "61000002" }),
    {
      ...paragraph({ text: "Note reference.", paraId: "61000003" }),
      content: [
        {
          type: "run",
          content: [
            { type: "text", text: "Note reference." },
            { type: "footnoteRef", id: 1 },
          ],
        },
      ],
    },
  ];
  document.package.footnotes = [
    {
      type: "footnote",
      id: 1,
      noteType: "normal",
      content: [
        paragraph({ text: "Anchor sentence.", paraId: "62000001" }),
        paragraph({ text: "Unrelated note.", paraId: "62000002" }),
      ],
    } satisfies Footnote,
  ];
  const bytes = (
    await ensureParaIds(new Uint8Array(await createDocx(document)))
  ).docx;
  return Uint8Array.from(bytes).buffer;
};

type ApplyToStoryOptions = {
  reviewer: Reviewer;
  story: FolioEditableDocumentStoryHandle;
  operations: FolioDocumentOperation[];
};

const applyToStory = ({ reviewer, story, operations }: ApplyToStoryOptions) =>
  reviewer.applyDocumentOperationsToStory({
    undefinedReferences: "refuse",
    story,
    batch: {
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode: "tracked-changes",
      operations,
    },
  });

const anchorInheritanceProjection = (
  reviewer: Reviewer,
  story: FolioEditableDocumentStoryHandle,
) => {
  const snapshot = reviewer.snapshotStory(story);
  if (!snapshot) throw new Error(`missing story ${JSON.stringify(story)}`);
  return snapshot.blocks
    .filter(({ text }) =>
      [
        "Inserted paragraph.",
        "Anchor sentence.",
        "Unrelated body.",
        "Unrelated note.",
      ].includes(text),
    )
    .map(({ text, kind, statedNumbering, listReference, displayLabel }) => ({
      text,
      kind,
      statedNumbering,
      listReference,
      displayLabel,
    }));
};

describe("new-list placement", () => {
  for (const { label, story } of [
    { label: "main", story: { type: "main" } as const },
    { label: "footnote", story: { type: "footnote", noteId: 1 } as const },
  ]) {
    for (const inheritFormatting of [true, false]) {
      test(`${label} insertion matches ordered new-list properties (inheritFormatting: ${inheritFormatting})`, async () => {
        const buffer = await anchorInheritanceDocx();
        const batchReviewer = await open(buffer);
        const sequentialReviewer = await open(buffer);
        const target = batchReviewer
          .snapshotStory(story)
          ?.blocks.find(({ text }) => text === "Anchor sentence.");
        if (!target) throw new Error(`missing anchor in ${label}`);
        const range = createFolioAITextRangeHandle({
          blockId: target.id,
          text: target.text,
          startOffset: 0,
          endOffset: 6,
        });
        if (!range) throw new Error(`missing range in ${label}`);
        const insert = {
          id: "insert-before-anchor",
          type: "insertBeforeBlock",
          blockId: target.id,
          text: "Inserted paragraph.",
          inheritFormatting,
        } satisfies FolioDocumentOperation;
        const format = {
          id: "format-anchor",
          type: "formatRange",
          range,
          formatting: { bold: true },
        } satisfies FolioDocumentOperation;
        const number = {
          id: "number-anchor",
          type: "setBlockParagraphProperties",
          blockId: target.id,
          properties: {
            numbering: { kind: "newList", format: "bullet" },
          },
        } satisfies FolioDocumentOperation;
        const before = anchorInheritanceProjection(batchReviewer, story);

        const batchResult = applyToStory({
          reviewer: batchReviewer,
          story,
          operations: [insert, format, number],
        });
        expect(batchResult.skipped).toEqual([]);
        for (const operation of [format, number, insert]) {
          expect(
            applyToStory({
              reviewer: sequentialReviewer,
              story,
              operations: [operation],
            }).skipped,
          ).toEqual([]);
        }

        const batchProjection = anchorInheritanceProjection(
          batchReviewer,
          story,
        );
        const sequentialProjection = anchorInheritanceProjection(
          sequentialReviewer,
          story,
        );
        expect(batchProjection).toEqual(sequentialProjection);
        const numberedAnchor = batchProjection.find(
          ({ text }) => text === "Anchor sentence.",
        );
        expect(numberedAnchor?.listReference).toBeDefined();
        const inserted = batchProjection.find(
          ({ text }) => text === "Inserted paragraph.",
        );
        if (!inserted) throw new Error(`missing insertion in ${label}`);
        expect(inserted.listReference !== undefined).toBe(inheritFormatting);
        if (inheritFormatting) {
          expect(inserted.listReference).toEqual(numberedAnchor?.listReference);
        }
        const unrelatedText =
          story.type === "main" ? "Unrelated body." : "Unrelated note.";
        expect(
          batchProjection.find(({ text }) => text === unrelatedText),
        ).toEqual(before.find(({ text }) => text === unrelatedText));
      });
    }
  }

  test("a direct merge into a pending-deleted list paragraph keeps the requested list once accepted", async () => {
    const source = await withPendingDeletion();
    const outcomes = await Promise.all(
      (["direct", "tracked-changes"] as const).map(async (mode) => {
        const reviewer = await open(source);
        const result = apply(reviewer, mode, [
          {
            id: "merge",
            type: "mergeBlockWithNext",
            blockId: idOf(reviewer, "Goods are packed securely"),
            separator: " ",
            mergedParagraphProperties: {
              numbering: { kind: "newList", format: "numbered" },
            },
          },
        ]);
        expect(result.skipped).toEqual([]);
        const saved = await open(await reviewer.toBuffer());
        saved.acceptAll();
        reviewer.acceptAll();
        // Accepting the live reviewer reads as accepting the saved package.
        expect(lists(reviewer)).toEqual(lists(saved));
        return lists(saved);
      }),
    );
    const [direct, tracked] = outcomes;
    expect(direct).toEqual(tracked);
    expect(
      direct?.find(({ text }) => text.startsWith("Goods are packed securely")),
    ).toMatchObject({ kind: "listItem", label: "1." });
  });

  for (const mode of ["direct", "tracked-changes"] as const) {
    test(`a merge that sets properties claims them against another property edit (${mode})`, async () => {
      const reviewer = await open(await listDocx());
      const heading = idOf(reviewer, "Delivery Terms");
      const result = apply(reviewer, mode, [
        {
          id: "bullets",
          type: "setBlockParagraphProperties",
          blockId: heading,
          properties: { numbering: { kind: "newList", format: "bullet" } },
        },
        {
          id: "merge",
          type: "mergeBlockWithNext",
          blockId: heading,
          separator: " ",
          mergedParagraphProperties: { numbering: { kind: "none" } },
        },
      ]);
      expect(result.applied.map(({ id }) => id)).toEqual(["bullets"]);
      expect(result.skipped).toEqual([
        expect.objectContaining({
          id: "merge",
          reason: "overlappingOperation",
        }),
      ]);
    });
  }

  test("a tracked merge that sets properties claims the joined paragraph's", async () => {
    const reviewer = await open(await listDocx());
    const result = apply(reviewer, "tracked-changes", [
      {
        id: "merge",
        type: "mergeBlockWithNext",
        blockId: idOf(reviewer, "The following apply to every order."),
        separator: " ",
        mergedParagraphProperties: {
          numbering: { kind: "newList", format: "numbered" },
        },
      },
      {
        id: "next",
        type: "setBlockParagraphProperties",
        blockId: idOf(reviewer, "Goods are packed securely"),
        properties: { numbering: { kind: "none" } },
      },
    ]);
    expect(result.skipped).toEqual([
      expect.objectContaining({ id: "next", reason: "overlappingOperation" }),
    ]);
  });

  test("a tracked new list inserted before deleted closing paragraphs is that list once accepted", async () => {
    const source = await withPendingDeletion();
    const operations = (reviewer: Reviewer): FolioDocumentOperation[] => [
      {
        id: "retention",
        type: "deleteBlock",
        blockId: idOf(reviewer, "Retention after inspection"),
      },
      {
        id: "insert",
        type: "insertAfterBlock",
        blockId: idOf(reviewer, "Balance on delivery"),
        text: "Inserted bullet.",
        numbering: { kind: "newList", format: "bullet" },
      },
      {
        id: "closing",
        type: "deleteBlock",
        blockId: idOf(reviewer, "Closing remarks."),
      },
    ];
    const tracked = await open(source);
    expect(
      apply(tracked, "tracked-changes", operations(tracked)).skipped,
    ).toEqual([]);
    const direct = await open(source);
    expect(apply(direct, "direct", operations(direct)).skipped).toEqual([]);

    const rejected = await open(await tracked.toBuffer());
    tracked.acceptAll();
    direct.acceptAll();
    expect(lists(tracked)).toEqual(lists(direct));
    expect(lists(tracked).at(-1)).toEqual({
      text: "Inserted bullet.",
      kind: "listItem",
      label: "•",
      list: 2,
    });

    rejected.rejectAll();
    const original = await open(source);
    original.rejectAll();
    expect(lists(rejected)).toEqual(lists(original));
  });

  test("a new list inserted after a table keeps its own break when the batch empties the closing paragraph", async () => {
    const bytes = (
      await ensureParaIds(
        new Uint8Array(
          await createDocx(
            fromMarkdown(
              [
                "# Schedule",
                "| Milestone | Date |\n| --- | --- |\n| Kick-off | May |",
                "Schedules may change by agreement.",
              ].join("\n\n"),
            ),
          ),
        ),
      )
    ).docx;
    const source = bytes.buffer.slice(
      bytes.byteOffset,
      bytes.byteOffset + bytes.byteLength,
    ) as ArrayBuffer;
    const operations = (reviewer: Reviewer) => {
      const insert: FolioDocumentOperation = {
        id: "insert",
        type: "insertAfterBlock",
        blockId: idOf(reviewer, "May"),
        text: "Delivery follows.",
        numbering: { kind: "newList", format: "numbered" },
      };
      const remove: FolioDocumentOperation = {
        id: "remove",
        type: "deleteBlock",
        blockId: idOf(reviewer, "Schedules may change by agreement."),
      };
      return { insert, remove };
    };
    const accepted = async (reviewer: Reviewer) => {
      const saved = await open(await reviewer.toBuffer());
      saved.acceptAll();
      return lists(saved);
    };

    const batch = await open(source);
    const { insert, remove } = operations(batch);
    expect(apply(batch, "tracked-changes", [insert, remove]).skipped).toEqual(
      [],
    );
    // One at a time, from the end backwards: the deletion runs first.
    const oneByOne = await open(source);
    const single = operations(oneByOne);
    expect(apply(oneByOne, "tracked-changes", [single.remove]).skipped).toEqual(
      [],
    );
    expect(apply(oneByOne, "tracked-changes", [single.insert]).skipped).toEqual(
      [],
    );
    const direct = await open(source);
    const directOperations = operations(direct);
    expect(
      apply(direct, "direct", [
        directOperations.insert,
        directOperations.remove,
      ]).skipped,
    ).toEqual([]);

    expect(lists(batch)).toEqual(lists(oneByOne));
    expect(await accepted(batch)).toEqual(await accepted(direct));
    expect(
      (await accepted(batch)).find(({ text }) => text === "Delivery follows."),
    ).toMatchObject({ kind: "listItem", label: "1." });
  });

  test("a suggested new list stays its own suggestion when a later batch reuses its operation id", async () => {
    const reviewer = await open(await listDocx());
    const first = apply(reviewer, "suggested", [
      {
        id: "op-1",
        type: "insertAfterBlock",
        blockId: idOf(reviewer, "Closing remarks."),
        text: "A new numbered item.",
        numbering: { kind: "newList", format: "numbered" },
      },
    ]);
    const second = apply(reviewer, "suggested", [
      {
        id: "op-1",
        type: "insertBeforeBlock",
        blockId: idOf(reviewer, "Payment happens in stages."),
        text: "An unnumbered note.",
      },
      {
        id: "op-2",
        suggestionId: "op-1",
        type: "replaceInBlock",
        blockId: idOf(reviewer, "Closing remarks."),
        find: "Closing",
        replace: "Final",
      },
    ]);
    const firstId = first.applied.at(0)?.suggestionId;
    const secondId = second.applied.find(
      ({ id }) => id === "op-1",
    )?.suggestionId;
    expect(firstId).toBe("op-1");
    expect(secondId).toBeDefined();
    expect(secondId).not.toBe(firstId);
    // An operation of the batch grouping with it by that id follows it.
    expect(second.applied.find(({ id }) => id === "op-2")?.suggestionId).toBe(
      secondId,
    );

    // Both stay pending, so the list the first one defines stays exported.
    expect(
      reviewer
        .exportPendingSuggestions()
        .map(({ suggestionId, operation }) => [suggestionId, operation.type]),
    ).toEqual(
      expect.arrayContaining([
        ["op-1", "insertAfterBlock"],
        [secondId, "insertBeforeBlock"],
      ]),
    );
    expect(reviewer.rejectSuggestion(secondId ?? "")).toBe(true);
    expect(
      reviewer.getContent().find(({ text }) => text === "A new numbered item."),
    ).toMatchObject({ kind: "listItem", displayLabel: "1." });
  });
});
