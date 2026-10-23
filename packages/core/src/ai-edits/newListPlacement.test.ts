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
import { FolioDocxReviewer } from "./headless";

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
type Mode = "direct" | "tracked-changes";

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

describe("new-list placement", () => {
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
});
