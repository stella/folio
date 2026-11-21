/**
 * A batch that deletes the paragraph a container ends with, and places a
 * block in the gap before it, reads what precedes that paragraph from the
 * finished batch. A paragraph placed there ends the container and the deleted
 * one goes; a table placed there leaves the deleted one emptied, so the
 * container still ends with a paragraph. The batch runs backwards, so its
 * deletion runs before the insertion: the outcome must not depend on that,
 * on the order the operations are stated in, or on the mode.
 */

import { describe, expect, test } from "bun:test";

import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperation,
} from "../document-operations";
import { ensureParaIds } from "../docx/ensureParaIds";
import { createDocx } from "../docx/rezip";
import { fromMarkdown } from "../markdown/fromMarkdown";
import type { HeadlessReviewerSession } from "../types/canonicalCapabilities";
import { FolioDocxReviewer } from "./headless";

const SESSIONS = [
  "legacy",
  "canonical",
] as const satisfies readonly HeadlessReviewerSession[];

type Reviewer = Awaited<ReturnType<typeof FolioDocxReviewer.fromBuffer>>;
type Mode = "direct" | "tracked-changes";

const CLOSING = "Closing remarks.";

const FIXTURES = {
  "a paragraph precedes it": {
    markdown: `# Terms\n\nBody text.\n\n${CLOSING}`,
    before: "Body text.",
  },
  "a table precedes it": {
    markdown: `# Terms\n\n| Item | Price |\n| --- | --- |\n| Widget | 10 |\n\n${CLOSING}`,
    before: "10",
  },
} as const;

type Fixture = keyof typeof FIXTURES;

const sources = new Map<Fixture, Uint8Array>();
const sourceOf = async (fixture: Fixture) => {
  const known = sources.get(fixture);
  if (known) return known;
  const bytes = (
    await ensureParaIds(
      await createDocx(fromMarkdown(FIXTURES[fixture].markdown)),
    )
  ).docx;
  sources.set(fixture, bytes);
  return bytes;
};

const opener = (session: HeadlessReviewerSession) => (bytes: Uint8Array) =>
  FolioDocxReviewer.fromBuffer(
    bytes.buffer.slice(
      bytes.byteOffset,
      bytes.byteOffset + bytes.byteLength,
    ) as ArrayBuffer,
    { author: "Test", session },
  );

type Open = ReturnType<typeof opener>;

const idOf = (reviewer: Reviewer, text: string) => {
  const block = reviewer.snapshot().blocks.find((row) => row.text === text);
  if (!block) throw new Error(`No block reads ${text}`);
  return block.id;
};

/** Each way to place a block in the gap before the closing paragraph. */
const GAP_INSERTIONS: Record<
  string,
  {
    fixtures: readonly Fixture[];
    operation: (ids: {
      before: string;
      closing: string;
    }) => FolioDocumentOperation;
  }
> = {
  "a paragraph after the block before it": {
    fixtures: ["a paragraph precedes it", "a table precedes it"],
    operation: ({ before }) => ({
      id: "insert",
      type: "insertAfterBlock",
      blockId: before,
      text: "Added clause.",
      styleId: "Heading2",
    }),
  },
  "a paragraph before it": {
    fixtures: ["a paragraph precedes it", "a table precedes it"],
    operation: ({ closing }) => ({
      id: "insert",
      type: "insertBeforeBlock",
      blockId: closing,
      text: "Added clause.",
      styleId: "Heading2",
    }),
  },
  "a table after the paragraph before it": {
    fixtures: ["a paragraph precedes it"],
    operation: ({ before }) => ({
      id: "insert",
      type: "insertTable",
      blockId: before,
      rows: [
        ["Term", "Value"],
        ["Period", "12 months"],
      ],
    }),
  },
};

const apply = (
  reviewer: Reviewer,
  mode: Mode,
  operations: FolioDocumentOperation[],
) => {
  const result = reviewer.applyDocumentOperations(
    { version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION, mode, operations },
    { undefinedReferences: "refuse" },
  );
  expect(result.skipped).toEqual([]);
};

type Row = { kind: string; text: string; inTable: boolean; styleId?: string };

/** The body a reader is left with once every change is accepted. */
const settled = async (open: Open, reviewer: Reviewer): Promise<Row[]> => {
  const saved = await open(new Uint8Array(await reviewer.toBuffer()));
  saved.acceptAll();
  const reopened = await open(new Uint8Array(await saved.toBuffer()));
  return reopened.getContent().map((block) => {
    const row: Row = {
      kind: block.kind,
      text: block.text,
      inTable: "table" in block && block.table !== undefined,
    };
    if ("styleId" in block && block.styleId !== undefined) {
      row.styleId = block.styleId;
    }
    return row;
  });
};

const rejected = async (open: Open, reviewer: Reviewer) => {
  const saved = await open(new Uint8Array(await reviewer.toBuffer()));
  saved.rejectAll();
  return saved.snapshot().blocks.map(({ text }) => text);
};

const operationsFor = (
  reviewer: Reviewer,
  fixture: Fixture,
  insertion: (typeof GAP_INSERTIONS)[string],
) => {
  const ids = {
    before: idOf(reviewer, FIXTURES[fixture].before),
    closing: idOf(reviewer, CLOSING),
  };
  const deletion: FolioDocumentOperation = {
    id: "delete",
    type: "deleteBlock",
    blockId: ids.closing,
  };
  return { insert: insertion.operation(ids), deletion };
};

describe.each(SESSIONS)(
  "a batch deleting a container's last paragraph beside a gap insertion (%s)",
  (session) => {
    const open = opener(session);
    for (const [name, insertion] of Object.entries(GAP_INSERTIONS)) {
      for (const fixture of insertion.fixtures) {
        test(`${name}, where ${fixture}`, async () => {
          const source = await sourceOf(fixture);
          const original = (await open(source))
            .snapshot()
            .blocks.map(({ text }) => text);
          const outcomes: Row[][] = [];
          for (const mode of ["direct", "tracked-changes"] as const) {
            // The batch, in both stated orders.
            for (const deletionFirst of [false, true]) {
              const reviewer = await open(source);
              const { insert, deletion } = operationsFor(
                reviewer,
                fixture,
                insertion,
              );
              apply(
                reviewer,
                mode,
                deletionFirst ? [deletion, insert] : [insert, deletion],
              );
              if (mode === "tracked-changes") {
                expect(await rejected(open, reviewer)).toEqual(original);
              }
              outcomes.push(await settled(open, reviewer));
            }
            // One at a time, with the insertion settled before the deletion.
            const reviewer = await open(source);
            const { insert, deletion } = operationsFor(
              reviewer,
              fixture,
              insertion,
            );
            apply(reviewer, mode, [insert]);
            apply(reviewer, mode, [deletion]);
            outcomes.push(await settled(open, reviewer));
          }
          const [first, ...rest] = outcomes;
          for (const outcome of rest) expect(outcome).toEqual(first);

          // The container still ends with a paragraph, and an emptied one only
          // where nothing but a table precedes it.
          const blocks = first ?? [];
          const last = blocks.at(-1);
          const beforeLast = blocks.at(-2);
          expect(last?.inTable).toBe(false);
          expect(blocks.some(({ text }) => text === CLOSING)).toBe(false);
          expect(last?.text !== "" || beforeLast?.inTable === true).toBe(true);
        });
      }
    }
  },
);

/*
 * A batch deleting a container's last paragraphs and placing a block in the
 * gap before them: tracked, in every stated order and one at a time, accepting
 * leaves what applying the operations directly does, and rejecting restores
 * the original.
 */
const CLOSINGS = ["First closing.", "Second closing.", "Third closing."];
const RUN_SOURCE = fromMarkdown(
  ["# Terms", "Body text.", ...CLOSINGS].join("\n\n"),
);

/** Each way to place a block in or before the deleted run. */
const RUN_GAP_INSERTIONS: Record<
  string,
  (ids: { before: string; first: string }) => FolioDocumentOperation
> = {
  "a paragraph after the paragraph before them": ({ before }) => ({
    id: "insert",
    type: "insertAfterBlock",
    blockId: before,
    text: "Added clause.",
    styleId: "Heading2",
  }),
  "a paragraph before the first of them": ({ first }) => ({
    id: "insert",
    type: "insertBeforeBlock",
    blockId: first,
    text: "Added clause.",
    styleId: "Heading2",
  }),
  "a paragraph after the first of them": ({ first }) => ({
    id: "insert",
    type: "insertAfterBlock",
    blockId: first,
    text: "Added clause.",
    styleId: "Heading2",
  }),
  "a table after the paragraph before them": ({ before }) => ({
    id: "insert",
    type: "insertTable",
    blockId: before,
    rows: [["Term", "Value"]],
  }),
};

describe.each(SESSIONS)(
  "a batch deleting a container's last paragraphs beside a gap insertion (%s)",
  (session) => {
    const open = opener(session);
    for (const count of [2, 3])
      for (const [name, insertion] of Object.entries(RUN_GAP_INSERTIONS))
        test(`${name}, deleting the last ${count}`, async () => {
          const source = (await ensureParaIds(await createDocx(RUN_SOURCE)))
            .docx;
          const original = (await open(source))
            .snapshot()
            .blocks.map(({ text }) => text);
          const operationsOf = (reviewer: Reviewer) => {
            const run = CLOSINGS.slice(-count);
            const deletions = run.map(
              (text, index): FolioDocumentOperation => ({
                id: `delete-${index}`,
                type: "deleteBlock",
                blockId: idOf(reviewer, text),
              }),
            );
            const before = CLOSINGS.at(-count - 1) ?? "Body text.";
            const insert = insertion({
              before: idOf(reviewer, before),
              first: idOf(reviewer, run[0] ?? ""),
            });
            return { deletions, insert };
          };
          // Applied directly one at a time, the insertion first.
          const direct = await open(source);
          const directOperations = operationsOf(direct);
          for (const operation of [
            directOperations.insert,
            ...directOperations.deletions,
          ])
            apply(direct, "direct", [operation]);
          const expected = await settled(open, direct);
          expect(expected.at(-1)?.inTable).toBe(false);
          // Tracked, as a batch in each stated order and one at a time.
          const orders: Record<
            string,
            (
              operations: ReturnType<typeof operationsOf>,
            ) => FolioDocumentOperation[][]
          > = {
            "insertion first": ({ insert, deletions }) => [
              [insert, ...deletions],
            ],
            "insertion last": ({ insert, deletions }) => [
              [...deletions, insert],
            ],
            "deletions backwards": ({ insert, deletions }) => [
              [...deletions.toReversed(), insert],
            ],
            "one at a time": ({ insert, deletions }) =>
              [insert, ...deletions].map((operation) => [operation]),
          };
          for (const batches of Object.values(orders)) {
            const reviewer = await open(source);
            for (const batch of batches(operationsOf(reviewer)))
              apply(reviewer, "tracked-changes", batch);
            expect(await rejected(open, reviewer)).toEqual(original);
            expect(await settled(open, reviewer)).toEqual(expected);
          }
        });
  },
);
