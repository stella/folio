/**
 * A batch that deletes a paragraph and puts blocks beside it: a member of a
 * gap replayed where it lands takes fresh ids when the ones it was admitted
 * with no longer fit, with or without a session lease; a table refused
 * before the paragraph after it keeps nothing apart, so the deletion takes
 * the paragraph whole; and a paragraph moved past the deletion copies the
 * paragraph it was put after.
 */

import { panic } from "better-result";
import { expect, test } from "bun:test";
import { EditorState } from "prosemirror-state";

import type { IdLease } from "@stll/docx-core/ops";

import { FolioDocxReviewer } from "../ai-edits/headless";
import {
  createFolioAIEditSnapshot,
  hashFolioAIBlockContent,
} from "../ai-edits/snapshot";
import type { FolioDocumentOperation } from "../document-operations";
import { ensureParaIds } from "../docx/ensureParaIds";
import { parseDocx } from "../docx/parser";
import { createDocx } from "../docx/rezip";
import { fromMarkdown } from "../markdown/fromMarkdown";
import { CanonicalPublicOperations } from "./canonicalPublicOperations";
import {
  type CanonicalCommit,
  createCanonicalSession,
  publishCanonicalProjection,
} from "./canonicalSession";

const TABLE = "| TBL a | TBL b |\n| --- | --- |\n| TBL c | TBL d |";

const bytesOf = async (markdown: string): Promise<ArrayBuffer> => {
  const { docx } = await ensureParaIds(
    await createDocx(fromMarkdown(markdown)),
  );
  return docx.slice().buffer;
};

const paragraphIdsOf = (reviewer: FolioDocxReviewer): string[] =>
  reviewer
    .snapshot()
    .blocks.filter(({ table }) => table === undefined)
    .map(({ id }) => id);

const LEASE: IdLease = {
  blocks: [
    { kind: "paragraph", start: 0x40000001, end: 0x40000041 },
    { kind: "revision", start: 10000, end: 10400 },
    { kind: "control", start: 20000, end: 20400 },
  ],
};

test.each(["unleased", "leased"] as const)(
  "a table after a table and a paragraph before the deleted last paragraph after it replay with their own revision ids (%s)",
  async (leasing) => {
    const earlier = await FolioDocxReviewer.fromBuffer(
      await bytesOf(`P0 clause.\n\nP1 clause.\n\n${TABLE}\n\nP2 clause.`),
      { author: "Earlier" },
    );
    const last = paragraphIdsOf(earlier).at(-1) ?? panic("No last paragraph");
    expect(
      earlier.applyOperations(
        [{ id: "final", type: "deleteBlock", blockId: last }],
        { mode: "tracked-changes", undefinedReferences: "refuse" },
      ).skipped,
    ).toEqual([]);
    const session = createCanonicalSession(
      await parseDocx(await earlier.toBuffer()),
      undefined,
      leasing === "leased" ? { idLease: LEASE } : {},
    ).unwrap();
    let state = EditorState.create({ doc: session.projection.doc });
    const publish = (commit: CanonicalCommit) => {
      const published = publishCanonicalProjection({ session, state, commit });
      if (published.isErr()) return false;
      state = published.value.state;
      return true;
    };
    const blocks = createFolioAIEditSnapshot(state.doc).blocks;
    const cell =
      blocks.find(({ table }) => table !== undefined)?.id ??
      panic("No table cell");
    const operations = new CanonicalPublicOperations({
      undoLimit: Number.POSITIVE_INFINITY,
      session,
      getState: () => state,
      publish,
    });
    const result = operations.apply({
      snapshot: createFolioAIEditSnapshot(state.doc),
      batch: {
        version: 1,
        mode: "tracked-changes",
        operations: [
          {
            id: "table",
            type: "insertTable",
            blockId: cell,
            position: "after",
            rows: [["Table term"]],
          },
          {
            id: "before",
            type: "insertBeforeBlock",
            blockId: last,
            text: "Before last.",
          },
          { id: "delete", type: "deleteBlock", blockId: last },
        ],
      },
      undefinedReferences: "refuse",
    });
    expect(result.skipped).toEqual([]);
    expect(
      createFolioAIEditSnapshot(state.doc).blocks.map(({ text }) => text),
    ).toEqual([
      "P0 clause.",
      "P1 clause.",
      "TBL a",
      "TBL b",
      "TBL c",
      "TBL d",
      "Table term",
      "Before last.",
    ]);
  },
);

test.each(["deletion first", "table first"] as const)(
  "a table refused before the next paragraph leaves the deleted paragraph nothing to keep apart (%s)",
  async (order) => {
    const reviewer = await FolioDocxReviewer.fromBuffer(
      await bytesOf(`${TABLE}\n\nP clause.\n\nQ clause.`),
      { author: "Reviewer" },
    );
    const [deleted, next] = paragraphIdsOf(reviewer);
    if (deleted === undefined || next === undefined)
      panic("Missing paragraphs");
    const operations: FolioDocumentOperation[] = [
      { id: "delete", type: "deleteBlock", blockId: deleted },
      {
        id: "table",
        type: "insertTable",
        blockId: next,
        position: "before",
        rows: [["Table term"]],
        precondition: { blockTextHash: hashFolioAIBlockContent("Stale.") },
      },
    ];
    const result = reviewer.applyOperations(
      order === "deletion first" ? operations : operations.toReversed(),
      { mode: "direct", undefinedReferences: "refuse" },
    );
    expect(result.skipped.map(({ id, reason }) => ({ id, reason }))).toEqual([
      { id: "table", reason: "preconditionFailed" },
    ]);
    expect(reviewer.snapshot().blocks.map(({ text }) => text)).toEqual([
      "TBL a",
      "TBL b",
      "TBL c",
      "TBL d",
      "Q clause.",
    ]);
  },
);

test("a paragraph moved past a deletion copies the paragraph it was put after", async () => {
  const reviewer = await FolioDocxReviewer.fromBuffer(
    await bytesOf(`## Heading A\n\nBody B.\n\n${TABLE}`),
    { author: "Reviewer" },
  );
  const [heading, body] = paragraphIdsOf(reviewer);
  if (heading === undefined || body === undefined) panic("Missing paragraphs");
  const result = reviewer.applyOperations(
    [
      { id: "x", type: "insertBeforeBlock", blockId: body, text: "X." },
      { id: "y", type: "insertAfterBlock", blockId: heading, text: "Y." },
      { id: "delete", type: "deleteBlock", blockId: body },
    ],
    { mode: "direct", undefinedReferences: "refuse" },
  );
  expect(result.skipped).toEqual([]);
  expect(
    reviewer
      .snapshot()
      .blocks.filter(({ table }) => table === undefined)
      .map(({ text, styleId }) => ({ text, styleId: styleId ?? null })),
  ).toEqual([
    { text: "Heading A", styleId: "Heading2" },
    { text: "X.", styleId: null },
    { text: "Y.", styleId: "Heading2" },
    { text: "", styleId: null },
  ]);
});
