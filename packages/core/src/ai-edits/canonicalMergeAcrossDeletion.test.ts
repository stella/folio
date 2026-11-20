import { afterAll, beforeAll, expect, setSystemTime, test } from "bun:test";

import { paragraphsDocx, textRun } from "../__tests__/operationBatchDocuments";
import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperation,
} from "../document-operations";
import { FolioDocxReviewer } from "./headless";

beforeAll(() => setSystemTime(new Date("2026-01-01T12:00:00Z")));
afterAll(() => setSystemTime());

const TEXTS = ["Alpha.", "Beta.", "Gamma.", "Delta."];
const idOf = (index: number) =>
  (0x1000_0000 + index).toString(16).toUpperCase();

type Outcome = { applied: string[]; texts: string[] };

const run = async (
  mode: "direct" | "tracked-changes",
  operations: FolioDocumentOperation[],
): Promise<Outcome> => {
  const reviewer = await FolioDocxReviewer.fromBuffer(
    await paragraphsDocx(TEXTS.map((text) => [textRun(text)])),
    { session: "canonical" },
  );
  const result = reviewer.applyDocumentOperations(
    { version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION, mode, operations },
    { undefinedReferences: "refuse" },
  );
  reviewer.acceptAll();
  return {
    applied: result.applied.map(({ id }) => id).toSorted(),
    texts: reviewer.getContent().map(({ text }) => text),
  };
};

const batch = (merged: number, deleted: readonly number[]) => [
  ...deleted.map((index): FolioDocumentOperation => ({
    id: `delete-${String(index)}`,
    type: "deleteBlock",
    blockId: idOf(index),
  })),
  {
    id: "merge",
    type: "mergeBlockWithNext",
    blockId: idOf(merged),
    separator: " ",
  } satisfies FolioDocumentOperation,
];

// A merge into a paragraph its batch deletes joins across it to the paragraph
// that then follows, directly as once its tracked counterpart is accepted.
test("a merge joins across the paragraphs its batch deletes", async () => {
  for (const deleted of [[1], [1, 2]]) {
    const direct = await run("direct", batch(0, deleted));
    expect(direct.applied).toHaveLength(deleted.length + 1);
    expect(direct.texts).toEqual(
      await run("tracked-changes", batch(0, deleted)).then(
        ({ texts }) => texts,
      ),
    );
    expect(direct.texts.at(0)).toBe(`Alpha. ${TEXTS[deleted.length + 1]}`);
  }
});

// With the deletions running to the end, nothing is left to join: directly the
// merge is refused, tracked it adds no separator, and both read alike.
test("a merge has nothing to join once its batch deletes every paragraph after it", async () => {
  const direct = await run("direct", batch(2, [3]));
  expect(direct.applied).toEqual(["delete-3"]);
  const tracked = await run("tracked-changes", batch(2, [3]));
  expect(tracked.texts).toEqual(direct.texts);
  expect(direct.texts).toEqual(["Alpha.", "Beta.", "Gamma."]);
});

// An edit of the first paragraph that changes its length runs before the
// direct merge that joins it away; the separator still goes at its end.
test("a direct merge puts its separator at the end of the paragraph its batch edits", async () => {
  const mergeAndEdit = (replace: string): FolioDocumentOperation[] => [
    {
      id: "merge",
      type: "mergeBlockWithNext",
      blockId: idOf(0),
      separator: " ",
    },
    {
      id: "replace",
      type: "replaceInBlock",
      blockId: idOf(0),
      find: "Alpha",
      replace,
    },
  ];
  for (const replace of ["Alphabet", "A"]) {
    const { texts } = await run("direct", mergeAndEdit(replace));
    expect(texts.at(0)).toBe(`${replace}. Beta.`);
    expect(texts).toEqual(
      (await run("tracked-changes", mergeAndEdit(replace))).texts,
    );
  }
});
