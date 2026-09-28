/**
 * An inserted paragraph's outline level.
 *
 * An insertion takes the anchor's paragraph properties unless it states its
 * own, and a heading written from Markdown states `w:outlineLvl` on the
 * paragraph. Without an `outlineLevel` of its own, a body paragraph inserted
 * beside that heading had no way to say it is not one.
 */

import { describe, expect, test } from "bun:test";

import { createDocx } from "../docx/rezip";
import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  parseFolioDocumentOperationBatch,
  type FolioDocumentOperationMode,
} from "../document-operations";
import { fromMarkdown } from "../markdown/fromMarkdown";
import { FolioDocxReviewer } from "./headless";

const HEADING = "Agreement";
const BODY = "First clause.";
const INSERTED = "Inserted clause.";

const MODES = [
  "direct",
  "tracked-changes",
] as const satisfies readonly FolioDocumentOperationMode[];

type InsertCase = {
  label: string;
  anchor: string;
  type: "insertBeforeBlock" | "insertAfterBlock";
  outlineLevel: unknown;
  expected: { kind: string; headingLevel?: number };
};

const CASES: readonly InsertCase[] = [
  {
    label: "null before a heading that states its level",
    anchor: HEADING,
    type: "insertBeforeBlock",
    outlineLevel: null,
    expected: { kind: "paragraph" },
  },
  {
    label: "body text before a heading that states its level",
    anchor: HEADING,
    type: "insertBeforeBlock",
    outlineLevel: { kind: "bodyText" },
    expected: { kind: "paragraph" },
  },
  {
    label: "a heading level after body text",
    anchor: BODY,
    type: "insertAfterBlock",
    outlineLevel: { kind: "heading", level: 2 },
    expected: { kind: "heading", headingLevel: 3 },
  },
];

const insertedAfterReopen = async (
  { anchor, type, outlineLevel }: InsertCase,
  mode: FolioDocumentOperationMode,
) => {
  const reviewer = await FolioDocxReviewer.fromBuffer(
    await createDocx(fromMarkdown(`# ${HEADING}\n\n${BODY}`)),
    { author: "AI" },
  );
  const blockId = reviewer.getContent().find((block) => block.text === anchor)?.id;
  if (blockId === undefined) throw new Error(`fixture must expose "${anchor}"`);
  // Through the wire parser, so the field is part of the accepted contract.
  const batch = parseFolioDocumentOperationBatch({
    version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
    mode,
    operations: [{ id: "op", type, blockId, text: INSERTED, styleId: null, outlineLevel }],
  });
  const result = reviewer.applyDocumentOperations(batch);
  expect(result.skipped).toEqual([]);
  if (mode !== "direct") reviewer.acceptAll();
  const reopened = await FolioDocxReviewer.fromBuffer(await reviewer.toBuffer());
  return reopened.getContent().find((block) => block.text === INSERTED);
};

describe("an insertion states its outline level", () => {
  test.each(CASES.flatMap((entry) => MODES.map((mode) => ({ entry, mode, label: entry.label }))))(
    "$label, $mode",
    async ({ entry, mode }) => {
      const inserted = await insertedAfterReopen(entry, mode);
      expect(inserted).toMatchObject(entry.expected);
      if (entry.expected.headingLevel === undefined) {
        expect(inserted?.headingLevel).toBeUndefined();
      }
    },
  );

  test("an invalid outline level is refused by the parser", () => {
    expect(() =>
      parseFolioDocumentOperationBatch({
        version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
        mode: "direct",
        operations: [
          {
            id: "op",
            type: "insertAfterBlock",
            blockId: "block",
            text: INSERTED,
            outlineLevel: { kind: "heading", level: 9 },
          },
        ],
      }),
    ).toThrow("outlineLevel");
  });
});
