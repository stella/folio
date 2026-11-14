/**
 * Both editing engines refuse an operation naming a paragraph style or a
 * numbering instance the document does not define, for the same reason and
 * with the same message, and leave the document as it was.
 */

import { describe, expect, mock, test } from "bun:test";
import { readFileSync } from "node:fs";
import path from "node:path";

import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperation,
} from "../document-operations";
import { prepareCanonicalDocxInput } from "../docx/canonicalSessionInput";
import { FolioDocxReviewer } from "./headless";
import * as undefinedReferences from "./undefinedReferences";

const FIXTURE = path.join(
  import.meta.dir,
  "../docx/__tests__/__fixtures__/corpus/upstream-styled-content.docx",
);

/** The package as a host stores it, so both engines read the same block ids. */
const ingested = async (): Promise<ArrayBuffer> => {
  const bytes = readFileSync(FIXTURE);
  const prepared = await prepareCanonicalDocxInput(
    bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength),
  );
  if (prepared.isErr()) throw prepared.error;
  return prepared.value.slice().buffer;
};

const open = (buffer: ArrayBuffer, session: "legacy" | "canonical") =>
  FolioDocxReviewer.fromBuffer(buffer.slice(0), {
    author: "Reviewer",
    session,
  });

const firstParagraphId = (reviewer: FolioDocxReviewer): string => {
  const block = reviewer
    .getContent()
    .find((candidate) => candidate.kind !== "diagnostic" && candidate.text);
  if (block === undefined) throw new Error("the fixture has no text block");
  return block.id;
};

const CASES: {
  name: string;
  reason: "missingStyle" | "missingNumbering";
  operation: (blockId: string) => FolioDocumentOperation;
}[] = [
  {
    name: "an inserted paragraph's style",
    reason: "missingStyle",
    operation: (blockId) => ({
      id: "op-1",
      type: "insertAfterBlock",
      blockId,
      text: "x",
      styleId: "NoSuchStyle",
    }),
  },
  {
    name: "a replaced paragraph's style",
    reason: "missingStyle",
    operation: (blockId) => ({
      id: "op-1",
      type: "replaceBlock",
      blockId,
      text: "Replaced.",
      styleId: "NoSuchStyle",
    }),
  },
  {
    name: "a paragraph's stated style",
    reason: "missingStyle",
    operation: (blockId) => ({
      id: "op-1",
      type: "setBlockParagraphProperties",
      blockId,
      properties: { styleId: "NoSuchStyle" },
    }),
  },
  {
    name: "a split half's style",
    reason: "missingStyle",
    operation: (blockId) => ({
      id: "op-1",
      type: "splitBlock",
      blockId,
      offset: 1,
      secondParagraphProperties: { styleId: "NoSuchStyle" },
    }),
  },
  {
    name: "an inserted paragraph's numbering instance",
    reason: "missingNumbering",
    operation: (blockId) => ({
      id: "op-1",
      type: "insertBeforeBlock",
      blockId,
      text: "x",
      numbering: { kind: "reference", numId: 9999 },
    }),
  },
];

describe("undefined style and numbering references", () => {
  for (const { name, reason, operation } of CASES) {
    for (const mode of ["direct", "tracked-changes"] as const) {
      test(`${name} is refused alike by both engines (${mode})`, async () => {
        const buffer = await ingested();
        const outcomes = await Promise.all(
          (["legacy", "canonical"] as const).map(async (session) => {
            const reviewer = await open(buffer, session);
            const before = reviewer.getContent();
            const result = reviewer.applyDocumentOperations(
              {
                version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
                mode,
                operations: [operation(firstParagraphId(reviewer))],
              },
              { undefinedReferences: "refuse" },
            );
            expect(result.applied).toEqual([]);
            expect(reviewer.getContent()).toEqual(before);
            return result.skipped.map((skip) => ({
              id: skip.id,
              reason: skip.reason,
              message: skip.message,
            }));
          }),
        );
        expect(outcomes[0]?.[0]?.reason).toBe(reason);
        expect(outcomes[1]).toEqual(outcomes[0]);
      });
    }
  }
});

// The refusal must come from the shared check: with it turned off, the
// canonical engine applies the operation and the refusal above would fail.
test("turning the shared check off lets the canonical engine apply an undefined style", async () => {
  const original = { ...undefinedReferences };
  await mock.module("./undefinedReferences", () => ({
    ...original,
    findUndefinedOperationReference: () => undefined,
  }));
  try {
    const reviewer = await open(await ingested(), "canonical");
    const result = reviewer.applyDocumentOperations(
      {
        version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
        mode: "direct",
        operations: [
          {
            id: "op-1",
            type: "insertAfterBlock",
            blockId: firstParagraphId(reviewer),
            text: "x",
            styleId: "NoSuchStyle",
          },
        ],
      },
      { undefinedReferences: "refuse" },
    );
    expect(result.skipped).toEqual([]);
    expect(result.applied).toHaveLength(1);
  } finally {
    await mock.module("./undefinedReferences", () => original);
  }
});
