/**
 * A direct word replacement in a table cell paragraph of a document that
 * holds tracked marks, hyperlinks, a note reference and preserved content
 * applies in the canonical reviewer: the edit publishes, undo gives every read
 * back, and the same replacement applies tracked.
 *
 * The document and the draw are a generated case of the canonical edit laws
 * in `headlessCanonicalEdits.property.test.ts`, kept as a fixed case.
 */

import { describe, expect, test } from "bun:test";
import { readFileSync } from "node:fs";
import path from "node:path";

import {
  bodyBlockIds,
  operationFor,
} from "../../../../test/generators/reviewerOperationDraws";
import { FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION } from "../document-operations";
import type { FolioDocumentOperation } from "../document-operations";
import { prepareCanonicalDocxInput } from "../docx/canonicalSessionInput";
import { createDocx } from "../docx/rezip";
import type { Document } from "../types/document";
import { FolioDocxReviewer } from "./headless";

const DOCUMENT = JSON.parse(
  readFileSync(
    path.join(
      import.meta.dir,
      "__fixtures__",
      "table-cell-replace-counterexample.json",
    ),
    "utf8",
  ),
) as Document;

const packaged = async (document: Document): Promise<ArrayBuffer> => {
  const prepared = await prepareCanonicalDocxInput(
    await createDocx(structuredClone(document)),
  );
  if (prepared.isErr()) throw prepared.error;
  return prepared.value.slice().buffer;
};

const open = (buffer: ArrayBuffer) =>
  FolioDocxReviewer.fromBuffer(buffer.slice(0), {
    author: "Canonical Reviewer",
    session: "canonical",
  });

const reads = (reviewer: FolioDocxReviewer) => ({
  content: reviewer.getContent(),
  text: reviewer.getContentAsText({ annotated: true }),
  notes: reviewer.getNotesAsText(),
  stories: reviewer.listStories(),
  changes: reviewer.getChanges(),
  comments: reviewer.getComments(),
  document: reviewer.toDocument(),
});

const apply = (
  reviewer: FolioDocxReviewer,
  operation: FolioDocumentOperation,
  mode: "direct" | "tracked-changes",
) =>
  reviewer.applyDocumentOperations(
    {
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      operations: [operation],
      mode,
    },
    { undefinedReferences: "refuse" },
  );

describe("a word replaced in a table cell paragraph", () => {
  test("applies directly, undoes, and applies tracked", async () => {
    const buffer = await packaged(DOCUMENT);
    const reviewer = await open(buffer);
    const operation = operationFor(
      reviewer,
      {
        kind: "replaceInBlock",
        block: 10,
        word: 1_882_668_977,
        text: "revised",
      },
      {
        replaceIn: "blocksWithWords",
        among: bodyBlockIds(reviewer.toDocument().package.document.content),
      },
    );
    if (operation === null) throw new Error("the draw names no block");
    expect(operation.type).toBe("replaceInBlock");

    const before = reads(reviewer);
    const direct = apply(reviewer, operation, "direct");
    expect(direct.skipped).toEqual([]);
    expect(direct.applied.map(({ id }) => id)).toEqual([operation.id]);
    expect(reviewer.getContentAsText()).toContain("revised");
    if (direct.undoHandle === null) throw new Error("no undo handle");
    expect(reviewer.undoDocumentOperations(direct.undoHandle)).toMatchObject({
      status: "undone",
    });
    expect(reads(reviewer)).toEqual(before);

    const tracked = apply(await open(buffer), operation, "tracked-changes");
    expect(tracked.skipped).toEqual([]);
    expect(tracked.applied.map(({ id }) => id)).toEqual([operation.id]);
  }, 60_000);
});
