/**
 * `FolioDocxReviewer.fork()` on both engines: the copy reads as the
 * original, pending suggestions included; edits, suggestions and undo on
 * either leave the other as it is; the copy's undo history starts empty; and
 * a method a host set on the original stays with the original.
 */

import { describe, expect, test } from "bun:test";
import { readFileSync } from "node:fs";
import path from "node:path";

import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperationBatch,
} from "../document-operations";
import { prepareCanonicalDocxInput } from "../docx/canonicalSessionInput";
import type { HeadlessReviewerSession } from "../types/canonicalCapabilities";
import { FolioDocxReviewer } from "./headless";

const FIXTURE = path.join(
  import.meta.dir,
  "../docx/__tests__/__fixtures__/corpus/upstream-styled-content.docx",
);

const fixture = async (): Promise<ArrayBuffer> => {
  const bytes = readFileSync(FIXTURE);
  const prepared = await prepareCanonicalDocxInput(
    bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength),
  );
  if (prepared.isErr()) throw prepared.error;
  return prepared.value.slice().buffer;
};

const open = async (session: HeadlessReviewerSession) =>
  FolioDocxReviewer.fromBuffer(await fixture(), {
    author: "Fork Reviewer",
    session,
  });

/** What a reader sees, pending suggestions included. */
const reads = (reviewer: FolioDocxReviewer) => ({
  content: reviewer.getContent(),
  text: reviewer.getContentAsText({ annotated: true }),
  changes: reviewer.getChanges(),
  comments: reviewer.getComments(),
  pending: reviewer.exportPendingSuggestions(),
});

/** A replacement of the first word of the first block that has one. */
const replaceFirstWord = (
  reviewer: FolioDocxReviewer,
  mode: FolioDocumentOperationBatch["mode"],
  replace: string,
) => {
  const block = reviewer
    .getContent()
    .find((candidate) => /\S/u.test(candidate.text));
  const find = block?.text.split(/\s+/u).find((word) => word !== "");
  if (!block || find === undefined)
    throw new Error("The fixture has no word to replace.");
  return reviewer.applyDocumentOperations(
    {
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      operations: [
        { id: "op", type: "replaceInBlock", blockId: block.id, find, replace },
      ],
      mode,
    },
    { undefinedReferences: "refuse" },
  );
};

describe.each<HeadlessReviewerSession>(["legacy", "canonical"])(
  "fork() on the %s engine",
  (session) => {
    test("the copy reads as the original, a pending suggestion included", async () => {
      const original = await open(session);
      expect(
        replaceFirstWord(original, "suggested", "Proposed").applied,
      ).toHaveLength(1);
      expect(original.exportPendingSuggestions().length).toBeGreaterThan(0);
      expect(reads(original.fork())).toEqual(reads(original));
    });

    test("edits on either leave the other as it is", async () => {
      const original = await open(session);
      const before = reads(original);
      const copy = original.fork();
      expect(
        replaceFirstWord(copy, "tracked-changes", "Edited").applied,
      ).toHaveLength(1);
      expect(reads(original)).toEqual(before);
      const copied = reads(copy);
      expect(
        replaceFirstWord(original, "direct", "Changed").applied,
      ).toHaveLength(1);
      expect(reads(copy)).toEqual(copied);
    });

    test("the copy holds its own pending suggestions", async () => {
      const original = await open(session);
      const copy = original.fork();
      expect(
        replaceFirstWord(copy, "suggested", "Proposed").applied,
      ).toHaveLength(1);
      expect(copy.exportPendingSuggestions().length).toBeGreaterThan(0);
      expect(original.exportPendingSuggestions()).toEqual([]);
      const copied = copy.exportPendingSuggestions();
      expect(
        replaceFirstWord(original, "suggested", "Offered").applied,
      ).toHaveLength(1);
      expect(original.exportPendingSuggestions().length).toBeGreaterThan(0);
      expect(copy.exportPendingSuggestions()).toEqual(copied);
    });

    test("the copy's undo history starts empty; the original's stays", async () => {
      const original = await open(session);
      const before = reads(original);
      const { undoHandle } = replaceFirstWord(original, "direct", "Changed");
      if (undoHandle === null) throw new Error("The edit gave no undo handle.");
      const copy = original.fork();
      expect(copy.undoDocumentOperations(undoHandle)).toMatchObject({
        status: "rejected",
        reason: "unknownHandle",
      });
      expect(original.undoDocumentOperations(undoHandle)).toMatchObject({
        status: "undone",
      });
      expect(reads(original)).toEqual(before);
      expect(reads(copy)).not.toEqual(before);
    });

    test("a method a host set on the original stays with the original", async () => {
      const original = await open(session);
      const calls: string[] = [];
      const own = original.applyDocumentOperations.bind(original);
      Object.defineProperty(original, "applyDocumentOperations", {
        configurable: true,
        value: (
          ...args: Parameters<FolioDocxReviewer["applyDocumentOperations"]>
        ) => {
          calls.push("recorded");
          return own(...args);
        },
      });
      const copy = original.fork();
      expect(Object.hasOwn(copy, "applyDocumentOperations")).toBe(false);
      replaceFirstWord(copy, "direct", "Changed");
      expect(calls).toEqual([]);
    });
  },
);
