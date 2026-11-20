/**
 * A comment added on text another comment's range runs past puts its
 * reference inside that range: a range is one stretch, so a save writes the
 * reference between the outer range's start and end, and reopening reads it
 * there. The live document holds it there too, so removing the text after the
 * reference leaves the outer range ending after it, saved or not.
 */

import { describe, expect, spyOn, test } from "bun:test";

import { MAX_REVISION_ID } from "@stll/docx-core/model";

import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperation,
  type FolioDocumentOperationMode,
} from "../document-operations";
import { ensureParaIds } from "../docx/ensureParaIds";
import { createDocx } from "../docx/rezip";
import { fromMarkdown } from "../markdown/fromMarkdown";
import * as commentIds from "../prosemirror/commentIdAllocator";
import type { Document } from "../types/document";
import { FolioDocxReviewer } from "./headless";
import { hashFolioAIBlockText } from "./snapshot";

const MODES = ["direct", "tracked-changes"] as const;

const open = async (bytes: ArrayBuffer | Uint8Array) =>
  FolioDocxReviewer.fromBuffer(bytes, { author: "Test" });

const apply = (
  reviewer: FolioDocxReviewer,
  mode: FolioDocumentOperationMode,
  operations: FolioDocumentOperation[],
) => {
  const result = reviewer.applyDocumentOperations(
    { version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION, mode, operations },
    { undefinedReferences: "refuse" },
  );
  expect(result.skipped).toEqual([]);
};

const commentText = (document: Document, id: number): string =>
  (document.package.document.comments ?? [])
    .find((comment) => comment.id === id)
    ?.content.flatMap((paragraph) => paragraph.content)
    .flatMap((child) => (child.type === "run" ? child.content : []))
    .map((item) => (item.type === "text" ? item.text : ""))
    .join("") ?? String(id);

/** The commented paragraph's words and comment markers, named by comment text. */
const markers = (document: Document) =>
  document.package.document.content.flatMap((block) =>
    block.type === "paragraph"
      ? block.content.flatMap((child) => {
          if (child.type === "run")
            return child.content.flatMap((item) =>
              item.type === "text" ? [item.text] : [],
            );
          if (
            child.type === "commentRangeStart" ||
            child.type === "commentRangeEnd" ||
            child.type === "commentReference"
          )
            return [`${child.type} ${commentText(document, child.id)}`];
          return [];
        })
      : [],
  );

type NestedCommentIds = { outerId: number; innerId: number };

/** "Outer" covers the paragraph; "Inner" covers its first word only. */
const nested = async ({ outerId, innerId }: NestedCommentIds) => {
  const reviewer = await open(
    (await ensureParaIds(await createDocx(fromMarkdown("Clause one tail."))))
      .docx,
  );
  const block = reviewer.snapshot().blocks[0];
  if (!block) throw new Error("Missing fixture");
  // Starts at the same position are ordered by ID. Pin the IDs while
  // adding both comments through public operations.
  const allocator = spyOn(commentIds, "allocateCommentId")
    .mockReturnValueOnce(outerId)
    .mockReturnValueOnce(innerId);
  try {
    apply(reviewer, "direct", [
      {
        id: "outer",
        type: "commentOnBlock",
        blockId: block.id,
        comment: { text: "Outer" },
      },
    ]);
    apply(reviewer, "direct", [
      {
        id: "inner",
        type: "commentOnRange",
        range: {
          type: "textRange",
          story: "main",
          blockId: block.id,
          startOffset: 0,
          endOffset: 6,
          selectedTextHash: hashFolioAIBlockText("Clause"),
        },
        comment: { text: "Inner" },
      },
    ]);
    expect(allocator).toHaveBeenCalledTimes(2);
  } finally {
    allocator.mockRestore();
  }
  return { reviewer, blockId: block.id };
};

const removeTail = (
  reviewer: FolioDocxReviewer,
  mode: FolioDocumentOperationMode,
  blockId: string,
) => {
  apply(reviewer, mode, [
    {
      id: "tail",
      type: "replaceInBlock",
      blockId,
      find: " one tail.",
      replace: "",
    },
  ]);
  return markers(
    mode === "direct"
      ? reviewer.toDocument()
      : reviewer.toResolvedDocument("accept"),
  );
};

const ID_ORDERS = [
  { name: "ascending IDs", outerId: 1, innerId: 2, starts: ["Outer", "Inner"] },
  {
    name: "wrapped IDs",
    outerId: MAX_REVISION_ID,
    innerId: 1,
    starts: ["Inner", "Outer"],
  },
] as const;

describe("a comment reference put inside another comment's range", () => {
  for (const mode of MODES) {
    for (const { name, outerId, innerId, starts } of ID_ORDERS) {
      test(`stays inside it after removing the tail (${mode}, ${name})`, async () => {
        const { reviewer, blockId } = await nested({ outerId, innerId });
        const liveBeforeSave = markers(reviewer.toDocument());
        const reopened = await open(await reviewer.toBuffer());
        expect(markers(reviewer.toDocument())).toEqual(liveBeforeSave);
        const expected = [
          ...starts.map((comment) => `commentRangeStart ${comment}`),
          "Clause",
          "commentRangeEnd Inner",
          "commentReference Inner",
          "commentRangeEnd Outer",
          "commentReference Outer",
        ];
        expect(removeTail(reviewer, mode, blockId)).toEqual(expected);
        expect(removeTail(reopened, mode, blockId)).toEqual(expected);
      });
    }
  }
});
