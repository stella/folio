/**
 * A comment added on text another comment's range runs past puts its
 * reference inside that range: a range is one stretch, so a save writes the
 * reference between the outer range's start and end, and reopening reads it
 * there. The live document holds it there too, so removing the text after the
 * reference leaves the outer range ending after it, saved or not.
 *
 * The saved copy is the live reviewer's own save: comment ids are allocated
 * realm-wide, so a second reviewer built the same way can hold other ids, and
 * ranges opening at one position are written in id order.
 */

import { describe, expect, test } from "bun:test";

import { MAX_REVISION_ID } from "@stll/docx-core/model";

import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperation,
  type FolioDocumentOperationMode,
} from "../document-operations";
import { ensureParaIds } from "../docx/ensureParaIds";
import { createDocx } from "../docx/rezip";
import { fromMarkdown } from "../markdown/fromMarkdown";
import { seedCommentIdAbove } from "../prosemirror/commentIdAllocator";
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

/** "Outer" covers the paragraph; "Inner" covers its first word only. */
const nested = async () => {
  const reviewer = await open(
    (await ensureParaIds(await createDocx(fromMarkdown("Clause one tail."))))
      .docx,
  );
  const block = reviewer.snapshot().blocks[0];
  if (!block) throw new Error("Missing fixture");
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

describe("a comment reference put inside another comment's range", () => {
  for (const mode of MODES) {
    test(`stays inside it once the text after it goes, saved or not (${mode})`, async () => {
      const live = await nested();
      const reopened = await open(await live.reviewer.toBuffer());
      const expected = removeTail(reopened, mode, live.blockId);
      expect(removeTail(live.reviewer, mode, live.blockId)).toEqual(expected);
      expect(expected).toEqual([
        "commentRangeStart Outer",
        "commentRangeStart Inner",
        "Clause",
        "commentRangeEnd Inner",
        "commentReference Inner",
        "commentRangeEnd Outer",
        "commentReference Outer",
      ]);
    });
  }

  test("its saved copy reads the same when the comment ids wrap past the last one", async () => {
    // The outer comment takes the last id and the inner one wraps to a low id.
    seedCommentIdAbove(MAX_REVISION_ID - 1);
    const live = await nested();
    const reopened = await open(await live.reviewer.toBuffer());
    expect(removeTail(reopened, "direct", live.blockId)).toEqual(
      removeTail(live.reviewer, "direct", live.blockId),
    );
  });
});
