/**
 * A comment added on text another comment's range runs past puts its
 * reference inside that range: a range is one stretch, so a save writes the
 * reference between the outer range's start and end, and reopening reads it
 * there. The live document holds it there too, so removing the text after the
 * reference leaves the outer range ending after it, saved or not.
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
  document.package.document.content.slice(0, 1).flatMap((block) =>
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
const nested = async (reservedId?: number) => {
  const document = fromMarkdown(
    reservedId === undefined
      ? "Clause one tail."
      : "Clause one tail.\n\nOther paragraph.",
  );
  if (reservedId !== undefined) {
    const other = document.package.document.content.at(1);
    if (other?.type !== "paragraph") throw new Error("Missing allocator seed");
    other.content = [
      { type: "commentRangeStart", id: reservedId },
      ...other.content,
      { type: "commentRangeEnd", id: reservedId },
      { type: "run", content: [{ type: "commentReference", id: reservedId }] },
    ];
    document.package.document.comments = [
      {
        id: reservedId,
        author: "Test",
        date: "2026-01-01T00:00:00Z",
        anchorKind: "range",
        content: [
          {
            type: "paragraph",
            content: [
              {
                type: "run",
                content: [{ type: "text", text: "Already occupied" }],
              },
            ],
          },
        ],
      },
    ];
  }
  const reviewer = await open(
    (await ensureParaIds(await createDocx(document))).docx,
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

const commentIdsByText = (reviewer: FolioDocxReviewer) => {
  const document = reviewer.toDocument();
  return Object.fromEntries(
    (document.package.document.comments ?? []).map(({ id }) => [
      commentText(document, id),
      id,
    ]),
  );
};

const ID_ORDERS = [
  {
    name: "ascending IDs",
    reservedId: undefined,
    outerId: 1,
    innerId: 2,
    starts: ["Outer", "Inner"],
  },
  {
    name: "wrapped IDs",
    reservedId: MAX_REVISION_ID - 1,
    outerId: MAX_REVISION_ID,
    innerId: 1,
    starts: ["Inner", "Outer"],
  },
] as const;

describe("a comment reference put inside another comment's range", () => {
  for (const mode of MODES) {
    for (const { name, reservedId, outerId, innerId, starts } of ID_ORDERS) {
      test(`stays inside it after removing the tail (${mode}, ${name})`, async () => {
        const { reviewer, blockId } = await nested(reservedId);
        const commentIds =
          reservedId === undefined
            ? { Outer: outerId, Inner: innerId }
            : {
                Outer: outerId,
                Inner: innerId,
                "Already occupied": reservedId,
              };
        expect(commentIdsByText(reviewer)).toEqual(commentIds);
        const liveBeforeSave = markers(reviewer.toDocument());
        const reopened = await open(await reviewer.toBuffer());
        expect(markers(reviewer.toDocument())).toEqual(liveBeforeSave);
        expect(commentIdsByText(reopened)).toEqual(commentIds);
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
