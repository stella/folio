import { expect, test } from "bun:test";
import { EditorState } from "prosemirror-state";

import {
  applyDocumentOp,
  applyDocumentOps,
  combineEdits,
} from "@stll/docx-core/ops";

import { rejectAIEditRevision } from "../prosemirror/commands/comments";
import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import { schema } from "../prosemirror/schema";
import type { Document } from "../types/document";

test("resolving a nested revision promotes its comment boundary ownership to the parent", () => {
  const document: Document = {
    package: {
      document: {
        comments: [{ id: 7, author: "Reviewer", content: [] }],
        content: [
          {
            type: "paragraph",
            paraId: "00000001",
            pPrMark: { kind: "del", info: { id: 3, author: "Reviewer" } },
            content: [
              {
                type: "insertion",
                info: { id: 1, author: "Reviewer" },
                content: [
                  {
                    type: "deletion",
                    info: { id: 2, author: "Reviewer" },
                    content: [
                      { type: "commentRangeStart", id: 7 },
                      {
                        type: "run",
                        content: [{ type: "text", text: "comment" }],
                      },
                    ],
                  },
                  { type: "commentRangeEnd", id: 7 },
                  { type: "commentReference", id: 7 },
                ],
              },
            ],
          },
          {
            type: "paragraph",
            paraId: "00000002",
            content: [
              { type: "run", content: [{ type: "text", text: "after" }] },
            ],
          },
        ],
      },
    },
  };
  let state = EditorState.create({ schema, doc: toProseDoc(document) });
  expect(
    rejectAIEditRevision(2)(state, (transaction) => {
      state = state.apply(transaction);
    }),
  ).toBe(true);

  const retained = state.doc.firstChild?.firstChild?.marks.find(
    (mark) => mark.type.name === "insertion",
  );
  expect(retained?.attrs["_docxCommentRangeBoundaries"]).toEqual([
    { type: "start", commentId: 7, owner: "parent" },
    { type: "end", commentId: 7, owner: "parent" },
  ]);
  expect(() =>
    fromProseDoc(state.doc, document, {
      stylesheetSource: { type: "package" },
    }),
  ).not.toThrow();
  const saved = fromProseDoc(state.doc, document, {
    stylesheetSource: { type: "package" },
  });
  expect(saved.package.document.comments?.map(({ id }) => id)).toEqual([7]);
});

test("deleting a whole comment range retires its definition before composing the inverse", () => {
  const document: Document = {
    package: {
      document: {
        comments: [
          {
            id: 7,
            author: "Reviewer",
            anchorKind: "range",
            content: [
              {
                type: "paragraph",
                paraId: "00000003",
                content: [
                  {
                    type: "run",
                    content: [{ type: "text", text: "Comment body" }],
                  },
                ],
              },
            ],
          },
        ],
        content: [
          {
            type: "paragraph",
            paraId: "00000001",
            content: [
              { type: "commentRangeStart", id: 7 },
              {
                type: "run",
                content: [{ type: "text", text: "Commented text" }],
              },
              { type: "commentRangeEnd", id: 7 },
              { type: "commentReference", id: 7 },
            ],
          },
          {
            type: "paragraph",
            paraId: "00000002",
            content: [
              { type: "run", content: [{ type: "text", text: "After" }] },
            ],
          },
        ],
      },
    },
  };
  const deletion = applyDocumentOp(document, {
    type: "deleteBlocks",
    story: "main",
    blockIds: ["00000001"],
    newIds: { revision: [], control: [] },
  }).unwrap();
  expect(
    deletion.document.package.document.comments?.find(({ id }) => id === 7),
  ).toBeUndefined();
  const composed = combineEdits(document, [deletion]);

  expect(composed.document.package.document.comments).toBeUndefined();
  const restored = applyDocumentOps(
    composed.document,
    composed.inverse,
  ).unwrap().document;
  expect(restored.package.document).toStrictEqual(document.package.document);
});
