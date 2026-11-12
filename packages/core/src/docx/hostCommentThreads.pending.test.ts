/**
 * A comment a pending suggestion proposes is held back with the suggestion:
 * the saved package has neither the suggested text nor the comment's thread.
 * Accepting the suggestion saves both; rejecting it saves neither.
 */

import { describe, expect, test } from "bun:test";

import {
  createHarnessState,
  HeadlessEditorView,
} from "../__tests__/editorHarness";
import { createFolioAIEditSnapshot } from "../ai-edits/snapshot";
import {
  applyFolioDocumentOperations,
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
} from "../document-operations";
import {
  acceptAllSuggestions,
  rejectAllSuggestions,
} from "../prosemirror/commands/comments";
import {
  commentIdsNamedIn,
  commentRecordsOf,
  createCommentRecordsPlugin,
  recordingCommentsView,
} from "../prosemirror/commentRecords";
import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import type { Comment } from "../types/content";
import type { Document } from "../types/document";
import { withHostCommentThreads } from "./hostCommentThreads";

const PROPOSED = "Proposed note";

const base = (): Document => ({
  package: {
    document: {
      content: [
        {
          type: "paragraph",
          content: [
            { type: "run", content: [{ type: "text", text: "Clause" }] },
          ],
        },
      ],
      comments: [],
    },
  },
});

/** The saved package's comment texts, as a host's save writes them. */
const savedComments = (view: HeadlessEditorView, document: Document) => {
  const saved = withHostCommentThreads(
    fromProseDoc(view.state.doc, document, {
      stylesheetSource: { type: "package" },
    }),
    commentRecordsOf(view.state) ?? [],
    { anchoredInEditor: commentIdsNamedIn(view.state.doc) },
  );
  return (saved.package.document.comments ?? []).map(({ content }) =>
    JSON.stringify(content),
  );
};

const suggested = () => {
  const document = base();
  const view = new HeadlessEditorView(
    createHarnessState(document, "editing", [createCommentRecordsPlugin([])]),
  );
  const snapshot = createFolioAIEditSnapshot(view.state.doc);
  const block = snapshot.blocks.at(0);
  if (!block) throw new Error("The fixture has no block");
  const created: Comment[] = [];
  let nextId = 30;
  const result = applyFolioDocumentOperations({
    view: recordingCommentsView(view, created),
    snapshot,
    undefinedReferences: "refuse",
    author: "Assistant",
    batch: {
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode: "suggested",
      operations: [
        {
          id: "proposal",
          type: "insertAfterBlock",
          blockId: block.id,
          text: "A proposed clause.",
          comment: { text: PROPOSED },
        },
      ],
    },
    createCommentId: (text) => {
      nextId += 1;
      created.push({
        id: nextId,
        author: "Assistant",
        content: [
          {
            type: "paragraph",
            content: [{ type: "run", content: [{ type: "text", text }] }],
          },
        ],
      });
      return nextId;
    },
  });
  expect(result.status).toBe("committed");
  return { document, view };
};

describe("a comment a pending suggestion proposes", () => {
  test("is not saved while the suggestion is pending", () => {
    const { document, view } = suggested();
    expect(commentRecordsOf(view.state)?.length).toBe(1);
    expect(savedComments(view, document).join()).not.toContain(PROPOSED);
  });

  test("is saved once the suggestion is accepted", () => {
    const { document, view } = suggested();
    view.execute(acceptAllSuggestions({ author: "Adopter" }));
    expect(savedComments(view, document).join()).toContain(PROPOSED);
  });

  test("is gone once the suggestion is rejected", () => {
    const { document, view } = suggested();
    view.execute(rejectAllSuggestions());
    expect(commentRecordsOf(view.state)).toEqual([]);
    expect(savedComments(view, document)).toEqual([]);
  });
});
