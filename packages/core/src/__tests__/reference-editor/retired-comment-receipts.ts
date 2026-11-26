import type { Node as PMNode } from "prosemirror-model";
import { EditorState } from "prosemirror-state";

import { anchoredCommentIdsInProseDoc } from "../../ai-edits/comment-lifecycle";
import { addRetiredCommentReceipts } from "../../ai-edits/retired-comment-receipts";
import {
  createDocumentNumberingPlugin,
  getStatedDocumentNumbering,
} from "../../prosemirror/plugins/documentNumbering";
import { createDocumentStyleContextPlugin } from "../../prosemirror/plugins/documentStyles";
import { getDocumentStyleResolver } from "../../prosemirror/plugins/documentStyleState";
import { acceptAllChanges, rejectEverySuggestion } from "./comments";

const acceptedComments = (
  before: EditorState,
  doc: PMNode,
): ReadonlySet<number> | null => {
  const state = rejectEverySuggestion(
    EditorState.create({
      doc,
      plugins: [
        createDocumentNumberingPlugin(getStatedDocumentNumbering(before)),
        createDocumentStyleContextPlugin(getDocumentStyleResolver(before)),
      ],
    }),
  );
  let accepted = state.doc;
  const resolved = acceptAllChanges()(state, (transaction) => {
    accepted = transaction.doc;
  });
  return resolved ? anchoredCommentIdsInProseDoc(accepted) : null;
};

export const addReferenceRetiredCommentReceipts = (
  options: Omit<
    Parameters<typeof addRetiredCommentReceipts>[0],
    "acceptedCommentIds"
  >,
): void => {
  const { before, after } = options;
  addRetiredCommentReceipts({
    ...options,
    acceptedCommentIds: (side) =>
      acceptedComments(before, side === "before" ? before.doc : after),
  });
};
