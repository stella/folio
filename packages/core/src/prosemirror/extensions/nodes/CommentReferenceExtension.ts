/** Zero-width `w:commentReference`: where a comment's visible mark is painted. */

import { Plugin, TextSelection } from "prosemirror-state";

import { readCommentReferenceAttrs } from "../../commentReferenceAttrs";
import {
  COMMENT_REFERENCE_NODE_NAME,
  resolveCommentReferences,
} from "../../commentReferenceIntegrity";
import { createNodeExtension } from "../create";
import type { NodeExtension } from "../types";
import { INLINE_ANCHOR_DOM_STYLE } from "./inlineAnchorDom";

export { COMMENT_REFERENCE_NODE_NAME };

type CommentReferenceOptions = {
  getInternalClipboardToken?: () => string;
};

export const CommentReferenceExtension: (
  options?: Partial<CommentReferenceOptions> | undefined,
) => NodeExtension = createNodeExtension<CommentReferenceOptions>({
  name: COMMENT_REFERENCE_NODE_NAME,
  schemaNodeName: COMMENT_REFERENCE_NODE_NAME,
  nodeSpec: (options) => ({
    inline: true,
    group: "inline",
    marks: "_",
    atom: true,
    selectable: false,
    attrs: {
      commentId: {},
    },
    parseDOM: [
      {
        tag: "span[data-docx-comment-reference]",
        getAttrs(dom) {
          const raw = dom.getAttribute("data-docx-comment-reference");
          if (raw === null || !/^-?(?:0|[1-9]\d*)$/.test(raw)) {
            return false;
          }
          const commentId = Number(raw);
          return Number.isSafeInteger(commentId) ? { commentId } : false;
        },
      },
    ],
    toDOM(node) {
      const attrs = readCommentReferenceAttrs(node);
      return [
        "span",
        {
          "data-docx-comment-reference": String(
            attrs.ok ? attrs.value.commentId : "",
          ),
          "aria-hidden": "true",
          contenteditable: "false",
          style: INLINE_ANCHOR_DOM_STYLE,
          ...(options.getInternalClipboardToken
            ? {
                "data-docx-internal-clipboard":
                  options.getInternalClipboardToken(),
              }
            : {}),
        },
      ];
    },
  }),
  onSchemaReady: () => ({
    plugins: [
      new Plugin({
        appendTransaction(transactions, oldState, newState) {
          if (!transactions.some(({ docChanged }) => docChanged)) {
            return null;
          }
          const resolution = resolveCommentReferences(
            newState.doc,
            oldState.doc,
          );
          if (!resolution) return null;
          const transaction = newState.tr;
          for (const step of resolution.steps) transaction.step(step);
          // A reference put back at the end of its comment's text stays after
          // the caret there, so what is typed next still joins the comment.
          const { selection } = newState;
          if (selection instanceof TextSelection && selection.empty) {
            transaction.setSelection(
              TextSelection.create(
                transaction.doc,
                transaction.mapping.map(selection.head, -1),
              ),
            );
          }
          return transaction;
        },
      }),
    ],
  }),
});
