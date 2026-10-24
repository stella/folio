/** Zero-width `w:commentReference`: where a comment's visible mark is painted. */

import { Plugin } from "prosemirror-state";

import { readCommentReferenceAttrs } from "../../commentReferenceAttrs";
import {
  COMMENT_REFERENCE_NODE_NAME,
  resolveCommentReferences,
} from "../../commentReferenceIntegrity";
import { createNodeExtension } from "../create";

export { COMMENT_REFERENCE_NODE_NAME };

type CommentReferenceOptions = {
  getInternalClipboardToken?: () => string;
};

export const CommentReferenceExtension =
  createNodeExtension<CommentReferenceOptions>({
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
            style: "display: none;",
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
          appendTransaction(transactions, _oldState, newState) {
            if (!transactions.some(({ docChanged }) => docChanged)) {
              return null;
            }
            const resolution = resolveCommentReferences(newState.doc);
            if (!resolution) return null;
            const transaction = newState.tr;
            for (const step of resolution.steps) transaction.step(step);
            return transaction;
          },
        }),
      ],
    }),
  });
