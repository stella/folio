/** Zero-width `w:commentReference`: where a comment's visible mark is painted. */

import type { Node as PMNode } from "prosemirror-model";
import { Plugin } from "prosemirror-state";

import { readCommentReferenceAttrs } from "../../commentReferenceAttrs";
import {
  planCommentReferenceRepairs,
  type CommentMarkExtent,
  type CommentReferenceOccurrence,
} from "../../commentReferenceIntegrity";
import { createNodeExtension } from "../create";

type CommentReferenceOptions = {
  getInternalClipboardToken?: () => string;
};

/** The node name, for callers asking whether a paragraph holds any content. */
export const COMMENT_REFERENCE_NODE_NAME = "commentReference";

type CommentReferenceSurvey = {
  references: CommentReferenceOccurrence[];
  markExtents: CommentMarkExtent[];
};

const surveyCommentReferences = (doc: PMNode): CommentReferenceSurvey => {
  const references: CommentReferenceOccurrence[] = [];
  const markEnds = new Map<number, number>();

  doc.descendants((node, position) => {
    if (node.type.name === COMMENT_REFERENCE_NODE_NAME) {
      const attrs = readCommentReferenceAttrs(node);
      references.push({
        commentId: attrs.ok ? attrs.value.commentId : -1,
        position,
        size: node.nodeSize,
        malformed: !attrs.ok,
      });
      return false;
    }
    if (!node.isInline) {
      return true;
    }
    for (const mark of node.marks) {
      if (mark.type.name !== "comment") {
        continue;
      }
      const commentId = mark.attrs["commentId"];
      if (typeof commentId !== "number") {
        continue;
      }
      markEnds.set(commentId, Math.max(markEnds.get(commentId) ?? 0, position + node.nodeSize));
    }
    return true;
  });

  return {
    references,
    markExtents: [...markEnds].map(([commentId, end]) => ({ commentId, end })),
  };
};

export const CommentReferenceExtension = createNodeExtension<CommentReferenceOptions>({
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
          if (raw === null || !/^(?:0|[1-9]\d*)$/.test(raw)) {
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
          "data-docx-comment-reference": String(attrs.ok ? attrs.value.commentId : ""),
          "aria-hidden": "true",
          contenteditable: "false",
          style: "display: none;",
          ...(options.getInternalClipboardToken
            ? { "data-docx-internal-clipboard": options.getInternalClipboardToken() }
            : {}),
        },
      ];
    },
  }),
  onSchemaReady: ({ schema }) => ({
    plugins: [
      new Plugin({
        appendTransaction(transactions, _oldState, newState) {
          if (!transactions.some(({ docChanged }) => docChanged)) {
            return null;
          }
          const referenceType = schema.nodes[COMMENT_REFERENCE_NODE_NAME];
          if (!referenceType) {
            return null;
          }

          const repairs = planCommentReferenceRepairs(surveyCommentReferences(newState.doc));
          if (repairs.length === 0) {
            return null;
          }

          const transaction = newState.tr;
          for (const repair of repairs) {
            if (repair.type === "delete") {
              transaction.delete(repair.position, repair.position + repair.size);
              continue;
            }
            transaction.insert(
              transaction.mapping.map(repair.position),
              referenceType.create({ commentId: repair.commentId }),
            );
          }
          return transaction;
        },
      }),
    ],
  }),
});
