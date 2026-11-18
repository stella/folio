import type {
  Document,
  Paragraph,
} from "../../packages/docx-core/src/model/document";
import { applyDocumentOp } from "../../packages/docx-core/src/ops/apply";
import {
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  OP_STORIES,
} from "../../packages/docx-core/src/ops/types";

const paragraph = (paraId: string, text: string): Paragraph => ({
  type: "paragraph",
  paraId,
  content: [{ type: "run", content: [{ type: "text", text }] }],
});
type NoteSourceOptions = {
  kind: "footnote" | "endnote";
  anchor: "point" | "range";
  text: string;
  depth: number;
  references: "sole" | "shared";
};
export const noteCommentDocument = ({
  kind,
  anchor,
  text,
  depth,
  references,
}: NoteSourceOptions) => {
  let document: Document = {
    package: {
      document: {
        content: [
          paragraph("10000001", text),
          paragraph("10000002", "standing"),
        ],
        // Parsed roots without source anchors have no anchorKind and remain preserved.
        comments: [
          {
            id: 99,
            author: "Loaded",
            content: [paragraph("20000063", "unanchored source definition")],
          },
        ],
      },
    },
  };
  const otherKind = kind === "footnote" ? "endnote" : "footnote";
  if (otherKind === "footnote")
    document.package.footnotes = [
      {
        type: otherKind,
        id: 3,
        content: [paragraph("30000003", "untouched note")],
      },
    ];
  else
    document.package.endnotes = [
      {
        type: otherKind,
        id: 3,
        content: [paragraph("30000003", "untouched note")],
      },
    ];
  const apply = (op: DocumentOp) => {
    document = applyDocumentOp(document, op).unwrap().document;
  };
  apply({
    type: DOCUMENT_OP_TYPES.CREATE_COMMENT,
    comment: {
      id: 11,
      author: "Reviewer",
      content: [paragraph("2000000B", "unrelated main comment")],
    },
    anchor: {
      kind: "point",
      at: { story: OP_STORIES.MAIN, blockId: "10000002", offset: 1 },
    },
  });
  apply({
    type: DOCUMENT_OP_TYPES.CREATE_COMMENT,
    comment: {
      id: 13,
      author: "Reviewer",
      content: [paragraph("2000000D", "untouched note comment")],
    },
    anchor: {
      kind: "point",
      at: { story: { kind: otherKind, id: 3 }, blockId: "30000003", offset: 1 },
    },
  });
  apply({
    type: DOCUMENT_OP_TYPES.ADD_NOTE,
    at: { story: OP_STORIES.MAIN, blockId: "10000001", offset: text.length },
    note: { type: kind, id: 1, content: [paragraph("20000001", text)] },
  });
  const at = { story: { kind, id: 1 }, blockId: "20000001", offset: 1 };
  apply({
    type: DOCUMENT_OP_TYPES.CREATE_COMMENT,
    comment: {
      id: 7,
      author: "Reviewer",
      done: false,
      content: [paragraph("20000007", text)],
    },
    anchor:
      anchor === "point"
        ? { kind: "point", at }
        : { kind: "range", from: at, to: { ...at, offset: text.length } },
  });
  const retiredIds = [7];
  for (let index = 0; index < depth; index++) {
    const id = 8 + index;
    apply({
      type: DOCUMENT_OP_TYPES.CREATE_COMMENT,
      comment: {
        id,
        author: "Reviewer",
        content: [
          paragraph((0x20000000 + id).toString(16).toUpperCase(), "reply"),
        ],
      },
      anchor: { kind: "reply", parentId: id - 1 },
    });
    retiredIds.push(id);
  }
  if (references === "shared")
    apply({
      type: DOCUMENT_OP_TYPES.INSERT_CONTENT,
      at: { story: OP_STORIES.MAIN, blockId: "10000002", offset: 0 },
      slice: {
        content: [
          {
            type: "run",
            content: [
              {
                type: kind === "footnote" ? "footnoteRef" : "endnoteRef",
                id: 1,
              },
            ],
          },
        ],
        openStart: 0,
        openEnd: 0,
      },
    });
  return { document, retiredIds, otherKind };
};
