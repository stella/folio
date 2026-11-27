/**
 * The intents and operations of one edit with the content they place
 * re-homed into the part it lands in (`storyRelationshipRehoming`): a move
 * from one story into another, and content its caller read from another
 * story than the one it is placed into.
 */

import {
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  type EditorIntent,
} from "@stll/docx-core/ops";

import {
  COMMENTS_PART,
  type RelationshipRehomer,
  type RelationshipStory,
} from "../docx/storyRelationshipRehoming";

/**
 * `intent` with the content it places re-homed: a move's from the story it
 * moves out of, and any other placement's from `contentStory`, the story its
 * content was read from (the one it is placed into when absent).
 */
export const rehomedIntent = (
  rehomer: RelationshipRehomer,
  intent: EditorIntent,
  contentStory: RelationshipStory | undefined,
): EditorIntent => {
  const from = (to: RelationshipStory) => ({ from: contentStory ?? to, to });
  switch (intent.type) {
    case "moveFragment":
      return {
        ...intent,
        paragraphs: rehomer.rehome(intent.paragraphs, {
          from: intent.from.story,
          to: intent.target.story,
        }),
      };
    case "moveBlocks":
      return {
        ...intent,
        blocks: rehomer.rehome(intent.blocks, {
          from: intent.from.story,
          to: intent.target.story,
        }),
      };
    case "replaceFragment":
      return {
        ...intent,
        paragraphs: rehomer.rehome(intent.paragraphs, from(intent.from.story)),
      };
    case "pasteBlocks":
      return {
        ...intent,
        blocks: rehomer.rehome(intent.blocks, from(intent.from.story)),
      };
    case "insertParagraphs":
      return {
        ...intent,
        paragraphs: rehomer.rehome(intent.paragraphs, from(intent.story)),
      };
    case "tableBlock":
      return intent.operation.type === DOCUMENT_OP_TYPES.INSERT_TABLE
        ? {
            ...intent,
            operation: {
              ...intent.operation,
              table: rehomer.rehome(
                intent.operation.table,
                from(intent.operation.story),
              ),
            },
          }
        : intent;
    case "table":
      return intent.cellContent === undefined
        ? intent
        : {
            ...intent,
            cellContent: rehomer.rehome(
              intent.cellContent,
              from(intent.operation.story),
            ),
          };
    default:
      return intent;
  }
};

/**
 * `op` with the content it places re-homed from `contentStory`, the story
 * every relationship id in that content was read from. What an operation
 * expects to find stays as the story it acts on holds it.
 */
export const rehomedOp = (
  rehomer: RelationshipRehomer,
  op: DocumentOp,
  contentStory: RelationshipStory,
): DocumentOp => {
  const placement = (to: RelationshipStory) => ({ from: contentStory, to });
  switch (op.type) {
    case DOCUMENT_OP_TYPES.INSERT_CONTENT:
      return {
        ...op,
        slice: {
          ...op.slice,
          content: rehomer.rehome(op.slice.content, placement(op.at.story)),
        },
      };
    case DOCUMENT_OP_TYPES.INSERT_BLOCKS:
      return {
        ...op,
        blocks: rehomer.rehome(op.blocks, placement(op.story)),
      };
    case DOCUMENT_OP_TYPES.REPLACE_BLOCKS:
      return {
        ...op,
        blocks: rehomer.rehome(op.blocks, placement(op.story)),
      };
    case DOCUMENT_OP_TYPES.SET_CONTAINER_BLOCKS:
      return {
        ...op,
        blocks: rehomer.rehome(op.blocks, placement(op.story)),
      };
    case DOCUMENT_OP_TYPES.REPLACE_INLINE:
      return {
        ...op,
        content: rehomer.rehome(op.content, placement(op.story)),
      };
    case DOCUMENT_OP_TYPES.INSERT_TABLE:
    case DOCUMENT_OP_TYPES.SET_TABLE:
      return { ...op, table: rehomer.rehome(op.table, placement(op.story)) };
    case DOCUMENT_OP_TYPES.INSERT_ROW:
      return { ...op, row: rehomer.rehome(op.row, placement(op.story)) };
    case DOCUMENT_OP_TYPES.SET_TABLE_ROWS:
      return { ...op, rows: rehomer.rehome(op.rows, placement(op.story)) };
    case DOCUMENT_OP_TYPES.CREATE_HEADER_FOOTER:
      return {
        ...op,
        content: rehomer.rehome(op.content, placement(op.story)),
      };
    case DOCUMENT_OP_TYPES.ADD_NOTE:
      return {
        ...op,
        note: {
          ...op.note,
          content: rehomer.rehome(
            op.note.content,
            placement({ kind: op.note.type, id: op.note.id }),
          ),
        },
      };
    case DOCUMENT_OP_TYPES.CREATE_COMMENT:
      return {
        ...op,
        comment: {
          ...op.comment,
          content: rehomer.rehome(op.comment.content, placement(COMMENTS_PART)),
        },
      };
    case DOCUMENT_OP_TYPES.UPDATE_COMMENT_CONTENT:
      return {
        ...op,
        content: rehomer.rehome(op.content, placement(COMMENTS_PART)),
      };
    default:
      return op;
  }
};
