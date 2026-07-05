/**
 * Comment + tracked-change (proposeChange) PM transaction builders shared by
 * the React and Vue adapters. They locate the target range, apply the mark(s),
 * dispatch, and return a result. ID allocation lives in `commentIdAllocator.ts`
 * and is injected. Adapter-specific state mutation, event emission, and sidebar
 * UI stay in each adapter.
 */

import type { EditorView } from "prosemirror-view";
import type { Comment } from "../types/content";
import type { CommentIdAllocator } from "./commentIdAllocator";
import { findParagraphByParaId } from "./utils/findParagraphByParaId";
import { findTextInPmParagraph } from "./paraText";

/** Build a Comment object with a freshly-allocated ID. */
export function createComment(
  allocator: CommentIdAllocator,
  text: string,
  authorName: string,
  parentId?: number,
): Comment {
  return {
    id: allocator.next(),
    author: authorName,
    date: new Date().toISOString(),
    content: [
      {
        type: "paragraph",
        formatting: {},
        content: [{ type: "run", formatting: {}, content: [{ type: "text", text }] }],
      },
    ],
    ...(parentId !== undefined && { parentId }),
  };
}

export interface AddCommentOptions {
  paraId: string;
  text: string;
  author: string;
  search?: string;
}

/**
 * Locate the comment range (paragraph, narrowed by `search`), add the comment
 * mark, dispatch, and return the created Comment. Returns null when the schema
 * lacks a comment mark or the range can't be resolved. The caller owns adding
 * the comment to its own state and showing the sidebar.
 */
export function addCommentToRange(
  view: EditorView,
  options: AddCommentOptions,
  allocator: CommentIdAllocator,
): Comment | null {
  const { schema } = view.state;
  if (!schema.marks["comment"]) return null;

  const range = findParagraphByParaId(view.state.doc, options.paraId);
  if (!range) return null;

  let from = range.from + 1;
  let to = range.to - 1;
  if (options.search) {
    const textRange = findTextInPmParagraph(view.state.doc, range.from, range.to, options.search);
    if (!textRange) return null;
    from = textRange.from;
    to = textRange.to;
  }
  // Refuse an empty range (e.g. an empty paragraph with no search) so we never
  // create an orphan comment — one serialized to comments.xml with no
  // commentRangeStart/End anchor in document.xml.
  if (from >= to) return null;

  const comment = createComment(allocator, options.text, options.author);
  view.dispatch(
    view.state.tr.addMark(from, to, schema.marks["comment"].create({ commentId: comment.id })),
  );
  return comment;
}

export interface ProposeChangeOptions {
  paraId: string;
  search: string;
  replaceWith: string;
  author: string;
}

/**
 * Apply a tracked change (insertion / deletion / replace) to the located range.
 * `search === ''` means a pure insertion at the paragraph end; `replaceWith ===
 * ''` means a pure deletion. Refuses to layer onto an existing tracked change.
 * Returns true when a change was dispatched, false otherwise.
 */
export function applyProposedChange(
  view: EditorView,
  options: ProposeChangeOptions,
  allocator: CommentIdAllocator,
): boolean {
  const { schema } = view.state;
  if (!schema.marks["deletion"] || !schema.marks["insertion"]) return false;

  const range = findParagraphByParaId(view.state.doc, options.paraId);
  if (!range) return false;

  const isInsertion = options.search === "";
  const isDeletion = options.replaceWith === "";
  if (isInsertion && isDeletion) return false; // nothing to do

  let textFrom: number;
  let textTo: number;
  if (isInsertion) {
    // Insert at end of paragraph (just before closing token).
    textFrom = range.to - 1;
    textTo = range.to - 1;
  } else {
    const textRange = findTextInPmParagraph(view.state.doc, range.from, range.to, options.search);
    if (!textRange) return false;
    textFrom = textRange.from;
    textTo = textRange.to;
  }

  // Refuse to layer onto an existing tracked change.
  if (textFrom < textTo) {
    let overlaps = false;
    view.state.doc.nodesBetween(textFrom, textTo, (node) => {
      for (const m of node.marks) {
        if (m.type === schema.marks["insertion"] || m.type === schema.marks["deletion"]) {
          overlaps = true;
          return false;
        }
      }
      return true;
    });
    if (overlaps) return false;
  }

  const revisionId = allocator.next();
  const date = new Date().toISOString();
  const deletionMark = schema.marks["deletion"].create({
    revisionId,
    author: options.author,
    date,
  });
  const insertionMark = schema.marks["insertion"].create({
    revisionId,
    author: options.author,
    date,
  });

  let tr = view.state.tr;
  if (!isInsertion) tr = tr.addMark(textFrom, textTo, deletionMark);
  if (!isDeletion) tr = tr.insert(textTo, schema.text(options.replaceWith, [insertionMark]));
  view.dispatch(tr);
  return true;
}
