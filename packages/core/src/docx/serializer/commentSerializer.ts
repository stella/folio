/**
 * Comment Serializer
 *
 * Serializes Comment[] to OOXML comments.xml format.
 */

import { deterministicHexId } from "../../utils/hexId";
import type { Comment, Paragraph, Run } from "../../types/content";
import { escapeXml } from "./xmlUtils";

function serializeRunContent(run: Run): string {
  let xml = "<w:r>";
  // Run properties (minimal — just preserve formatting basics)
  const rPr: string[] = [];
  if (run.formatting?.bold) {
    rPr.push("<w:b/>");
  }
  if (run.formatting?.italic) {
    rPr.push("<w:i/>");
  }
  if (rPr.length > 0) {
    xml += `<w:rPr>${rPr.join("")}</w:rPr>`;
  }

  for (const c of run.content) {
    if (c.type === "text") {
      const preserveSpace = c.text !== c.text.trim() || c.text.includes("  ");
      xml += preserveSpace
        ? `<w:t xml:space="preserve">${escapeXml(c.text)}</w:t>`
        : `<w:t>${escapeXml(c.text)}</w:t>`;
    } else if (c.type === "break") {
      xml += "<w:br/>";
    }
  }
  xml += "</w:r>";
  return xml;
}

/**
 * `w14:paraId` / `w14:textId` open tag for a comment paragraph. Word keys
 * commentsExtended.xml (reply threading) and commentsExtensible.xml (UTC
 * dates) on the paraId of a comment's LAST paragraph, so the id must survive
 * serialization; the parser stores it on `Paragraph.paraId`.
 */
function commentParagraphOpenTag(p: Paragraph): string {
  const attrs: string[] = [];
  if (p.paraId) {
    attrs.push(`w14:paraId="${p.paraId}"`);
  }
  if (p.textId) {
    attrs.push(`w14:textId="${p.textId}"`);
  }
  return attrs.length > 0 ? `<w:p ${attrs.join(" ")}>` : "<w:p>";
}

function serializeParagraph(p: Paragraph): string {
  let xml = commentParagraphOpenTag(p);
  for (const item of p.content) {
    if (item.type === "run") {
      xml += serializeRunContent(item);
    }
  }
  xml += "</w:p>";
  return xml;
}

/** Serialize a paragraph, prepending an annotationRef run (required by Word in first paragraph of a comment) */
function serializeParagraphWithAnnotationRef(p: Paragraph): string {
  let xml = commentParagraphOpenTag(p);
  xml += '<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:annotationRef/></w:r>';
  for (const item of p.content) {
    if (item.type === "run") {
      xml += serializeRunContent(item);
    }
  }
  xml += "</w:p>";
  return xml;
}

function serializeComment(comment: Comment): string {
  const attrs: string[] = [
    `w:id="${comment.id}"`,
    `w:author="${escapeXml(comment.author)}"`,
  ];
  if (comment.initials) {
    attrs.push(`w:initials="${escapeXml(comment.initials)}"`);
  }
  if (comment.date) {
    attrs.push(`w:date="${escapeXml(comment.date)}"`);
  }

  let xml = `<w:comment ${attrs.join(" ")}>`;
  if (comment.content.length > 0) {
    // First paragraph must contain an annotationRef run for Word to link the comment
    // SAFETY: length > 0 verified by condition above
    xml += serializeParagraphWithAnnotationRef(comment.content[0]!);
    for (let i = 1; i < comment.content.length; i++) {
      // SAFETY: i < comment.content.length in for loop
      xml += serializeParagraph(comment.content[i]!);
    }
  } else {
    // Empty comment — still needs a paragraph with annotationRef
    xml +=
      '<w:p><w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:annotationRef/></w:r></w:p>';
  }
  xml += "</w:comment>";
  return xml;
}

const COMMENTS_HEADER =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  '<w:comments xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" ' +
  'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ' +
  'xmlns:o="urn:schemas-microsoft-com:office:office" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" ' +
  'xmlns:v="urn:schemas-microsoft-com:vml" ' +
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
  'xmlns:w10="urn:schemas-microsoft-com:office:word" ' +
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" ' +
  'xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" ' +
  'xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" ' +
  'xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" ' +
  'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" ' +
  'mc:Ignorable="w14 wp14">';

/**
 * Serialize comments array to comments.xml content. Returns a valid empty
 * `<w:comments/>` document for an empty array so callers can overwrite an
 * existing `word/comments.xml` part when the editor has removed the last
 * comment — leaving the previous file in place would otherwise re-emit
 * the orphaned comment threads on every save.
 */
export function serializeComments(comments: Comment[]): string {
  // Separate top-level comments and replies in a single pass
  const topLevel: Comment[] = [];
  const replies: Comment[] = [];
  for (const c of comments) {
    const comment: { parentId?: number | null } = c;
    const { parentId } = comment;
    (parentId === null || parentId === undefined ? topLevel : replies).push(c);
  }

  let xml = COMMENTS_HEADER;

  // Serialize top-level comments first, then replies
  for (const comment of topLevel) {
    xml += serializeComment(comment);
  }
  for (const reply of replies) {
    xml += serializeComment(reply);
  }

  xml += "</w:comments>";
  return xml;
}

/** The `w14:paraId` Word threads a comment by: its LAST paragraph's paraId. */
function commentThreadParaId(comment: Comment): string | undefined {
  return comment.content?.at(-1)?.paraId;
}

/**
 * The ids of comments that need a `commentsExtended.xml` entry: a reply, a
 * reply's parent (the thread root), or any comment carrying a resolved state.
 * A plain top-level comment set yields an empty set, so no part is written.
 */
function threadedCommentIds(comments: readonly Comment[]): Set<number> {
  const replyParents = new Set<number>();
  for (const comment of comments) {
    if (comment.parentId !== undefined) {
      replyParents.add(comment.parentId);
    }
  }
  const ids = new Set<number>();
  for (const comment of comments) {
    if (
      comment.parentId !== undefined ||
      replyParents.has(comment.id) ||
      comment.done !== undefined
    ) {
      ids.add(comment.id);
    }
  }
  return ids;
}

const commentPlainText = (comment: Comment): string => {
  let text = "";
  for (const paragraph of comment.content ?? []) {
    for (const item of paragraph.content ?? []) {
      if (item.type !== "run") {
        continue;
      }
      for (const runItem of item.content ?? []) {
        if (runItem.type === "text") {
          text += runItem.text;
        }
      }
    }
  }
  return text;
};

/**
 * Assign a deterministic `w14:paraId` to the LAST paragraph of every comment
 * that needs a commentsExtended entry (a reply, a reply's parent, or a resolved
 * comment) but has none. A document authored or loaded without comment paraIds
 * would otherwise have no stable key to thread through commentsExtended.xml, and
 * the thread link would be silently dropped. Word-authored ids are preserved;
 * only threaded, id-less paragraphs are filled. Deterministic (content- and
 * id-derived) so repeated saves mint the SAME id. MUST run before serializing
 * BOTH comments.xml and commentsExtended.xml so the two reference the same id.
 */
export function ensureThreadedCommentParaIds(comments: readonly Comment[]): void {
  const threaded = threadedCommentIds(comments);
  if (threaded.size === 0) {
    return;
  }

  const used = new Set<string>();
  for (const comment of comments) {
    for (const paragraph of comment.content ?? []) {
      if (paragraph.paraId) {
        used.add(paragraph.paraId.toUpperCase());
      }
    }
  }

  for (const comment of comments) {
    if (!threaded.has(comment.id)) {
      continue;
    }
    const last = (comment.content ?? []).at(-1);
    if (!last || last.paraId) {
      continue;
    }
    let paraId = deterministicHexId(`comment:${comment.id}:${commentPlainText(comment)}`);
    for (let salt = 1; used.has(paraId.toUpperCase()); salt++) {
      paraId = deterministicHexId(`comment:${comment.id}:${salt}`);
    }
    used.add(paraId.toUpperCase());
    last.paraId = paraId;
  }
}

type CommentExtendedEntry = {
  paraId: string;
  paraIdParent?: string;
  done: boolean;
};

/**
 * Build the `commentsExtended.xml` entries for the comments that participate in
 * a thread or carry a resolved state. Returns `null` when no comment needs an
 * entry (no replies, no parents-of-replies, no `done` state) so callers can
 * leave any existing part untouched — a plain top-level comment set gets no
 * `commentsExtended.xml`, matching how Word omits it.
 *
 * A `w15:commentEx` keys on the comment's LAST paragraph paraId; a reply's
 * `w15:paraIdParent` points at its parent comment's last-paragraph paraId. Ids
 * are guaranteed present by {@link ensureThreadedCommentParaIds}; the `!paraId`
 * skip is a last-ditch safety for a malformed model.
 */
function buildCommentExtendedEntries(comments: readonly Comment[]): CommentExtendedEntry[] | null {
  const threaded = threadedCommentIds(comments);
  if (threaded.size === 0) {
    return null;
  }

  const paraIdByCommentId = new Map<number, string>();
  for (const comment of comments) {
    const paraId = commentThreadParaId(comment);
    if (paraId) {
      paraIdByCommentId.set(comment.id, paraId);
    }
  }

  const entries: CommentExtendedEntry[] = [];
  for (const comment of comments) {
    if (!threaded.has(comment.id)) {
      continue;
    }
    const paraId = paraIdByCommentId.get(comment.id);
    if (!paraId) {
      continue;
    }
    const parentParaId =
      comment.parentId !== undefined ? paraIdByCommentId.get(comment.parentId) : undefined;
    entries.push({
      paraId,
      ...(parentParaId !== undefined ? { paraIdParent: parentParaId } : {}),
      done: comment.done ?? false,
    });
  }

  return entries.length > 0 ? entries : null;
}

const COMMENTS_EXTENDED_HEADER =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  '<w15:commentsEx xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ' +
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" ' +
  'mc:Ignorable="w15">';

/**
 * Serialize `commentsExtended.xml` (`w15:commentsEx`) for reply threading and
 * resolved state, or `null` when no comment needs an entry (see
 * {@link buildCommentExtendedEntries}). This is the part that makes Word render
 * a comment as a REPLY rather than a separate top-level thread.
 */
export function serializeCommentsExtended(comments: readonly Comment[]): string | null {
  const entries = buildCommentExtendedEntries(comments);
  if (!entries) {
    return null;
  }

  let xml = COMMENTS_EXTENDED_HEADER;
  for (const entry of entries) {
    const parentAttr =
      entry.paraIdParent !== undefined ? ` w15:paraIdParent="${entry.paraIdParent}"` : "";
    xml += `<w15:commentEx w15:paraId="${entry.paraId}"${parentAttr} w15:done="${entry.done ? "1" : "0"}"/>`;
  }
  xml += "</w15:commentsEx>";
  return xml;
}
