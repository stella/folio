/**
 * Comment Serializer
 *
 * Serializes Comment[] to OOXML comments.xml format.
 */

import { deterministicHexId } from "../../utils/hexId";
import type { Comment, Paragraph } from "../../types/content";
import type { TextFormatting } from "../../types/formatting";
import { serializePartElement, type OoxmlNamespacePrefix } from "./partNamespaces";
import { serializeParagraph } from "./paragraphSerializer";
import { serializeTextFormatting } from "./runSerializer";
import { escapeXml } from "./xmlUtils";

const DEFAULT_ANNOTATION_REFERENCE_PROPERTIES =
  '<w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>';
const PARAGRAPH_PROPERTIES_END = "</w:pPr>";

const serializeAnnotationReference = (formatting: TextFormatting | undefined): string => {
  const properties = formatting
    ? serializeTextFormatting(formatting)
    : DEFAULT_ANNOTATION_REFERENCE_PROPERTIES;
  return `<w:r>${properties}<w:annotationRef/></w:r>`;
};

/** Serialize a paragraph, prepending an annotationRef run (required by Word in first paragraph of a comment) */
function serializeParagraphWithAnnotationRef(
  paragraph: Paragraph,
  formatting: TextFormatting | undefined,
): string {
  const xml = serializeParagraph(paragraph);
  const annotationReference = serializeAnnotationReference(formatting);
  const propertiesEnd = xml.indexOf(PARAGRAPH_PROPERTIES_END);
  if (propertiesEnd !== -1) {
    const contentStart = propertiesEnd + PARAGRAPH_PROPERTIES_END.length;
    return `${xml.slice(0, contentStart)}${annotationReference}${xml.slice(contentStart)}`;
  }

  return xml.replace(
    /^<w:p(?=[\s>])[^>]*>/u,
    (openingTag) => `${openingTag}${annotationReference}`,
  );
}

function serializeComment(comment: Comment): string {
  const attrs: string[] = [`w:id="${comment.id}"`, `w:author="${escapeXml(comment.author)}"`];
  if (comment.initials !== undefined) {
    attrs.push(`w:initials="${escapeXml(comment.initials)}"`);
  }
  if (comment.date) {
    attrs.push(`w:date="${escapeXml(comment.date)}"`);
  }

  let xml = `<w:comment ${attrs.join(" ")}>`;
  if (comment.content.length > 0) {
    // First paragraph must contain an annotationRef run for Word to link the comment
    // SAFETY: length > 0 verified by condition above
    xml += serializeParagraphWithAnnotationRef(
      comment.content[0]!,
      comment.annotationReferenceFormatting,
    );
    for (let i = 1; i < comment.content.length; i++) {
      // SAFETY: i < comment.content.length in for loop
      xml += serializeParagraph(comment.content[i]!);
    }
  } else {
    // Empty comment — still needs a paragraph with annotationRef
    xml += `<w:p>${serializeAnnotationReference(comment.annotationReferenceFormatting)}</w:p>`;
  }
  xml += "</w:comment>";
  return xml;
}

// Prefixes word/comments.xml declares whether or not the comment bodies use
// them, so a comment carrying a drawing or a raw-replayed extension lands on a
// root that already declares its prefix.
const COMMENTS_BASELINE_PREFIXES = [
  "wpc",
  "mc",
  "o",
  "r",
  "m",
  "v",
  "wp",
  "w10",
  "w",
  // Every ignorable prefix needs a binding even when no comment uses its elements.
  "w14",
  "wp14",
  "wpg",
  "wpi",
  "wne",
  "wps",
] as const satisfies readonly OoxmlNamespacePrefix[];

const COMMENTS_EXTENDED_BASELINE_PREFIXES = [
  "mc",
  "w",
  "w15",
] as const satisfies readonly OoxmlNamespacePrefix[];

const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

/**
 * Serialize comments array to comments.xml content. Returns a valid empty
 * `<w:comments/>` document for an empty array so callers can overwrite an
 * existing `word/comments.xml` part when the editor has removed the last
 * comment — leaving the previous file in place would otherwise re-emit
 * the orphaned comment threads on every save.
 */
export function serializeComments(
  comments: Comment[],
  sourceBindings?: ReadonlyMap<string, string>,
): string {
  // Separate top-level comments and replies in a single pass
  const topLevel: Comment[] = [];
  const replies: Comment[] = [];
  for (const c of comments) {
    const comment: { parentId?: number | null } = c;
    const { parentId } = comment;
    (parentId === null || parentId === undefined ? topLevel : replies).push(c);
  }

  // Serialize top-level comments first, then replies
  const body =
    topLevel.map((comment) => serializeComment(comment)).join("") +
    replies.map((reply) => serializeComment(reply)).join("");

  return (
    XML_DECLARATION +
    serializePartElement({
      partPath: "word/comments.xml",
      rootName: "w:comments",
      baselinePrefixes: COMMENTS_BASELINE_PREFIXES,
      sourceBindings,
      body,
    })
  );
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

  const body = entries
    .map((entry) => {
      const parentAttr =
        entry.paraIdParent !== undefined
          ? ` w15:paraIdParent="${escapeXml(entry.paraIdParent)}"`
          : "";
      return `<w15:commentEx w15:paraId="${escapeXml(entry.paraId)}"${parentAttr} w15:done="${entry.done ? "1" : "0"}"/>`;
    })
    .join("");

  return (
    XML_DECLARATION +
    serializePartElement({
      partPath: "word/commentsExtended.xml",
      rootName: "w15:commentsEx",
      baselinePrefixes: COMMENTS_EXTENDED_BASELINE_PREFIXES,
      sourceBindings: undefined,
      body,
    })
  );
}
