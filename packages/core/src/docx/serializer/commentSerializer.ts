/**
 * Comment Serializer
 *
 * Serializes Comment[] to OOXML comments.xml format.
 */

import { commentThreadParaId } from "../commentThreadKey";
import { deterministicHexId } from "../../utils/hexId";
import type { Comment, Paragraph } from "../../types/content";
import type { TextFormatting } from "../../types/formatting";
import { serializePartElement, type OoxmlNamespacePrefix } from "./partNamespaces";
import { serializeWithPreservedChildren } from "../containerChildren";
import { serializeParagraph } from "./paragraphSerializer";
import { serializeTextFormatting } from "./textFormattingSerializer";
import { escapeXmlAttribute } from "@stll/docx-core";

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
  const attrs: string[] = [
    `w:id="${comment.id}"`,
    `w:author="${escapeXmlAttribute(comment.author)}"`,
  ];
  if (comment.initials !== undefined) {
    attrs.push(`w:initials="${escapeXmlAttribute(comment.initials)}"`);
  }
  if (comment.date) {
    attrs.push(`w:date="${escapeXmlAttribute(comment.date)}"`);
  }

  const paragraphs = comment.content.map((paragraph, index) =>
    // First paragraph must contain an annotationRef run for Word to link the comment
    index === 0
      ? serializeParagraphWithAnnotationRef(paragraph, comment.annotationReferenceFormatting)
      : serializeParagraph(paragraph),
  );
  if (paragraphs.length === 0 && comment.preserved === undefined) {
    // Empty comment — still needs a paragraph with annotationRef
    paragraphs.push(
      `<w:p>${serializeAnnotationReference(comment.annotationReferenceFormatting)}</w:p>`,
    );
  }

  const body = serializeWithPreservedChildren(paragraphs, comment.preserved);
  return `<w:comment ${attrs.join(" ")}>${body}</w:comment>`;
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
 * What every comment part is written from: one ordered list of comment ids, and
 * the facts each id carries.
 *
 * `word/comments.xml` and `word/commentsExtended.xml` describe the same
 * comments and have to agree about which is which. They agreed by each walking
 * the `Comment[]` they were handed and pairing entries up by position, so any
 * disagreement about order — one part reordered, the other not — silently moved
 * a comment's thread link, resolved state and paraId onto a different comment.
 * Positions cannot disagree if neither part has one: both walk `order` and look
 * every fact up by `w:id`, the only identity the package itself has.
 */
export type CommentPartPlan = {
  /** The `w:id`s to write, in order, each exactly once. */
  readonly order: readonly number[];
  readonly byId: ReadonlyMap<number, Comment>;
  /** The `w14:paraId` each comment is threaded by, minted where the model had none. */
  readonly threadParaIdById: ReadonlyMap<number, string>;
  /** The comments needing a `commentsExtended.xml` entry: replies, their parents, resolved comments. */
  readonly threadedIds: ReadonlySet<number>;
};

/**
 * Plan both comment parts from the model, in the model's own order.
 *
 * The order is the document's: the order `word/comments.xml` listed the
 * comments in is the order it is written back in, so a save neither reshuffles
 * a reviewer's threads nor hands the next parse a different `comments[]` than
 * the one it read.
 *
 * Planning also mints the `w14:paraId` a thread needs when the model has none —
 * a comment written in the editor has no Word-authored id, and without one the
 * `commentsExtended.xml` link would be dropped. Minting is deterministic
 * (comment id and text derived) so repeated saves mint the same id, and it
 * happens here rather than at each call site because both parts must see it.
 */
export const planCommentParts = (comments: readonly Comment[]): CommentPartPlan => {
  const byId = new Map<number, Comment>();
  const order: number[] = [];
  // A duplicate `w:id` is resolved the way the parser resolves it
  // (`normalizeCommentIds`): the first definition is the one every marker in
  // the body addresses, so it is the one that gets written.
  for (const comment of comments) {
    if (byId.has(comment.id)) {
      continue;
    }
    byId.set(comment.id, comment);
    order.push(comment.id);
  }

  const planned = [...byId.values()];
  const threadedIds = threadedCommentIds(planned);
  const used = new Set<string>();
  for (const comment of planned) {
    for (const paragraph of comment.content ?? []) {
      if (paragraph.paraId) {
        used.add(paragraph.paraId.toUpperCase());
      }
    }
  }

  const threadParaIdById = new Map<number, string>();
  for (const id of order) {
    // SAFETY: `order` holds exactly the keys of `byId`.
    const comment = byId.get(id)!;
    const existing = commentThreadParaId((comment.content ?? []).map(({ paraId }) => paraId));
    if (existing) {
      threadParaIdById.set(id, existing);
      continue;
    }
    const last = (comment.content ?? []).at(-1);
    if (!threadedIds.has(id) || !last) {
      continue;
    }
    let minted = deterministicHexId(`comment:${id}:${commentPlainText(comment)}`);
    for (let salt = 1; used.has(minted.toUpperCase()); salt++) {
      minted = deterministicHexId(`comment:${id}:${salt}`);
    }
    used.add(minted.toUpperCase());
    last.paraId = minted;
    threadParaIdById.set(id, minted);
  }

  return { order, byId, threadParaIdById, threadedIds };
};

/**
 * Serialize a comment plan to comments.xml content. Returns a valid empty
 * `<w:comments/>` document for an empty plan so callers can overwrite an
 * existing `word/comments.xml` part when the editor has removed the last
 * comment — leaving the previous file in place would otherwise re-emit
 * the orphaned comment threads on every save.
 */
export function serializeComments(
  { order, byId }: CommentPartPlan,
  sourceBindings?: ReadonlyMap<string, string>,
): string {
  let body = "";
  for (const id of order) {
    // SAFETY: `order` holds exactly the keys of `byId`.
    body += serializeComment(byId.get(id)!);
  }

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
 * A `w15:commentEx` keys on the comment's thread paraId; a reply's
 * `w15:paraIdParent` on its parent's. Both come from the plan, so a reply and
 * its parent name each other by `w:id` right up to the attribute.
 */
function buildCommentExtendedEntries({
  order,
  byId,
  threadParaIdById,
  threadedIds,
}: CommentPartPlan): CommentExtendedEntry[] | null {
  if (threadedIds.size === 0) {
    return null;
  }

  const entries: CommentExtendedEntry[] = [];
  for (const id of order) {
    if (!threadedIds.has(id)) {
      continue;
    }
    const paraId = threadParaIdById.get(id);
    // A malformed model can hold a comment with no paragraph to carry an id.
    if (!paraId) {
      continue;
    }
    // SAFETY: `order` holds exactly the keys of `byId`.
    const { parentId, done } = byId.get(id)!;
    const parentParaId = parentId !== undefined ? threadParaIdById.get(parentId) : undefined;
    entries.push({
      paraId,
      ...(parentParaId !== undefined ? { paraIdParent: parentParaId } : {}),
      done: done ?? false,
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
export function serializeCommentsExtended(plan: CommentPartPlan): string | null {
  const entries = buildCommentExtendedEntries(plan);
  if (!entries) {
    return null;
  }

  const body = entries
    .map((entry) => {
      const parentAttr =
        entry.paraIdParent !== undefined
          ? ` w15:paraIdParent="${escapeXmlAttribute(entry.paraIdParent)}"`
          : "";
      return `<w15:commentEx w15:paraId="${escapeXmlAttribute(entry.paraId)}"${parentAttr} w15:done="${entry.done ? "1" : "0"}"/>`;
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
