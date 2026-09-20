/**
 * Comment Parser - Parse comments.xml, commentsExtensible.xml, and
 * commentsExtended.xml.
 *
 * - `comments.xml` (w) carries the comment author, local date, and body.
 * - `commentsExtensible.xml` (w16cex, Word 2016+) carries reliable UTC
 *   timestamps via `w16cex:dateUtc` — Word's `w:date` is local time
 *   without an offset and so is ambiguous.
 * - `commentsExtended.xml` (w15, Word 2013+) carries reply-thread
 *   parent links via `w15:paraIdParent` and the resolved/done state
 *   via `w15:done`. Cross-referenced via the `w14:paraId` on
 *   `w:comment` and the matching `w15:paraId` on `w15:commentEx`.
 *
 * OOXML Reference:
 * - Comments: w:comments
 * - Comment: w:comment (w:id, w:author, w:date, w:initials, w14:paraId)
 * - Comment content: child w:p elements
 */

import { PARSE_WARNING_CODES } from "@stll/docx-core/model";

import { CAPTURE, dispatchChildren } from "./containerChildren";
import { commentThreadParaId } from "./commentThreadKey";
import { PARA_ID_NAMESPACE_URIS, paraIdAttribute, paraIdParentAttribute } from "./paraIdAttribute";
import type { ParseContext } from "./parseContext";
import type {
  Comment,
  Paragraph,
  Theme,
  RelationshipMap,
  MediaFile,
  TextFormatting,
} from "../types/document";
import { parseParagraph } from "./paragraphParser";
import { cloneParagraphWithPropertySource } from "./paragraphPropertySource";
import { parseRunProperties } from "./runParser";
import type { StyleMap } from "./styleParser";
import {
  parseXml,
  findChild,
  getChildElements,
  getAttribute,
  getAttributeByNamespaceUri,
  getLocalName,
  type XmlElement,
  parseOnOffValue,
} from "./xmlParser";

/** The whole lexical form of `ST_DecimalNumber`: an optional sign and digits. */
const DECIMAL_NUMBER = /^[+-]?\d+$/u;

type ParsedFirstCommentParagraph = {
  paragraph: Paragraph;
  annotationReferenceFormatting?: TextFormatting;
};

const DEFAULT_ANNOTATION_REFERENCE_STYLE_ID = "CommentReference";

const normalizeAnnotationReferenceFormatting = (
  formatting: TextFormatting | undefined,
): TextFormatting | undefined => {
  if (
    formatting?.styleId === DEFAULT_ANNOTATION_REFERENCE_STYLE_ID &&
    Object.keys(formatting).length === 1
  ) {
    return undefined;
  }
  return formatting;
};

const normalizeFirstCommentParagraph = (
  paragraphElement: XmlElement,
  paragraph: Paragraph,
  theme: Theme | null,
): ParsedFirstCommentParagraph => {
  const firstContentElement = getChildElements(paragraphElement).find(
    (element) => getLocalName(element.name) !== "pPr",
  );
  if (
    !firstContentElement ||
    getLocalName(firstContentElement.name) !== "r" ||
    !findChild(firstContentElement, "w", "annotationRef")
  ) {
    return { paragraph };
  }

  const runProperties = findChild(firstContentElement, "w", "rPr");
  const annotationReferenceFormatting = normalizeAnnotationReferenceFormatting(
    runProperties ? parseRunProperties(runProperties, theme) : undefined,
  );
  // The source run is the comment's reference mark, which `serializeComment`
  // re-emits from `annotationReferenceFormatting`. Drop it from the editable
  // content whether it parsed to nothing or to the preserved `w:annotationRef`
  // capture; keeping it would give the comment two reference marks on save.
  const firstParsedContent = paragraph.content.at(0);
  const normalizedParagraph =
    firstParsedContent?.type === "run" &&
    firstParsedContent.content.every((item) => item.type === "preservedXml" && item.text === "")
      ? cloneParagraphWithPropertySource(paragraph, { content: paragraph.content.slice(1) })
      : paragraph;

  return {
    paragraph: normalizedParagraph,
    ...(annotationReferenceFormatting !== undefined ? { annotationReferenceFormatting } : {}),
  };
};

/**
 * Build a lookup from paraId → dateUtc from commentsExtensible.xml
 *
 * The XML structure is:
 * <w16cex:commentsExtensible>
 *   <w16cex:comment w16cex:paraId="..." w16cex:dateUtc="2024-02-10T14:30:45Z"/>
 * </w16cex:commentsExtensible>
 */
function parseCommentsExtensible(xml: string): Map<string, string> {
  const dateUtcByParaId = new Map<string, string>();

  const root = parseXml(xml);

  // Find the root element (may be w16cex:commentsExtensible or similar)
  const container = findChild(root, "w16cex", "commentsExtensible") ?? root;
  for (const child of getChildElements(container)) {
    const localName = child.name?.replace(/^.*:/u, "") ?? "";
    if (localName !== "comment") {
      continue;
    }

    const paraId = paraIdAttribute(child);

    // `dateUtc` rides on the same element and is read by the same rule: a
    // timestamp from a foreign namespace is not this comment's timestamp.
    const dateUtc =
      getAttributeByNamespaceUri(child, PARA_ID_NAMESPACE_URIS, "dateUtc") ??
      child.attributes?.["w16cex:dateUtc"] ??
      child.attributes?.["w15:dateUtc"];

    // First entry wins on a duplicate paraId, as it does in commentsExtended
    // and for a duplicate `w:id`: a second timestamp for the same key would
    // otherwise land on whichever comment the key resolves to.
    const key = paraId === undefined ? null : paraId.toUpperCase();
    if (key && dateUtc && !dateUtcByParaId.has(key)) {
      dateUtcByParaId.set(key, String(dateUtc));
    }
  }

  return dateUtcByParaId;
}

export type CommentExtendedInfo = {
  parentParaId?: string;
  done?: boolean;
};

/**
 * Build a lookup from paraId → reply-thread info from
 * commentsExtended.xml. The XML structure is:
 *
 * ```xml
 * <w15:commentsEx>
 *   <w15:commentEx w15:paraId="..." w15:done="0"/>
 *   <w15:commentEx w15:paraId="..." w15:paraIdParent="..." w15:done="1"/>
 * </w15:commentsEx>
 * ```
 *
 * `w15:paraIdParent` points at the parent thread's paraId; `w15:done`
 * (`"1"` / `"true"`) marks the thread resolved.
 */
export function parseCommentsExtended(xml: string): Map<string, CommentExtendedInfo> {
  const infoByParaId = new Map<string, CommentExtendedInfo>();

  const root = parseXml(xml);
  const container = findChild(root, "w15", "commentsEx") ?? root;
  for (const child of getChildElements(container)) {
    const localName = child.name?.replace(/^.*:/u, "") ?? "";
    if (localName !== "commentEx") {
      continue;
    }

    const paraId = paraIdAttribute(child);
    if (!paraId) {
      continue;
    }
    // Two entries for one paraId make the thread link and the resolved state
    // ambiguous. Keep the first, the way a duplicate `w:id` keeps the first
    // `w:comment` (see `normalizeCommentIds`), so the reading is deterministic
    // rather than dependent on where the duplicate sits in the part.
    const key = paraId.toUpperCase();
    if (infoByParaId.has(key)) {
      continue;
    }

    const parentParaId = paraIdParentAttribute(child);
    const doneAttr = getAttribute(child, "w15", "done") ?? child.attributes?.["w15:done"];

    const info: CommentExtendedInfo = {};
    if (parentParaId) {
      info.parentParaId = String(parentParaId).toUpperCase();
    }
    if (doneAttr !== undefined) {
      info.done = parseOnOffValue(String(doneAttr).toLowerCase()) ?? false;
    }
    infoByParaId.set(key, info);
  }

  return infoByParaId;
}

/**
 * Parse comments.xml into an array of Comment objects.
 *
 * If `commentsExtensibleXml` is provided, UTC timestamps are
 * cross-referenced via paraId and preferred over the ambiguous w:date
 * local time. If `commentsExtendedXml` is provided, reply-thread
 * parent links (`parentId`) and resolved state (`done`) are populated.
 */
export function parseComments(
  commentsXml: string | null,
  styles: StyleMap | null,
  theme: Theme | null,
  rels: RelationshipMap,
  media: Map<string, MediaFile>,
  commentsExtensibleXml?: string | null,
  commentsExtendedXml?: string | null,
  context?: ParseContext,
): Comment[] {
  if (!commentsXml) {
    return [];
  }

  const root = parseXml(commentsXml);

  // Build UTC date lookup from commentsExtensible (Word 2016+).
  const dateUtcByParaId = commentsExtensibleXml
    ? parseCommentsExtensible(commentsExtensibleXml)
    : new Map<string, string>();

  // Build reply-thread + done lookup from commentsExtended (Word 2013+).
  const extendedByParaId = commentsExtendedXml
    ? parseCommentsExtended(commentsExtendedXml)
    : new Map<string, CommentExtendedInfo>();

  const commentsEl = findChild(root, "w", "comments") ?? root;
  const children = getChildElements(commentsEl);
  // Each comment is carried next to the paraId it threads by rather than
  // looked up by position later: `w15:paraIdParent` names a comment, and a
  // lookup that pairs the two arrays up by index attributes one comment's
  // thread link and resolved state to another the moment the arrays differ.
  const parsed: { comment: Comment; threadParaId: string | null }[] = [];
  // Track the paraId → comment-id mapping so we can resolve
  // `w15:paraIdParent` (which references the parent comment's paraId,
  // not its `w:id`) to a numeric `parentId` once every comment is parsed.
  const commentIdByParaId = new Map<string, number>();

  for (const child of children) {
    const localName = child.name?.replace(/^.*:/u, "") ?? "";
    if (localName !== "comment") {
      continue;
    }

    // Reading a missing or unparseable `w:id` as 0 manufactured a duplicate of
    // whichever comment genuinely holds id 0, and made a comment no marker can
    // address look addressable. A comment with no id anchors nothing, so drop
    // it and say so.
    const rawId = getAttribute(child, "w", "id");
    // The whole attribute has to be the number: `parseInt` reads `7pt` as 7,
    // which is the duplicate this guard exists to prevent, wearing an id that
    // another comment genuinely holds.
    const id =
      rawId !== null && DECIMAL_NUMBER.test(rawId) ? Number.parseInt(rawId, 10) : Number.NaN;
    if (Number.isNaN(id)) {
      context?.warn({
        code: PARSE_WARNING_CODES.missingCommentId,
        element: "w:comment",
        ...(rawId === null ? {} : { value: rawId }),
      });
      continue;
    }
    const rawAuthor = getAttribute(child, "w", "author");
    const author = parseCommentAuthor(rawAuthor);
    const rawInitials = getAttribute(child, "w", "initials");
    const initials = rawInitials !== null ? String(rawInitials) : undefined;
    const rawDate = getAttribute(child, "w", "date");
    const localDate = rawDate !== null ? String(rawDate) : undefined;

    // The paraId join key used by commentsExtensible (UTC dates) and
    // commentsExtended (reply links) may live on `w:comment` itself,
    // or on its paragraphs — both layouts occur in the wild and
    // exporters disagree. Check the wrapper first, then apply the
    // shared paragraph rule the serializer writes the key back by.
    const rawParaId =
      paraIdAttribute(child) ??
      commentThreadParaId(
        getChildElements(child)
          .filter((sub) => (sub.name?.replace(/^.*:/u, "") ?? "") === "p")
          .map(paraIdAttribute),
      );
    const paraId = rawParaId ? String(rawParaId).toUpperCase() : null;

    const dateUtc = paraId ? dateUtcByParaId.get(paraId) : undefined;
    // Prefer UTC date over ambiguous local date
    const date = dateUtc ?? localDate;

    const extendedInfo = paraId ? extendedByParaId.get(paraId) : undefined;
    const done = extendedInfo?.done;

    // Parse comment content. `Comment.content` is a paragraph list, so every
    // other body child a `CT_Comment` may hold — a table, an equation, a
    // content control, the bookmark and range markers a reviewer's selection
    // leaves behind — goes to the sink at its source position rather than on
    // the floor.
    const paragraphs: Paragraph[] = [];
    let annotationReferenceFormatting: TextFormatting | undefined;
    const preserved = dispatchChildren({
      element: child,
      container: "w:comment",
      capturePosition: () => paragraphs.length,
      handlers: {
        p: (contentChild) => {
          const paragraph = parseParagraph(contentChild, styles, theme, null, rels, media);
          if (paragraphs.length > 0) {
            paragraphs.push(paragraph);
            return;
          }
          const normalized = normalizeFirstCommentParagraph(contentChild, paragraph, theme);
          annotationReferenceFormatting = normalized.annotationReferenceFormatting;
          paragraphs.push(normalized.paragraph);
        },
        altChunk: CAPTURE,
        bookmarkEnd: CAPTURE,
        bookmarkStart: CAPTURE,
        commentRangeEnd: CAPTURE,
        commentRangeStart: CAPTURE,
        customXml: CAPTURE,
        customXmlDelRangeEnd: CAPTURE,
        customXmlDelRangeStart: CAPTURE,
        customXmlInsRangeEnd: CAPTURE,
        customXmlInsRangeStart: CAPTURE,
        customXmlMoveFromRangeEnd: CAPTURE,
        customXmlMoveFromRangeStart: CAPTURE,
        customXmlMoveToRangeEnd: CAPTURE,
        customXmlMoveToRangeStart: CAPTURE,
        del: CAPTURE,
        ins: CAPTURE,
        moveFrom: CAPTURE,
        moveFromRangeEnd: CAPTURE,
        moveFromRangeStart: CAPTURE,
        moveTo: CAPTURE,
        moveToRangeEnd: CAPTURE,
        moveToRangeStart: CAPTURE,
        permEnd: CAPTURE,
        permStart: CAPTURE,
        proofErr: CAPTURE,
        sdt: CAPTURE,
        tbl: CAPTURE,
      },
    });

    // Two comments sharing a paraId make every `w15:paraIdParent` naming it
    // ambiguous. Resolve it to the first, as a duplicate `w:id` resolves to
    // the first `w:comment` (see `normalizeCommentIds`), so which comment a
    // reply hangs off does not depend on how far down the part the twin sits.
    if (paraId && !commentIdByParaId.has(paraId)) {
      commentIdByParaId.set(paraId, id);
    }

    parsed.push({
      comment: {
        id,
        author,
        ...(initials !== undefined ? { initials } : {}),
        ...(date !== undefined ? { date } : {}),
        ...(done !== undefined ? { done } : {}),
        ...(annotationReferenceFormatting !== undefined ? { annotationReferenceFormatting } : {}),
        ...(preserved !== undefined ? { preserved } : {}),
        content: paragraphs,
      },
      threadParaId: paraId,
    });
  }

  // Second pass: resolve `w15:paraIdParent` → numeric parent comment id.
  // A reply whose parent paraId is unknown (e.g. the parent was deleted
  // from comments.xml but a stale `w15:commentEx` remains) is left as a
  // top-level comment so it isn't silently dropped.
  const comments: Comment[] = [];
  for (const { comment, threadParaId } of parsed) {
    const parentParaId = threadParaId
      ? extendedByParaId.get(threadParaId)?.parentParaId
      : undefined;
    const parentId = parentParaId ? commentIdByParaId.get(parentParaId) : undefined;
    if (parentId !== undefined && parentId !== comment.id) {
      comment.parentId = parentId;
    }
    comments.push(comment);
  }
  return comments;
}

const parseCommentAuthor = (author: string | null): string => {
  if (author === "") {
    return "";
  }

  return author?.trim() || "Unknown";
};
