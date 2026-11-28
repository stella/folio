/**
 * The source's comments carried onto the output.
 *
 * With `comments: "keep"` a comment travels when the output places its
 * reference: the comment itself goes into the output's comments part (a
 * new one, or after the template's own comments with ids above theirs),
 * and its range is written where the output places both of its ends. The
 * parts that extend comments (replies and done state, durable ids, dates,
 * people) come along, holding only the entries of the comments carried.
 * A comment whose reference the output does not place (it sits in an
 * omitted block, in a note or a header) is left out and counted.
 */
import { Result } from "better-result";

import { captureVerbatimXml } from "../../verbatimCapture";
import { cloneElement, getLocalName, type XmlElement } from "../../xmlParser";
import type { AnalyzedDocument } from "./blocks";
import { readPartText } from "./package";
import { type ExpectedSegment, segmentPieces } from "./plan";
import { readFamilyPart, type SourcePart } from "./sourceHeaders";
import type { TemplatePlanError } from "./types";
import {
  elementChildren,
  integerOrNull,
  maxOf,
  rangeOf,
  readRawPart,
  spliceXml,
  walkElements,
  wAttr,
  wName,
} from "./xml";

const WORDPROCESSINGML =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.";

/** The comments part and the parts that extend it, by what they hold. */
export const COMMENT_PARTS = {
  comments: {
    type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
    contentType: `${WORDPROCESSINGML}comments+xml`,
    fileName: "comments.xml",
  },
  commentsExtended: {
    type: "http://schemas.microsoft.com/office/2011/relationships/commentsExtended",
    contentType: `${WORDPROCESSINGML}commentsExtended+xml`,
    fileName: "commentsExtended.xml",
  },
  commentsIds: {
    type: "http://schemas.microsoft.com/office/2016/09/relationships/commentsIds",
    contentType: `${WORDPROCESSINGML}commentsIds+xml`,
    fileName: "commentsIds.xml",
  },
  commentsExtensible: {
    type: "http://schemas.microsoft.com/office/2018/08/relationships/commentsExtensible",
    contentType: `${WORDPROCESSINGML}commentsExtensible+xml`,
    fileName: "commentsExtensible.xml",
  },
  people: {
    type: "http://schemas.microsoft.com/office/2011/relationships/people",
    contentType: `${WORDPROCESSINGML}people+xml`,
    fileName: "people.xml",
  },
} as const;

export type CommentExtension = Exclude<keyof typeof COMMENT_PARTS, "comments">;

export const COMMENT_EXTENSIONS: readonly CommentExtension[] = [
  "commentsExtended",
  "commentsIds",
  "commentsExtensible",
  "people",
];

/** A part of a document read raw. */
export type CommentExtensionPart = {
  kind: CommentExtension;
  path: string;
  xml: string;
  root: XmlElement;
};

export type SourceComments = {
  part: SourcePart;
  /** Every comment of the part a copy can carry, by its id. */
  byId: Map<string, XmlElement>;
  /** Comments the part holds, carriable or not. */
  total: number;
  extensions: CommentExtensionPart[];
};

/** The path of the part a document's main part names for `type`. */
export const commentPartPath = (
  { pkg }: AnalyzedDocument,
  type: string,
): string | null => {
  for (const relationship of pkg.documentRelationships.values()) {
    if (relationship.type === type && relationship.targetMode !== "External") {
      const path = pkg.resolveDocumentTarget(relationship.target);
      return pkg.paths.has(path) ? path : null;
    }
  }
  return null;
};

/** A part read raw; null where the package holds none there. */
export const readCommentExtension = async (
  document: AnalyzedDocument,
  kind: CommentExtension,
): Promise<Result<CommentExtensionPart | null, TemplatePlanError>> => {
  const path = commentPartPath(document, COMMENT_PARTS[kind].type);
  const xml = await readPartText(document.pkg, path);
  if (Result.isError(xml)) {
    return xml;
  }
  const part =
    path === null || xml.value === null ? null : readRawPart(path, xml.value);
  return Result.ok(
    path === null || part === null
      ? null
      : { kind, path, xml: part.xml, root: part.root },
  );
};

/** The source's comments and the parts that extend them; null without any. */
export const readSourceComments = async (
  source: AnalyzedDocument,
): Promise<Result<SourceComments | null, TemplatePlanError>> => {
  const path = commentPartPath(source, COMMENT_PARTS.comments.type);
  const part = await readFamilyPart(source, path);
  if (Result.isError(part)) {
    return part;
  }
  if (part.value === null) {
    return Result.ok(null);
  }
  const byId = new Map<string, XmlElement>();
  for (const block of part.value.blocks) {
    const id = wName(block) === "comment" ? wAttr(block, "id") : null;
    if (id !== null) {
      byId.set(id, block);
    }
  }
  const extensions: CommentExtensionPart[] = [];
  for (const kind of COMMENT_EXTENSIONS) {
    // oxlint-disable-next-line no-await-in-loop -- serialized reads keep the archive budget exact
    const extension = await readCommentExtension(source, kind);
    if (Result.isError(extension)) {
      return extension;
    }
    if (extension.value !== null) {
      extensions.push(extension.value);
    }
  }
  return Result.ok({
    part: part.value,
    byId,
    total: elementChildren(part.value.root).filter(
      (child) => wName(child) === "comment",
    ).length,
    extensions,
  });
};

/** The comment marks a list of placed paragraphs' segments holds, by kind. */
export const placedCommentMarks = (
  paragraphs: Iterable<readonly ExpectedSegment[]>,
): Record<"start" | "end" | "reference", Set<string>> => {
  const placed = {
    start: new Set<string>(),
    end: new Set<string>(),
    reference: new Set<string>(),
  };
  for (const segments of paragraphs) {
    for (const piece of segmentPieces(segments)) {
      const mark = piece.atom?.comment;
      if (piece.kind === "commentMark" && mark !== undefined) {
        placed[mark.kind].add(mark.id);
      }
    }
  }
  return placed;
};

/** The highest comment id a document's comments part uses; -1 for none. */
export const highestCommentId = (root: XmlElement | null): number =>
  maxOf(
    elementChildren(root)
      .filter((child) => wName(child) === "comment")
      .map((child) => integerOrNull(wAttr(child, "id")))
      .filter((id) => id !== null),
    -1,
  );

const attributeByLocalName = (
  element: XmlElement,
  localName: string,
): string | null => {
  for (const [name, value] of Object.entries(element.attributes ?? {})) {
    if (
      value !== undefined &&
      getLocalName(name) === localName &&
      !name.startsWith("xmlns")
    ) {
      return String(value);
    }
  }
  return null;
};

/** Paragraph ids (`w14:paraId`) the given comments hold. */
export const commentParagraphIds = (
  comments: Iterable<XmlElement>,
): Set<string> => {
  const ids = new Set<string>();
  for (const comment of comments) {
    walkElements(comment, (element) => {
      const id = attributeByLocalName(element, "paraId");
      if (id !== null) {
        ids.add(id.toUpperCase());
      }
    });
  }
  return ids;
};

/**
 * The entries of an extension part the carried comments use: an entry
 * that names a paragraph (`paraId`) is kept when a carried comment holds
 * it, one that names only a durable id when a kept entry names that id.
 * People stay whole.
 */
export const keptExtensionEntries = (
  extensions: readonly CommentExtensionPart[],
  paragraphIds: ReadonlySet<string>,
): Map<CommentExtension, XmlElement[]> => {
  const durable = new Set<string>();
  const out = new Map<CommentExtension, XmlElement[]>();
  const ordered = extensions.toSorted(
    (a, b) =>
      COMMENT_EXTENSIONS.indexOf(a.kind) - COMMENT_EXTENSIONS.indexOf(b.kind),
  );
  for (const { kind, root } of ordered) {
    const kept = elementChildren(root).filter((entry) => {
      if (kind === "people") {
        return true;
      }
      const paragraph = attributeByLocalName(entry, "paraId");
      const durableId = attributeByLocalName(entry, "durableId");
      if (paragraph !== null) {
        const keep = paragraphIds.has(paragraph.toUpperCase());
        if (keep && durableId !== null) {
          durable.add(durableId.toUpperCase());
        }
        return keep;
      }
      return durableId !== null && durable.has(durableId.toUpperCase());
    });
    out.set(kind, kept);
  }
  return out;
};

/** The part's own XML without the entries it does not keep; null when a
 *  removal cannot be placed. */
export const withEntriesOnly = (
  part: CommentExtensionPart,
  kept: readonly XmlElement[],
): string | null => {
  const keep = new Set(kept);
  const splices = [];
  for (const entry of elementChildren(part.root)) {
    if (keep.has(entry)) {
      continue;
    }
    const range = rangeOf(entry);
    if (range === null) {
      return null;
    }
    splices.push({ ...range, newXml: "" });
  }
  return splices.length === 0 ? part.xml : spliceXml(part.xml, splices);
};

/** Entries of a source part as markup for another part's root: each brings
 *  the namespaces its own root declares. */
export const entriesMarkup = (
  root: XmlElement,
  entries: readonly XmlElement[],
): string => {
  const declared: Record<string, string | number | undefined> = {};
  for (const [name, value] of Object.entries(root.attributes ?? {})) {
    if (name.startsWith("xmlns") || name.endsWith(":Ignorable")) {
      declared[name] = value;
    }
  }
  return entries
    .map((entry) =>
      captureVerbatimXml(
        cloneElement(entry, {
          attributes: { ...declared, ...entry.attributes },
        }),
      ),
    )
    .join("");
};
