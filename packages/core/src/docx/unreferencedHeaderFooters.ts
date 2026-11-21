/**
 * A header or footer goes with the last section reference to it. Resolving
 * the section change that added or removed a reference can leave a part the
 * opened package referenced with no section naming it, as rejecting a
 * comparison's imported header does; a save then drops it. A part the
 * opened package held without a reference, or one the session added and has
 * not referenced yet, is content, not a resolution's leftover, and stays.
 */

import { setDocumentBodyComments } from "@stll/docx-core/model";

import { withoutLostCommentThreads } from "../ai-edits/comment-lifecycle";
import type {
  BlockContent,
  Document,
  DocumentBody,
  HeaderFooter,
  SectionProperties,
} from "../types/document";
import { visitCommentMarkers } from "./commentAnchorIndex";

const HEADER_FOOTER_PARTS = {
  HEADER: "header",
  FOOTER: "footer",
} as const;

type HeaderFooterPart =
  (typeof HEADER_FOOTER_PARTS)[keyof typeof HEADER_FOOTER_PARTS];

const partKey = (part: HeaderFooterPart, relationshipId: string) =>
  `${part}:${relationshipId}`;

/** Every section record of the body: its paragraphs' and its final one. */
const sectionRecords = (body: DocumentBody): SectionProperties[] => {
  const records: SectionProperties[] = [];
  const visit = (blocks: readonly BlockContent[]) => {
    for (const block of blocks) {
      if (block.type === "paragraph") {
        if (block.sectionProperties !== undefined)
          records.push(block.sectionProperties);
      } else if (block.type === "table") {
        for (const row of block.rows)
          for (const cell of row.cells) visit(cell.content);
      } else if (block.type === "blockSdt" || block.type === "blockCustomXml") {
        visit(block.content);
      }
    }
  };
  visit(body.content);
  if (body.finalSectionProperties !== undefined)
    records.push(body.finalSectionProperties);
  for (const section of body.sections ?? []) records.push(section.properties);
  return records;
};

/**
 * The parts a section names, now or in a pending change of its references
 * (which accepting or rejecting gives back).
 */
const referencedParts = (document: Document): Set<string> => {
  const referenced = new Set<string>();
  const add = (
    selection:
      | Pick<SectionProperties, "headerReferences" | "footerReferences">
      | undefined,
  ) => {
    for (const { rId } of selection?.headerReferences ?? [])
      referenced.add(partKey(HEADER_FOOTER_PARTS.HEADER, rId));
    for (const { rId } of selection?.footerReferences ?? [])
      referenced.add(partKey(HEADER_FOOTER_PARTS.FOOTER, rId));
  };
  for (const record of sectionRecords(document.package.document)) {
    add(record);
    for (const change of record.propertyChanges ?? []) {
      add(change.previousReferences);
      add(change.previousProperties);
      add(change.currentProperties);
    }
  }
  return referenced;
};

type PartEntry = {
  part: HeaderFooterPart;
  relationshipId: string;
  story: HeaderFooter;
};

const partsOf = (document: Document): PartEntry[] => {
  const entries: PartEntry[] = [];
  for (const [relationshipId, story] of document.package.headers ?? [])
    entries.push({ part: HEADER_FOOTER_PARTS.HEADER, relationshipId, story });
  for (const [relationshipId, story] of document.package.footers ?? [])
    entries.push({ part: HEADER_FOOTER_PARTS.FOOTER, relationshipId, story });
  return entries;
};

/**
 * The headers and footers a section of `source` names, now or in a pending
 * change of its references.
 */
export const referencedHeaderFooters = (
  source: Document,
): ReadonlySet<string> => referencedParts(source);

/**
 * The `header:<rId>` and `footer:<rId>` keys of the parts `document` holds
 * that the opened package referenced (`referencedAtOpen`, from
 * {@link referencedHeaderFooters}) and no section names any more: the parts a
 * resolution retired.
 */
export const retiredHeaderFooters = (
  document: Document,
  referencedAtOpen: ReadonlySet<string>,
): ReadonlySet<string> => {
  const retired = new Set<string>();
  if (referencedAtOpen.size === 0) return retired;
  const referenced = referencedParts(document);
  for (const { part, relationshipId } of partsOf(document)) {
    const key = partKey(part, relationshipId);
    if (referencedAtOpen.has(key) && !referenced.has(key)) retired.add(key);
  }
  return retired;
};

/** Ids of the comments anchored anywhere in these stories. */
const anchoredComments = (
  stories: readonly (readonly BlockContent[])[],
): Set<number> => {
  const ids = new Set<number>();
  for (const blocks of stories)
    visitCommentMarkers(blocks, ({ item }) => ids.add(item.id));
  return ids;
};

/**
 * The document without the headers and footers the opened package
 * referenced (`referencedAtOpen`, from {@link referencedHeaderFooters}) that
 * no section names any more, and the comment threads anchored only in them.
 */
export const withoutUnreferencedHeaderFooters = (
  document: Document,
  referencedAtOpen: ReadonlySet<string>,
): Document => {
  const retired = retiredHeaderFooters(document, referencedAtOpen);
  if (retired.size === 0) return document;
  const keeps = (part: HeaderFooterPart, relationshipId: string) =>
    !retired.has(partKey(part, relationshipId));
  const parts = partsOf(document);
  const removed = parts.filter(
    ({ part, relationshipId }) => !keeps(part, relationshipId),
  );
  if (removed.length === 0) return document;
  const kept = (part: HeaderFooterPart, map: Map<string, HeaderFooter>) =>
    new Map([...map].filter(([relationshipId]) => keeps(part, relationshipId)));
  const { headers, footers, footnotes, endnotes } = document.package;
  const keptHeaders =
    headers === undefined
      ? undefined
      : kept(HEADER_FOOTER_PARTS.HEADER, headers);
  const keptFooters =
    footers === undefined
      ? undefined
      : kept(HEADER_FOOTER_PARTS.FOOTER, footers);
  const comments = document.package.document.comments;
  const keptComments = comments
    ? withoutLostCommentThreads(comments, {
        anchoredBefore: anchoredComments(
          removed.map(({ story }) => story.content),
        ),
        anchoredNow: anchoredComments([
          document.package.document.content,
          ...[
            ...(keptHeaders?.values() ?? []),
            ...(keptFooters?.values() ?? []),
          ].map((part) => part.content),
          ...[...(footnotes ?? []), ...(endnotes ?? [])].map(
            (note) => note.content,
          ),
        ]),
      })
    : undefined;
  const documentBody = { ...document.package.document };
  const commentsChanged =
    keptComments !== undefined && keptComments.length !== comments?.length;
  if (commentsChanged) setDocumentBodyComments(documentBody, keptComments);
  return {
    ...document,
    package: {
      ...document.package,
      ...(keptHeaders !== undefined && { headers: keptHeaders }),
      ...(keptFooters !== undefined && { footers: keptFooters }),
      ...(commentsChanged && { document: documentBody }),
    },
  };
};
