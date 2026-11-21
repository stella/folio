/**
 * Which ids a package's unnamed story paragraphs take, and the marking of a
 * story part's tree with them before its paragraphs are parsed. See
 * `unnamedParagraphIds` for why the ids stay out of the model.
 */
import { Result } from "better-result";

import {
  assignPackageParaIds,
  EnsureParaIdsError,
  packageStoryParts,
  storyParagraphSlots,
  type UnnamedParagraphIds,
} from "./ensureParaIds";
import { markSlotParagraphId } from "./unnamedParagraphIds";
import type { XmlElement } from "./xmlParser";

/** Story parts by name, as the package names them. */
export type PackageUnnamedParagraphIds = ReadonlyMap<
  string,
  UnnamedParagraphIds
>;

const MAIN_DOCUMENT_PART = "word/document.xml";

/** A paragraph open tag, under any prefix. */
const PARAGRAPH_OPEN_TAG = /<(?:[\w.-]+:)?p(?=[\s/>])[^>]*>/gu;
/** An id the parser keeps: eight hex digits, not all zero. */
const USABLE_PARA_ID = /:paraId\s*=\s*["'](?!0{8})[0-9A-Fa-f]{8}["']/u;
/**
 * A cell that may end without a paragraph, which it is then read with: one
 * that is empty, or whose last block may be a table (markers aside).
 */
const CELL_WITHOUT_LAST_PARAGRAPH =
  /<(?:[\w.-]+:)?tc(?:\s[^>]*)?\/>|<\/(?:[\w.-]+:)?(?:tbl|tcPr|sdt|customXml)>(?:\s*<(?:[\w.-]+:)?(?:bookmarkStart|bookmarkEnd|commentRangeStart|commentRangeEnd|permStart|permEnd|proofErr)\b[^>]*\/>)*\s*<\/(?:[\w.-]+:)?tc>|<(?:[\w.-]+:)?tc(?:\s[^>]*)?>\s*<\/(?:[\w.-]+:)?tc>/u;

/**
 * Whether a story part may hold a paragraph without a usable id: one whose
 * open tag names none, or a cell whose paragraph is implied. A cheap
 * overestimate that lets a fully named package skip the assignment.
 */
const holdsUnnamedParagraph = (xml: string): boolean => {
  if (CELL_WITHOUT_LAST_PARAGRAPH.test(xml)) return true;
  for (const [tag] of xml.matchAll(PARAGRAPH_OPEN_TAG))
    if (!USABLE_PARA_ID.test(tag)) return true;
  return false;
};

/**
 * The ids the package's unnamed story paragraphs take, from its parts as the
 * package holds them. A package the assignment refuses opens on neither
 * canonical surface, so there is nothing to agree with: none.
 */
export const unnamedPackageParagraphIds = (
  allXml: ReadonlyMap<string, string>,
  documentRels: string | null,
): PackageUnnamedParagraphIds => {
  const documentPartName = [...allXml.keys()].find(
    (name) => name.toLowerCase() === MAIN_DOCUMENT_PART,
  );
  if (documentPartName === undefined) return new Map();
  const { isTargetPart } = packageStoryParts({
    documentPartName,
    documentRels,
    entries: [...allXml.keys()],
  });
  // Most packages name every paragraph: then there is nothing to assign, and
  // reading each story part is skipped.
  const storyTexts = [...allXml].flatMap(([name, xml]) =>
    name === documentPartName || isTargetPart(name) ? [xml] : [],
  );
  if (!storyTexts.some(holdsUnnamedParagraph)) return new Map();
  const assigned = Result.try({
    try: () =>
      assignPackageParaIds({
        partTexts: allXml,
        documentPartName,
        isTargetPart,
      }),
    catch: (cause) => cause,
  });
  if (assigned.isOk()) return assigned.value;
  if (assigned.error instanceof EnsureParaIdsError) return new Map();
  throw assigned.error;
};

/** One story part's ids, by the part's name in any spelling of its case. */
export const unnamedPartParagraphIds = (
  ids: PackageUnnamedParagraphIds,
  partPath: string,
): UnnamedParagraphIds | undefined => {
  const exact = ids.get(partPath);
  if (exact !== undefined) return exact;
  const lower = partPath.toLowerCase();
  for (const [name, part] of ids) if (name.toLowerCase() === lower) return part;
  return undefined;
};

/**
 * Mark each unnamed paragraph of a story part's tree with its id. The tree's
 * paragraphs are numbered as the assignment numbered them; a tree that counts
 * them otherwise is not the part the assignment read, and is left unmarked.
 */
export const markUnnamedParagraphIds = (
  root: XmlElement,
  part: UnnamedParagraphIds | undefined,
): void => {
  if (part === undefined || part.byOrdinal.size === 0) return;
  const slots = storyParagraphSlots(root);
  if (slots.length !== part.paragraphs) return;
  for (const [ordinal, id] of part.byOrdinal) {
    const slot = slots[ordinal - 1];
    if (slot === undefined) continue;
    markSlotParagraphId(
      slot.type === "paragraph" ? slot.element : slot.cell,
      id,
    );
  }
};
