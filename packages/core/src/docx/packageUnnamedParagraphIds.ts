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
