/**
 * Selective Save Module
 *
 * Orchestrates selective XML patching for the save flow.
 * Serializes full document.xml, validates patch safety, builds patched XML,
 * and calls applyUpdatesToZip() to produce the final DOCX.
 *
 * Returns null on any failure, signaling the caller to fall back to full repack.
 */

import type JSZip from "jszip";

import type { Document, BlockContent, Comment } from "../types/document";
import { parseCommentsExtended, type CommentExtendedInfo } from "./commentParser";
import { hasUnsynthesizedReplyRanges } from "./commentReplyMarkers";
import { validateFolioDocumentModel } from "./modelValidation";
import { RELATIONSHIP_TYPES } from "./relsParser";
import {
  applyUpdatesToZip,
  findMaxRId,
  updateCoreProperties,
  collectHeaderFooterUpdates,
  hasUnmaterializedHeaderFooter,
  hasModelDrivenPictureWatermark,
  COMMENTS_CONTENT_TYPE,
  COMMENTS_EXTENDED_PART,
  COMMENTS_EXTENDED_PART_LOWER,
  addCommentsExtendedOverride,
  addCommentsExtendedRelationship,
} from "./rezip";
import { DEFAULT_SELECTIVE_SAVE_MAX_BYTES } from "./selectiveSaveFlags";
import { buildPatchedDocumentXml, buildPatchedNoteXml, collectParaIds } from "./selectiveXmlPatch";
import {
  ensureThreadedCommentParaIds,
  serializeComments,
  serializeCommentsExtended,
} from "./serializer/commentSerializer";
import { serializeDocument } from "./serializer/documentSerializer";
import { serializeEndnotes, serializeFootnotes } from "./serializer/noteSerializer";

const SYNTHETIC_IMAGE_RID_PREFIX = "rId_img_";

/**
 * Check if document content has new images (data: URL without rId) or
 * new hyperlinks (href without rId). Combined into a single traversal
 * to avoid walking the block tree twice.
 */
function hasNewImagesOrHyperlinks(blocks: BlockContent[]): boolean {
  const runHasNewImage = (run: {
    content: { type: string; image?: { src?: string; rId?: string } }[];
  }): boolean =>
    run.content.some(
      (c) =>
        c.type === "drawing" &&
        c.image?.src?.startsWith("data:") === true &&
        (!c.image.rId || c.image.rId.startsWith(SYNTHETIC_IMAGE_RID_PREFIX)),
    );

  for (const block of blocks) {
    if (block.type === "paragraph") {
      for (const item of block.content) {
        if (item.type === "run") {
          if (runHasNewImage(item)) {
            return true;
          }
        } else if (item.type === "hyperlink" && item.href && !item.rId && !item.anchor) {
          return true;
        } else if (
          // A picture inserted/deleted/moved under track changes lives inside
          // an ins/del/moveFrom/moveTo wrapper. Without descending into them,
          // a freshly tracked image gets no rId allocated and the saved DOCX
          // references missing media. eigenpal #641.
          item.type === "insertion" ||
          item.type === "deletion" ||
          item.type === "moveFrom" ||
          item.type === "moveTo"
        ) {
          for (const sub of item.content) {
            if (sub.type === "run" && runHasNewImage(sub)) {
              return true;
            }
          }
        }
      }
    } else if (block.type === "table") {
      for (const row of block.rows) {
        for (const cell of row.cells) {
          if (hasNewImagesOrHyperlinks(cell.content)) {
            return true;
          }
        }
      }
    }
  }
  return false;
}

/**
 * Splice edited footnote/endnote paragraphs into their parts and return the
 * candidate ids that could NOT be routed to a note part (an id that is neither
 * a body nor a note paragraph — the caller bails on those).
 *
 * The document model only retains the normal notes, so this never rewrites a
 * whole note part (that would drop the separator notes the model omits): it
 * splices only the edited note paragraphs by `paraId` via
 * {@link buildPatchedNoteXml}, keeping separators and unedited notes byte-exact.
 *
 * Returns null to signal the caller should bail to a full repack (a note
 * paragraph could not be spliced safely).
 */
async function patchNoteParts(
  zip: JSZip,
  doc: Document,
  candidateIds: Set<string>,
  updates: Map<string, string>,
): Promise<Set<string> | null> {
  const footnotes = doc.package.footnotes ?? [];
  const endnotes = doc.package.endnotes ?? [];
  const remainingIds = new Set(candidateIds);
  if (footnotes.length === 0 && endnotes.length === 0) {
    return remainingIds;
  }

  // Word writes note parts at the conventional lowercase path; match
  // case-insensitively so a producer that cased them differently still hits the
  // existing entry (writing a new-cased path would duplicate the part).
  const findPart = (conventionalLowerPath: string) => {
    const direct = zip.file(conventionalLowerPath);
    if (direct) {
      return direct;
    }
    for (const [path, file] of Object.entries(zip.files)) {
      if (!file.dir && path.toLowerCase() === conventionalLowerPath) {
        return file;
      }
    }
    return null;
  };

  const patchPart = async (
    conventionalLowerPath: string,
    serialize: () => string,
  ): Promise<boolean> => {
    const file = findPart(conventionalLowerPath);
    if (!file) {
      return true;
    }
    const originalXml = await file.async("text");
    const originalIds = collectParaIds(originalXml);
    const partIds = new Set<string>();
    for (const id of remainingIds) {
      if (originalIds.has(id)) {
        partIds.add(id);
      }
    }
    if (partIds.size === 0) {
      return true;
    }
    const patched = buildPatchedNoteXml(originalXml, serialize(), partIds);
    if (patched === null) {
      return false;
    }
    updates.set(file.name, patched);
    for (const id of partIds) {
      remainingIds.delete(id);
    }
    return true;
  };

  if (
    footnotes.length > 0 &&
    !(await patchPart("word/footnotes.xml", () => serializeFootnotes(footnotes)))
  ) {
    return null;
  }
  if (
    endnotes.length > 0 &&
    !(await patchPart("word/endnotes.xml", () => serializeEndnotes(endnotes)))
  ) {
    return null;
  }

  return remainingIds;
}

const findZipEntryCaseInsensitive = (zip: JSZip, lowerPath: string): JSZip.JSZipObject | null => {
  const direct = zip.file(lowerPath);
  if (direct) {
    return direct;
  }
  for (const [path, file] of Object.entries(zip.files)) {
    if (!file.dir && path.toLowerCase() === lowerPath) {
      return file;
    }
  }
  return null;
};

const commentsExtendedInfoEqual = (
  a: Map<string, CommentExtendedInfo>,
  b: Map<string, CommentExtendedInfo>,
): boolean => {
  if (a.size !== b.size) {
    return false;
  }
  for (const [paraId, infoA] of a) {
    const infoB = b.get(paraId);
    if (!infoB || infoA.parentParaId !== infoB.parentParaId || infoA.done !== infoB.done) {
      return false;
    }
  }
  return true;
};

/**
 * Reconcile commentsExtended.xml on the selective path. Returns `false` to
 * signal the caller must bail to a full repack (which owns removing / rewiring
 * the part + content-type + relationship as a unit); `true` when the part is
 * unchanged or was safely patched into `updates`.
 *
 * - Threading unchanged (compared through the same parser, so attribute-order
 *   noise is ignored): leave the part byte-exact.
 * - Threading removed (no desired part but a baseline exists): bail — removing
 *   the part plus its dangling override/relationship cleanly is the full
 *   repack's job.
 * - Threading changed: write the part and ensure its override + relationship,
 *   bailing if the packaging files needed to wire it are absent.
 */
async function patchCommentsExtended(
  zip: JSZip,
  comments: Comment[],
  updates: Map<string, string>,
): Promise<boolean> {
  const desiredXml = serializeCommentsExtended(comments);
  const existing = findZipEntryCaseInsensitive(zip, COMMENTS_EXTENDED_PART_LOWER);

  if (!desiredXml) {
    return existing === null;
  }

  const baselineInfo = existing
    ? parseCommentsExtended(await existing.async("text"))
    : new Map<string, CommentExtendedInfo>();
  if (commentsExtendedInfoEqual(parseCommentsExtended(desiredXml), baselineInfo)) {
    return true;
  }

  updates.set(existing?.name ?? COMMENTS_EXTENDED_PART, desiredXml);
  return ensureCommentsExtendedPackaging(zip, updates);
}

/**
 * Add the content-type override + relationship for commentsExtended.xml if
 * absent, composing over any pending `updates`. Returns `false` when a packaging
 * file needed to wire the part is missing (the caller then bails to full repack
 * rather than write a part nothing references).
 */
async function ensureCommentsExtendedPackaging(
  zip: JSZip,
  updates: Map<string, string>,
): Promise<boolean> {
  const ctPath = "[Content_Types].xml";
  const ctXml = updates.get(ctPath) ?? (await zip.file(ctPath)?.async("text"));
  if (ctXml === undefined) {
    return false;
  }
  const nextCt = addCommentsExtendedOverride(ctXml);
  if (nextCt !== ctXml) {
    updates.set(ctPath, nextCt);
  }

  const relsPath = "word/_rels/document.xml.rels";
  const relsXml = updates.get(relsPath) ?? (await zip.file(relsPath)?.async("text"));
  if (relsXml === undefined) {
    return false;
  }
  const nextRels = addCommentsExtendedRelationship(relsXml);
  if (nextRels !== relsXml) {
    updates.set(relsPath, nextRels);
  }
  return true;
}

export type SelectiveSaveOptions = {
  /** Changed paragraph IDs to selectively patch */
  changedParaIds: Set<string>;
  /** Whether structural changes occurred (paragraph add/delete) */
  structuralChange: boolean;
  /** Whether any changes affected paragraphs without paraId */
  hasUntrackedChanges: boolean;
  /**
   * Maximum allowed `originalBuffer.byteLength` for the selective path. Above
   * this size the function returns null and the caller falls back to full
   * repack. Defaults to {@link DEFAULT_SELECTIVE_SAVE_MAX_BYTES}.
   */
  maxBytes?: number;
};

/**
 * Attempt a selective save — patch only changed paragraphs in document.xml.
 * Also updates comments, headers/footers, and core properties so that
 * all document parts stay in sync even when only paragraphs are patched.
 *
 * Returns the saved ArrayBuffer, or null if selective save is not possible
 * (caller should fall back to full repack).
 */
export async function attemptSelectiveSave(
  doc: Document,
  originalBuffer: ArrayBuffer,
  options: SelectiveSaveOptions,
): Promise<ArrayBuffer | null> {
  const { changedParaIds, structuralChange, hasUntrackedChanges } = options;
  const maxBytes = options.maxBytes ?? DEFAULT_SELECTIVE_SAVE_MAX_BYTES;

  // Bail out conditions — fall back to full repack
  if (structuralChange) {
    return null;
  }
  if (hasUntrackedChanges) {
    return null;
  }
  // Refuse very large inputs: the JSZip overhead on top of the original buffer
  // dominates memory cost, and full repack is the safer path at that scale.
  if (originalBuffer.byteLength > maxBytes) {
    return null;
  }
  // Check for new images/hyperlinks that need relationship management
  const content = doc.package.document.content;
  if (hasNewImagesOrHyperlinks(content)) {
    return null;
  }
  // A header/footer created in memory needs a new part + relationship +
  // [Content_Types] Override, which only the full repack path writes.
  if (hasUnmaterializedHeaderFooter(doc)) {
    return null;
  }
  // A picture watermark spanning multiple headers needs per-header image
  // relationship rebinding, which only the full repack path performs.
  if (hasModelDrivenPictureWatermark(doc)) {
    return null;
  }
  // A reply with no anchor of its own must get its parent's commentRange
  // markers, which means editing the parent's paragraph — not necessarily one of
  // `changedParaIds`. The full repack owns that synthesis, so hand off to it.
  if (hasUnsynthesizedReplyRanges(doc)) {
    return null;
  }
  if (!validateFolioDocumentModel(doc).valid) {
    return null;
  }

  const comments = doc.package.document.comments ?? [];
  const hasComments = comments.length > 0;
  const headerFooterUpdates = collectHeaderFooterUpdates(doc);

  try {
    const JSZip = (await import("jszip")).default;
    const zip = await JSZip.loadAsync(originalBuffer);
    const updates = new Map<string, string>();

    // Patch document.xml and the note parts (footnotes.xml / endnotes.xml). A
    // changed paraId lives in exactly one part, so partition against the body's
    // own paraIds: body ids patch document.xml, the rest are routed to the note
    // parts. A plain body edit has no note candidates, so the note parts are
    // never read or rewritten and stay verbatim.
    if (changedParaIds.size > 0) {
      const docXmlFile = zip.file("word/document.xml");
      if (!docXmlFile) {
        return null;
      }
      const originalDocXml = await docXmlFile.async("text");
      const bodyParaIds = collectParaIds(originalDocXml);

      const bodyChangedIds = new Set<string>();
      const noteCandidateIds = new Set<string>();
      for (const id of changedParaIds) {
        if (bodyParaIds.has(id)) {
          bodyChangedIds.add(id);
        } else {
          noteCandidateIds.add(id);
        }
      }

      if (noteCandidateIds.size > 0) {
        const unrouted = await patchNoteParts(zip, doc, noteCandidateIds, updates);
        // A splice failure, or a changed id that is neither a body nor a note
        // paragraph, falls back to full repack — matching the prior safety when
        // buildPatchedDocumentXml met a paraId it could not resolve.
        if (unrouted === null || unrouted.size > 0) {
          return null;
        }
      }

      if (bodyChangedIds.size > 0) {
        const serializedDocXml = serializeDocument(doc);
        const patchedDocXml = buildPatchedDocumentXml(
          originalDocXml,
          serializedDocXml,
          bodyChangedIds,
        );
        if (!patchedDocXml) {
          return null;
        }
        updates.set("word/document.xml", patchedDocXml);
      }
    }

    // Overwrite `word/comments.xml` whenever the source already had one,
    // even if the editor now has zero comments — otherwise the stale
    // entries linger in the saved file (the rezip baseline copies the
    // previous part as-is) and round-trip back as phantom threads.
    const hadCommentsFile = zip.file("word/comments.xml") !== null;
    if (hasComments || hadCommentsFile) {
      // Threaded/resolved comments need a stable last-paragraph paraId so
      // comments.xml and commentsExtended.xml reference the same key.
      ensureThreadedCommentParaIds(comments);
      updates.set("word/comments.xml", serializeComments(comments));
    }
    if (hasComments) {
      // Ensure [Content_Types].xml has an Override for comments.xml
      const ctFile = zip.file("[Content_Types].xml");
      if (ctFile) {
        const ctXml = await ctFile.async("text");
        if (!ctXml.includes("/word/comments.xml")) {
          updates.set(
            "[Content_Types].xml",
            ctXml.replace(
              "</Types>",
              `<Override PartName="/word/comments.xml" ContentType="${COMMENTS_CONTENT_TYPE}"/></Types>`,
            ),
          );
        }
      }

      // Ensure word/_rels/document.xml.rels has a Relationship for comments.xml
      const relsPath = "word/_rels/document.xml.rels";
      const relsFile = zip.file(relsPath);
      if (relsFile) {
        const relsXml = await relsFile.async("text");
        if (!relsXml.includes("comments.xml")) {
          const maxId = findMaxRId(relsXml);
          updates.set(
            relsPath,
            relsXml.replace(
              "</Relationships>",
              `<Relationship Id="rId${maxId + 1}" Type="${RELATIONSHIP_TYPES.comments}" Target="comments.xml"/></Relationships>`,
            ),
          );
        }
      }
    }

    // Reply threading / resolved state lives in commentsExtended.xml. Rewrite it
    // only when the model's threading differs from what the file already
    // encodes, so an unrelated body edit leaves the part byte-exact. Bail to
    // full repack when the part cannot be reconciled safely (e.g. a removal).
    if (!(await patchCommentsExtended(zip, comments, updates))) {
      return null;
    }

    // Serialize modified headers/footers
    for (const [path, xml] of headerFooterUpdates) {
      updates.set(path, xml);
    }

    // Update modification date in docProps/core.xml
    const corePropsFile = zip.file("docProps/core.xml");
    if (corePropsFile) {
      const corePropsXml = await corePropsFile.async("text");
      updates.set(
        "docProps/core.xml",
        updateCoreProperties(corePropsXml, { updateModifiedDate: true }),
      );
    }

    // Use the already-loaded zip to avoid a redundant decompression pass
    return await applyUpdatesToZip(zip, updates);
  } catch {
    // Any error — fall back to full repack
    return null;
  }
}
