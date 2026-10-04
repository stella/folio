/**
 * Headless API for the document watermark.
 *
 * Mirrors the `content-controls` module's shape: pure functions that
 * read and write a `Document` immutably. Word stores watermarks inside
 * header parts; folio's getter scans every header and returns the
 * first watermark found, and the setter writes the modeled watermark
 * to every header part so all sections render it.
 */

import { copyParagraphPropertySource } from "../docx/paragraphPropertySource";
import type { Document, Watermark } from "../types/document";
import {
  planDocumentWatermarkCoverage,
  setDocumentWatermarkWithCoverage,
  ensureDocumentWatermarkHeaderCoverage,
} from "@stll/docx-core/ops";

export type { Watermark, TextWatermark, PictureWatermark } from "../types/document";

/**
 * Schemes allowed for a picture watermark's external image target
 * (`TargetMode="External"` in the saved package's relationships).
 */
const ALLOWED_EXTERNAL_IMAGE_SCHEMES = new Set(["http:", "https:"]);

/**
 * Whether `target` is safe to save as a watermark picture's external image
 * relationship. A watermark dialog lets an author type an arbitrary string
 * for this field; without a scheme check it becomes a `TargetMode="External"`
 * relationship verbatim (see `docx/rezip.ts`), which would let a `file:` URL
 * or UNC path (`\\host\share\...`) into the exported `.docx` — a resource
 * whatever later opens the file (e.g. Word) would try to resolve. Only
 * `http:`/`https:` targets are allowed.
 */
export function isAllowedExternalWatermarkImageUrl(target: string): boolean {
  try {
    return ALLOWED_EXTERNAL_IMAGE_SCHEMES.has(new URL(target).protocol);
  } catch {
    return false;
  }
}

/**
 * Read the document's watermark. Walks every header part and returns
 * the first watermark encountered (header insertion order). Returns
 * `undefined` when no header carries one.
 */
export function getDocumentWatermark(doc: Document): Watermark | undefined {
  const headers = doc.package.headers;
  if (!headers) {
    return undefined;
  }
  for (const header of headers.values()) {
    if (header.watermark) {
      return header.watermark;
    }
  }
  return undefined;
}

/**
 * Set (or clear) the document's watermark. Writes the modeled `watermark` to
 * every existing header part (clearing the captured raw VML so the serializer
 * synthesizes from the model), then extends coverage to the header parts a
 * document needs but lacks — see {@link ensureWatermarkHeaderCoverage}. Pass
 * `undefined` to remove the watermark from every header.
 */
const legacyWatermarkCoverage = (document: Document) => {
  const used = new Set([
    ...(document.package.relationships?.keys() ?? []),
    ...(document.package.headers?.keys() ?? []),
    ...(document.package.footers?.keys() ?? []),
  ]);
  return planDocumentWatermarkCoverage(document).map((type) => {
    const base = `rId_wm_${type}`;
    let rId = base;
    let suffix = 1;
    while (used.has(rId)) rId = `${base}_${++suffix}`;
    used.add(rId);
    return { type, rId, content: [] };
  });
};

type RestoreLegacyWatermarkSourcesOptions = { before: Document; after: Document };
/** The model owner cannot transfer folio's private paragraph source registry. */
const restoreLegacyWatermarkSources = ({ before, after }: RestoreLegacyWatermarkSourcesOptions) => {
  for (const [index, source] of before.package.document.content.entries()) {
    const target = after.package.document.content.at(index);
    if (source.type === "paragraph" && target?.type === "paragraph" && source !== target)
      copyParagraphPropertySource(target, source);
  }
  return after;
};

export function setDocumentWatermark(doc: Document, watermark: Watermark | undefined): Document {
  return restoreLegacyWatermarkSources({
    before: doc,
    after: setDocumentWatermarkWithCoverage({
      document: doc,
      authority: "legacy",
      change: watermark === undefined ? { kind: "remove" } : { kind: "set", watermark },
      coverage: watermark === undefined ? [] : legacyWatermarkCoverage(doc),
    }).unwrap(),
  });
}

/** Extend forward-inherited coverage without replacing existing watermark source payloads. */
export function ensureWatermarkHeaderCoverage(doc: Document, watermark: Watermark): Document {
  return restoreLegacyWatermarkSources({
    before: doc,
    after: ensureDocumentWatermarkHeaderCoverage({
      document: doc,
      authority: "legacy",
      watermark,
      coverage: legacyWatermarkCoverage(doc),
    }).unwrap(),
  });
}
