/** Allocate package identities before the deterministic watermark operation is journaled. */
import { Result, TaggedError, panic } from "better-result";
import {
  DOCUMENT_OP_TYPES,
  planDocumentWatermarkCoverage,
  planDocumentWatermarkHosts,
  type SetDocumentWatermarkOp,
} from "@stll/docx-core/ops";
import type { Document } from "../types/document";
import { CANONICAL_GAP } from "../types/canonicalCapabilities";
import type { CanonicalWatermarkRequest } from "../types/canonicalWatermark";
import { isAllowedExternalWatermarkImageUrl } from "../watermark/index";
import { withCanonicalParagraphIds } from "./canonicalOperations";

class CanonicalWatermarkError extends TaggedError("CanonicalWatermarkError")<{
  message: string;
  gap: typeof CANONICAL_GAP.watermark;
}> {}

export const createCanonicalWatermarkOperation = (
  document: Document,
  change: CanonicalWatermarkRequest,
) => {
  if (
    change.kind === "set" &&
    change.watermark.kind === "picture" &&
    change.watermark.imageTargetExternal === true &&
    (change.watermark.imageTarget === undefined ||
      !isAllowedExternalWatermarkImageUrl(change.watermark.imageTarget))
  )
    return Result.err(
      new CanonicalWatermarkError({
        gap: CANONICAL_GAP.watermark,
        message: "Picture watermarks require an HTTP or HTTPS image target.",
      }),
    );
  const types = change.kind === "remove" ? [] : planDocumentWatermarkCoverage(document);
  const hostRIds = change.kind === "remove" ? [] : planDocumentWatermarkHosts(document);
  const paragraphs = withCanonicalParagraphIds(
    Array.from({ length: types.length + hostRIds.length }, () => ({
      type: "paragraph",
      content: [],
    })),
    document,
  );
  const paragraphIdAt = (index: number) => {
    const paragraph = paragraphs.at(index);
    if (paragraph?.type !== "paragraph" || paragraph.paraId === undefined)
      return panic("Watermark host allocation requires an identified paragraph.");
    return paragraph.paraId;
  };
  const occupied = new Set([
    ...(document.package.relationships?.keys() ?? []),
    ...(document.package.headers?.keys() ?? []),
    ...(document.package.footers?.keys() ?? []),
  ]);
  const coverage = types.map((type, index) => {
    const base = `rId_wm_${type}`;
    let rId = base;
    let suffix = 1;
    while (occupied.has(rId)) rId = `${base}_${++suffix}`;
    occupied.add(rId);
    return { type, rId, paraId: paragraphIdAt(index) };
  });
  const hosts = hostRIds.map((rId, index) => ({
    rId,
    paraId: paragraphIdAt(types.length + index),
  }));
  return Result.ok({
    type: DOCUMENT_OP_TYPES.SET_DOCUMENT_WATERMARK,
    change,
    coverage,
    hosts,
  } satisfies SetDocumentWatermarkOp);
};
