import { TaggedError } from "better-result";
import type { Document } from "../types/document";
import { CANONICAL_GAP } from "../types/canonicalCapabilities";
import { canonicalJson } from "../utils/canonicalJson";
import { getDocumentSourceBaseline } from "./headerFooterVerbatim";
import type { SaveDiagnostic } from "./saveDiagnostics";

export class CanonicalResourceSaveRefusalError extends TaggedError(
  "CanonicalResourceSaveRefusalError",
)<{
  message: string;
  gap: typeof CANONICAL_GAP.resourceReplacement;
  diagnostic: Extract<SaveDiagnostic, { type: "canonicalResourceReplacement" }>;
}> {}

/** Full and selective package writers preserve existing style/media source entries. */
export const canonicalResourceReplacementOf = (document: Document) => {
  const source = getDocumentSourceBaseline(document);
  if (source.type !== "captured" || !document.originalBuffer) return undefined;
  const baseline = source.resourceStyles;
  const current = document.package.styles;
  if (baseline) {
    const { styles: sourceStyles, ...sourceDefaults } = baseline;
    const { styles: currentStyles, ...currentDefaults } = current ?? { styles: [] };
    const currentById = new Map(currentStyles.map((style) => [style.styleId, style]));
    if (
      canonicalJson(sourceDefaults) !== canonicalJson(currentDefaults) ||
      sourceStyles.some(
        (style) => canonicalJson(style) !== canonicalJson(currentById.get(style.styleId)),
      )
    )
      return "word/styles.xml";
  }
  for (const [path, media] of source.resourceMedia) {
    const currentMedia = document.package.media?.get(path);
    if (!currentMedia || media.mimeType !== currentMedia.mimeType) return path;
    const sourceBytes = new Uint8Array(media.data);
    const currentBytes = new Uint8Array(currentMedia.data);
    if (
      sourceBytes.length !== currentBytes.length ||
      sourceBytes.some((byte, index) => byte !== currentBytes[index])
    )
      return path;
  }
  return undefined;
};
