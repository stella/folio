/** Canonical snapshots own explicit section references; the PM projection does not. */
import JSZip from "jszip";
import { withSectionReferenceResolution } from "../internal/sectionReferenceResolution";
import type { RemovedSectionReference } from "../internal/sectionEndpointResolution";
import type { Document } from "../types/document";
import { readDocumentSectionFacts, type HeaderFooterReference } from "./documentSectionFacts";
import { serializeDocument } from "./serializer/documentSerializer";

const referenceKey = ({ element, type, rId }: HeaderFooterReference): string =>
  `${element}:${type}:${rId}`;

type CanonicalStoryRepackOptions<T> = {
  document: Document;
  repack: () => Promise<T>;
};

/** Authorize only references missing from the committed canonical snapshot, with exact multiplicity. */
export const repackWithCanonicalStoryRemovals = async <T>({
  document,
  repack,
}: CanonicalStoryRepackOptions<T>): Promise<T> => {
  if (!document.originalBuffer) return repack();
  const zip = await JSZip.loadAsync(document.originalBuffer);
  const xml = await zip.file("word/document.xml")?.async("text");
  if (!xml) return repack();
  const before = readDocumentSectionFacts(xml);
  const after = readDocumentSectionFacts(serializeDocument(document));
  const remaining = new Map<string, number>();
  for (const reference of after.headerFooterReferences) {
    const key = referenceKey(reference);
    remaining.set(key, (remaining.get(key) ?? 0) + 1);
  }
  const removedReferences: RemovedSectionReference[] = [];
  for (const reference of before.headerFooterReferences) {
    const key = referenceKey(reference);
    const count = remaining.get(key) ?? 0;
    if (count > 0) {
      remaining.set(key, count - 1);
      continue;
    }
    removedReferences.push({
      part: reference.element === "headerReference" ? "header" : "footer",
      type: reference.type,
      relationshipId: reference.rId,
    });
  }
  return withSectionReferenceResolution({ document, removedReferences, repack });
};
