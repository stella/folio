/** Canonical snapshots own explicit section references; the PM projection does not. */
import JSZip from "jszip";
import { parseDocx } from "./parser";
import { withTrackedSectionEndpointRemoval } from "../internal/sectionEndpointResolution";
import { withSectionReferenceResolution } from "../internal/sectionReferenceResolution";
import type { RemovedSectionReference } from "../internal/sectionEndpointResolution";
import type { Document } from "../types/document";
import { readDocumentSectionFacts, type HeaderFooterReference } from "./documentSectionFacts";
import { serializeDocument } from "./serializer/documentSerializer";
import { assertSectionCarriersMatchModel } from "./rezip";

const referenceKey = ({ element, type, rId }: HeaderFooterReference): string =>
  `${element}:${type}:${rId}`;

const xmlFingerprint = async (xml: string): Promise<string> => {
  const digest = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(xml));
  let fingerprint = "";
  for (const byte of new Uint8Array(digest)) fingerprint += byte.toString(16).padStart(2, "0");
  return fingerprint;
};

type CanonicalStoryRepackOptions<T> = {
  document: Document;
  repack: () => Promise<T>;
};

/** Authorize exact canonical section removals and references missing from the committed snapshot. */
export const repackWithCanonicalStoryRemovals = async <T>({
  document,
  repack,
}: CanonicalStoryRepackOptions<T>): Promise<T> => {
  const originalBuffer = document.originalBuffer;
  if (!originalBuffer) return repack();
  const zip = await JSZip.loadAsync(originalBuffer);
  const xml = await zip.file("word/document.xml")?.async("text");
  if (!xml) return repack();
  const before = readDocumentSectionFacts(xml);
  const serializedXml = serializeDocument(document);
  const after = readDocumentSectionFacts(serializedXml);
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
  return withSectionReferenceResolution({
    document,
    removedReferences,
    repack: async () => {
      if (after.sectionCount >= before.sectionCount) return repack();
      const baseline = await parseDocx(originalBuffer, { preloadFonts: false });
      const baselineXml = serializeDocument(baseline);
      const baselineFacts = readDocumentSectionFacts(baselineXml);
      if (baselineFacts.sectionCount !== before.sectionCount) return repack();
      assertSectionCarriersMatchModel({
        doc: baseline,
        serializedSectionCount: baselineFacts.sectionCount,
      });
      return withTrackedSectionEndpointRemoval({
        document,
        resolution: {
          type: "tracked-section-endpoint-removal",
          sourceParagraphEndpointCount: before.sectionCount,
          expectedParagraphEndpointCount: after.sectionCount,
          sourceEndpointFingerprint: await xmlFingerprint(baselineXml),
          expectedEndpointFingerprint: await xmlFingerprint(serializedXml),
          removedReferences,
        },
        repack,
      });
    },
  });
};
