import { panic } from "better-result";

import type { Document, HeaderFooterType } from "../types/document";

export type RemovedSectionReference = {
  part: "header" | "footer";
  type: HeaderFooterType;
  relationshipId: string;
};

/**
 * Exact resolution evidence produced when tracked-change resolution removes
 * section endpoints. The tracker binds this count transition to exact endpoint
 * fingerprints; removed references bind the only relationship losses that the
 * same resolution may legitimately cause during package serialization.
 */
export type TrackedSectionEndpointRemoval = {
  type: "tracked-section-endpoint-removal";
  sourceParagraphEndpointCount: number;
  expectedParagraphEndpointCount: number;
  sourceEndpointFingerprint: string;
  expectedEndpointFingerprint: string;
  removedReferences: readonly RemovedSectionReference[];
};

type ActiveSectionEndpointRepack =
  | { status: "available"; authorization: TrackedSectionEndpointRemoval }
  | { status: "consumed"; authorization: TrackedSectionEndpointRemoval };
const activeSectionEndpointRepacks = new WeakMap<Document, ActiveSectionEndpointRepack>();

const snapshotAuthorization = (
  authorization: TrackedSectionEndpointRemoval,
): TrackedSectionEndpointRemoval => {
  const removedReferences = Object.freeze(
    authorization.removedReferences.map((reference) => Object.freeze({ ...reference })),
  );
  return Object.freeze({ ...authorization, removedReferences });
};

type WithTrackedSectionEndpointRemovalOptions<T> = {
  document: Document;
  resolution: TrackedSectionEndpointRemoval;
  repack: () => Promise<T>;
};

/** Run one repack with an exact authorization that no public save API accepts. */
export const withTrackedSectionEndpointRemoval = async <T>({
  document,
  resolution,
  repack,
}: WithTrackedSectionEndpointRemovalOptions<T>): Promise<T> => {
  if (activeSectionEndpointRepacks.has(document)) {
    panic("A tracked section-endpoint repack is already active for this document");
  }
  const snapshot = snapshotAuthorization(resolution);
  activeSectionEndpointRepacks.set(document, { status: "available", authorization: snapshot });
  try {
    const result = await repack();
    if (activeSectionEndpointRepacks.get(document)?.status !== "consumed") {
      panic("Tracked section-endpoint repack did not consume its authorization");
    }
    return result;
  } finally {
    activeSectionEndpointRepacks.delete(document);
  }
};

/** @internal Consume the active authorization once for the exact document object. */
export const consumeTrackedSectionEndpointRemoval = (
  document: Document,
): TrackedSectionEndpointRemoval | null => {
  const active = activeSectionEndpointRepacks.get(document);
  if (!active || active.status === "consumed") {
    return null;
  }
  activeSectionEndpointRepacks.set(document, {
    status: "consumed",
    authorization: active.authorization,
  });
  return active.authorization;
};
