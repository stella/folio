/** Documents a headless reviewer reads, for properties that compare its engines. */
import fc from "fast-check";

import type { Document } from "../../packages/docx-core/src/model/document";
import { reviewDocumentArbitrary } from "../../packages/docx-core/src/ops/__tests__/documentArbitraries";
import { packageDocumentArbitrary } from "./packageOperationArbitraries";

/**
 * Every story a reviewer lists (headers, footers, notes, a comment thread,
 * content controls and a table), or a body with review marks of every shape
 * the review properties draw.
 */
export const reviewerDocumentArbitrary: fc.Arbitrary<Document> = fc.oneof(
  packageDocumentArbitrary,
  reviewDocumentArbitrary,
);

/**
 * Range markers that a package holds only in pairs, by their pair. A bookmark
 * boundary without a partner is source content a package keeps, so the
 * generator's lone bookmarks stay.
 */
const PAIRED_MARKERS: Readonly<Record<string, string>> = {
  moveToRangeStart: "moveTo",
  moveToRangeEnd: "moveTo",
  moveFromRangeStart: "moveFrom",
  moveFromRangeEnd: "moveFrom",
};

/** A plain object: a Map, a Date or a buffer is kept as it is. */
const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" &&
  value !== null &&
  Object.getPrototypeOf(value) === Object.prototype;

/** The pair a marker belongs to and its id, or nothing for any other node. */
const markerOf = (node: unknown): { key: string; end: boolean } | undefined => {
  if (!isRecord(node) || typeof node["type"] !== "string") return undefined;
  const pair = PAIRED_MARKERS[node["type"]];
  if (pair === undefined) return undefined;
  return {
    key: `${pair}:${String(node["id"])}`,
    end: node["type"].endsWith("End"),
  };
};

/**
 * The document with every move range marker whose partner is missing left
 * out, so that the writer packages it: the review generator draws markers one
 * by one.
 */
const withPairedMarkers = (document: Document): Document => {
  const starts = new Set<string>();
  const ends = new Set<string>();
  const collect = (value: unknown): void => {
    if (Array.isArray(value)) {
      for (const item of value) collect(item);
      return;
    }
    if (!isRecord(value)) return;
    const marker = markerOf(value);
    if (marker) (marker.end ? ends : starts).add(marker.key);
    for (const child of Object.values(value)) collect(child);
  };
  const keep = (value: unknown): unknown => {
    if (Array.isArray(value))
      return value
        .filter((item) => {
          const marker = markerOf(item);
          return (
            marker === undefined ||
            (starts.has(marker.key) && ends.has(marker.key))
          );
        })
        .map(keep);
    if (!isRecord(value)) return value;
    return Object.fromEntries(
      Object.entries(value).map(([key, child]) => [key, keep(child)]),
    );
  };
  collect(document);
  // SAFETY: only whole marker nodes leave their arrays; every record keeps
  // its shape.
  return keep(document) as Document;
};

/**
 * Bodies with review marks of every shape that the writer packages, for
 * properties about the body's edits and resolutions. Headers, footers, notes
 * and text boxes have their own story properties.
 */
export const bodyReviewerDocumentArbitrary: fc.Arbitrary<Document> =
  reviewDocumentArbitrary.map(withPairedMarkers);
