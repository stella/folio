import { TaggedError } from "better-result";

import {
  findAttributeByNamespaceUri,
  getLocalName,
  getNamespaceUri,
  type XmlElement,
  WORDPROCESSINGML_NAMESPACE_URIS,
} from "./xmlParser";
import { rewriteStreamingXmlDecimalAttributes } from "./streamingXmlParser";
import { assertXmlResourceLimits, XmlResourceLimitError } from "./xmlResourceLimits";

/**
 * Two revision elements in one package claimed the same `w:id`.
 *
 * `w:id` on a revision element is unique across the package, so a collision is
 * a package a consumer may reject rather than a cosmetic detail. The
 * normalization below hands every id it emits to one choke point, which throws
 * this rather than letting the duplicate reach the ZIP.
 */
export class RevisionIdCollisionError extends TaggedError("RevisionIdCollisionError")<{
  message: string;
  revisionId: number;
  part: string;
}> {}

export const REVISION_ELEMENT_NAMES = new Set([
  "cellDel",
  "cellIns",
  "cellMerge",
  "del",
  "ins",
  "moveFrom",
  "moveTo",
  "numberingChange",
  "pPrChange",
  "rPrChange",
  "sectPrChange",
  "tblGridChange",
  "tblPrChange",
  "tblPrExChange",
  "tcPrChange",
  "trPrChange",
]);

const REVISION_ELEMENT_CANDIDATE = new RegExp(
  `<(?:[^\\s<>/:]+:)?(?:${[...REVISION_ELEMENT_NAMES].join("|")})(?:[\\s/>])`,
  "u",
);

/**
 * Whether an XML part could hold a revision element at all. A save that wrote
 * none cannot have created a collision, so this is what lets an exit skip the
 * pass instead of reading every part of the package to prove nothing changed.
 */
export const containsRevisionElement = (xml: string): boolean =>
  REVISION_ELEMENT_CANDIDATE.test(xml);

type RevisionAttribute = { name: string; id: number };

const revisionAttribute = (element: XmlElement): RevisionAttribute | null => {
  if (
    !element.name ||
    !WORDPROCESSINGML_NAMESPACE_URIS.has(getNamespaceUri(element) ?? "") ||
    !REVISION_ELEMENT_NAMES.has(getLocalName(element.name))
  ) {
    return null;
  }
  const attribute = findAttributeByNamespaceUri(element, WORDPROCESSINGML_NAMESPACE_URIS, "id");
  if (!attribute) {
    return null;
  }
  const id = Number(attribute.value);
  return Number.isSafeInteger(id) && id >= 0 ? { name: attribute.name, id } : null;
};

/**
 * Keep physical tracked-change element ids unique across a package.
 *
 * Live editor marks and operation receipts keep their logical revision IDs.
 * Saving assigns fresh IDs only where one logical change was split into
 * multiple physical OOXML wrappers; reopening therefore exposes the physical
 * wrapper IDs that the file format requires.
 */
export const normalizeRevisionIdsInXmlParts = (
  parts: ReadonlyMap<string, string>,
): Map<string, string> => {
  const candidates = [...parts].filter(([, xml]) => REVISION_ELEMENT_CANDIDATE.test(xml));
  const occurrencesByPath = new Map<string, number[]>();
  const reserved = new Set<number>();
  for (const [path, xml] of candidates) {
    assertXmlResourceLimits(xml);
    const ids: number[] = [];
    const scanned = rewriteStreamingXmlDecimalAttributes(xml, (element) => {
      const attribute = revisionAttribute(element);
      if (attribute) {
        ids.push(attribute.id);
        reserved.add(attribute.id);
      }
      return null;
    });
    if (scanned.status === "unsupported") {
      throw new XmlResourceLimitError({
        message: `Revision-id normalization could not safely scan ${path}`,
        limit: "syntax",
      });
    }
    occurrencesByPath.set(path, ids);
  }

  const repeatedPaths = new Set<string>();
  const firstSeen = new Set<number>();
  for (const [path, ids] of occurrencesByPath) {
    for (const id of ids) {
      if (firstSeen.has(id)) {
        repeatedPaths.add(path);
      } else {
        firstSeen.add(id);
      }
    }
  }

  let nextId = 0;
  const allocate = (): number => {
    while (reserved.has(nextId)) {
      nextId += 1;
    }
    const allocated = nextId;
    reserved.add(allocated);
    nextId += 1;
    return allocated;
  };

  const seen = new Set<number>();
  const normalized = new Map(parts);
  for (const [path, xml] of candidates) {
    const ids = occurrencesByPath.get(path);
    if (!ids) {
      continue;
    }
    // Every id this pass lets stand or mints goes through `claim`, so a shape
    // the branches below did not anticipate surfaces as a typed error instead
    // of a package carrying two revisions under one id.
    const claim = (id: number): void => {
      if (seen.has(id)) {
        throw new RevisionIdCollisionError({
          message: `Revision id ${String(id)} is claimed twice in ${path}`,
          revisionId: id,
          part: path,
        });
      }
      seen.add(id);
    };
    if (!repeatedPaths.has(path) && ids.every((id) => !seen.has(id))) {
      for (const id of ids) {
        claim(id);
      }
      continue;
    }
    const rewritten = rewriteStreamingXmlDecimalAttributes(xml, (element) => {
      const attribute = revisionAttribute(element);
      if (!attribute) {
        return null;
      }
      if (!seen.has(attribute.id)) {
        claim(attribute.id);
        return null;
      }
      const replacement = allocate();
      claim(replacement);
      return new Map([[attribute.name, String(replacement)]]);
    });
    if (rewritten.status === "unsupported") {
      throw new XmlResourceLimitError({
        message: `Revision-id normalization could not safely rewrite ${path}`,
        limit: "syntax",
      });
    }
    normalized.set(path, rewritten.value);
  }
  return normalized;
};
