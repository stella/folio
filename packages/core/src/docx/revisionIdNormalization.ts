import {
  findAttributeByNamespaceUri,
  getLocalName,
  getNamespaceUri,
  type XmlElement,
  WORDPROCESSINGML_NAMESPACE_URIS,
} from "./xmlParser";
import { rewriteStreamingXmlDecimalAttributes } from "./streamingXmlParser";
import { assertXmlResourceLimits, XmlResourceLimitError } from "./xmlResourceLimits";

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
    if (!repeatedPaths.has(path) && ids.every((id) => !seen.has(id))) {
      for (const id of ids) {
        seen.add(id);
      }
      continue;
    }
    const rewritten = rewriteStreamingXmlDecimalAttributes(xml, (element) => {
      const attribute = revisionAttribute(element);
      if (!attribute) {
        return null;
      }
      if (!seen.has(attribute.id)) {
        seen.add(attribute.id);
        return null;
      }
      const replacement = allocate();
      seen.add(replacement);
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
