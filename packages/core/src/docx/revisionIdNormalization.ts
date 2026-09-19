// PARSE-WARNING-EXEMPT: renumbers revision save ids, which carry no authored
// meaning and are regenerated on every save.

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
 * The rest of the annotation id space.
 *
 * A comment, a bookmark, a protected range and a tracked change all draw their
 * `w:id` from one space: Word allocates from a single counter, which is why a
 * package carrying several kinds almost never repeats a value across them. So
 * an id this pass mints must avoid these as well, or a renumbered `w:ins`
 * lands on a live comment.
 *
 * They are only ever reserved, never claimed. A comment id legitimately
 * appears four times (`w:comment`, both range markers and the reference) and a
 * bookmark id twice, so feeding them to the uniqueness machinery would reject
 * a package Word wrote. Their pairing is also why they are not revision
 * elements: renumbering one end of a range would unpair it.
 */
const ANNOTATION_ELEMENT_NAMES = new Set([
  "bookmarkEnd",
  "bookmarkStart",
  "comment",
  "commentRangeEnd",
  "commentRangeStart",
  "commentReference",
  "customXmlDelRangeEnd",
  "customXmlDelRangeStart",
  "customXmlInsRangeEnd",
  "customXmlInsRangeStart",
  "customXmlMoveFromRangeEnd",
  "customXmlMoveFromRangeStart",
  "customXmlMoveToRangeEnd",
  "customXmlMoveToRangeStart",
  "moveFromRangeEnd",
  "moveFromRangeStart",
  "moveToRangeEnd",
  "moveToRangeStart",
  "permEnd",
  "permStart",
]);

const ANNOTATION_ELEMENT_CANDIDATE = new RegExp(
  `<(?:[^\\s<>/:]+:)?(?:${[...ANNOTATION_ELEMENT_NAMES].join("|")})(?:[\\s/>])`,
  "u",
);

const ID_KINDS = { revision: "revision", annotation: "annotation" } as const;

type IdKind = (typeof ID_KINDS)[keyof typeof ID_KINDS];

/** One lookup for both halves of the space, so an element is classified once. */
const ID_KIND_BY_ELEMENT_NAME: ReadonlyMap<string, IdKind> = new Map([
  ...[...REVISION_ELEMENT_NAMES].map((name): [string, IdKind] => [name, ID_KINDS.revision]),
  ...[...ANNOTATION_ELEMENT_NAMES].map((name): [string, IdKind] => [name, ID_KINDS.annotation]),
]);

type IdentifiedElement = { kind: IdKind; name: string; id: number };

/**
 * An element's `w:id` and which half of the annotation space it belongs to.
 *
 * One classifier rather than two, because it runs on every element of every
 * scanned part: resolving the namespace and the local name twice to ask two
 * questions measured 18% on a 17.6 MiB package.
 *
 * `w:permStart` types its id as a string, so a protected range named
 * `everyone` yields nothing. That is correct: a value the allocator can never
 * mint is not one it has to avoid.
 */
const identifiedElement = (element: XmlElement): IdentifiedElement | null => {
  if (!element.name || !WORDPROCESSINGML_NAMESPACE_URIS.has(getNamespaceUri(element) ?? "")) {
    return null;
  }
  const localName = getLocalName(element.name);
  const kind = ID_KIND_BY_ELEMENT_NAME.get(localName);
  if (kind === undefined) {
    return null;
  }
  const attribute = findAttributeByNamespaceUri(element, WORDPROCESSINGML_NAMESPACE_URIS, "id");
  if (!attribute) {
    return null;
  }
  const id = Number(attribute.value);
  return Number.isSafeInteger(id) && id >= 0 ? { kind, name: attribute.name, id } : null;
};

const revisionAttribute = (element: XmlElement): IdentifiedElement | null => {
  const identified = identifiedElement(element);
  return identified?.kind === ID_KINDS.revision ? identified : null;
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
    assertXmlResourceLimits({ xml, partPath: path });
    const ids: number[] = [];
    // One walk collects both: the revision ids this pass owns, and the rest of
    // the annotation space it must not mint into.
    const scanned = rewriteStreamingXmlDecimalAttributes(xml, (element) => {
      const identified = identifiedElement(element);
      if (identified === null) {
        return null;
      }
      if (identified.kind === ID_KINDS.revision) {
        ids.push(identified.id);
      }
      reserved.add(identified.id);
      return null;
    });
    if (scanned.status === "unsupported") {
      throw new XmlResourceLimitError({
        message: `Revision-id normalization could not safely scan ${path}`,
        limit: "syntax",
        observed: 0,
        allowed: 0,
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

  // The parts with no revision element carry annotation ids too, and
  // `word/comments.xml` is the obvious one. Walking them is a second pass over
  // the package, so it is paid only when an id is actually going to be minted:
  // with no repeated revision id every part takes the fast path below and
  // nothing is allocated.
  if (repeatedPaths.size > 0) {
    for (const [path, xml] of parts) {
      if (occurrencesByPath.has(path) || !ANNOTATION_ELEMENT_CANDIDATE.test(xml)) {
        continue;
      }
      assertXmlResourceLimits({ xml, partPath: path });
      const scanned = rewriteStreamingXmlDecimalAttributes(xml, (element) => {
        const identified = identifiedElement(element);
        if (identified !== null) {
          reserved.add(identified.id);
        }
        return null;
      });
      if (scanned.status === "unsupported") {
        throw new XmlResourceLimitError({
          message: `Revision-id normalization could not safely scan ${path}`,
          limit: "syntax",
          observed: 0,
          allowed: 0,
        });
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
        observed: 0,
        allowed: 0,
      });
    }
    normalized.set(path, rewritten.value);
  }
  return normalized;
};
