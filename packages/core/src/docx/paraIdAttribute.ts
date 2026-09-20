/**
 * The one reader for "the paraId written on this element".
 *
 * A `paraId` is a join key: `w15:commentEx/@w15:paraId`,
 * `w16cex:comment/@w16cex:paraId` and `w15:paraIdParent` all name a paragraph
 * through the `w14:paraId` it carries. Three Word generations write the same
 * key under three prefixes, and a producer is free to bind any of those
 * namespaces to a prefix of its own choosing, so the key must be resolved by
 * namespace URI. Resolving it by local name alone is worse than missing it: an
 * unrelated `vendor:paraId` would key a comment into another thread and carry
 * that thread's date, parent and resolved state.
 *
 * Every reader of the key calls this, so no call site can come to disagree
 * with another about which attribute is a paraId.
 */

import { OOXML_NAMESPACES } from "./serializer/partNamespaces";
import {
  NAMESPACES,
  WORDPROCESSINGML_NAMESPACE_URIS,
  getAttributeByNamespaceUri,
  type XmlElement,
} from "./xmlParser";

/**
 * The namespaces a paraId-valued attribute may be written in: the Word 2010,
 * 2012 and 2018 wordml extensions.
 *
 * The 2018 URI is read from the serializer's table rather than retyped, so the
 * spelling folio writes and the spelling it accepts cannot drift.
 */
export const PARA_ID_NAMESPACE_URIS: ReadonlySet<string> = new Set([
  NAMESPACES.w14,
  NAMESPACES.w15,
  OOXML_NAMESPACES.w16cex.uri,
]);

/**
 * The conventional prefixes, read literally as a last resort.
 *
 * A part that writes `w14:paraId` without declaring `xmlns:w14` is malformed,
 * and no URI lookup can resolve it; Word reads those files, so folio does too.
 */
const CONVENTIONAL_PARA_ID_PREFIXES = ["w14", "w15", "w16cex"] as const;

const paraIdValuedAttribute = (element: XmlElement, localName: string): string | undefined => {
  const byUri = getAttributeByNamespaceUri(element, PARA_ID_NAMESPACE_URIS, localName);
  if (byUri !== null) {
    return byUri;
  }

  for (const prefix of CONVENTIONAL_PARA_ID_PREFIXES) {
    const literal = element.attributes?.[`${prefix}:${localName}`];
    if (literal !== undefined) {
      return String(literal);
    }
  }

  const inWordprocessingml = getAttributeByNamespaceUri(
    element,
    WORDPROCESSINGML_NAMESPACE_URIS,
    localName,
  );
  if (inWordprocessingml !== null) {
    return inWordprocessingml;
  }

  const literalW = element.attributes?.[`w:${localName}`];
  return literalW === undefined ? undefined : String(literalW);
};

/** The paraId this element carries, or `undefined` when it carries none. */
export const paraIdAttribute = (element: XmlElement): string | undefined =>
  paraIdValuedAttribute(element, "paraId");

/** The paraId of the parent this element names, by the same rule. */
export const paraIdParentAttribute = (element: XmlElement): string | undefined =>
  paraIdValuedAttribute(element, "paraIdParent");

/**
 * The textId this element carries: a paragraph identity written beside
 * `paraId`, in the same namespaces and by the same rule.
 */
export const textIdAttribute = (element: XmlElement): string | undefined =>
  paraIdValuedAttribute(element, "textId");

/**
 * How a save addresses one paragraph of the part it is patching.
 *
 * Word 2010+ writes a `w14:paraId` on every paragraph; LibreOffice, Google
 * Docs, python-docx and docx4j write none, and a package can carry both (Word
 * stamps what it rewrites and leaves the rest). folio mints an id for every
 * paragraph that arrives without one, so the model always has a key — but a
 * minted key names nothing in the file, and looking it up there answers
 * "absent" for every paragraph at once. Naming the two cases apart is what lets
 * the patcher keep the id lookup where the file has an id and fall back to
 * position only where it does not.
 */
export type ParagraphIdentity =
  /** The source part writes this paraId, so the id locates the paragraph. */
  | { type: "authored"; paraId: string; ordinal: number }
  /** The model's id was minted at parse; the ordinal locates the paragraph. */
  | { type: "minted"; paraId: string; ordinal: number }
  /** Neither side writes an id: nothing can name this paragraph to change it. */
  | { type: "anonymous"; ordinal: number };

/**
 * Every paragraph's identity, plus whether ordinals may be trusted.
 *
 * `ordinalsAligned` is the evidence a `minted` splice needs: the ordinal only
 * locates a paragraph when the source part and the model's serialization agree
 * on the paragraph sequence. Every paraId the source writes is a witness to
 * that, so a package with ids on some paragraphs proves its own alignment, and
 * one with none is aligned vacuously — the count check the caller already runs
 * is then the whole of the evidence.
 */
export type ParagraphIdentityPlan = {
  identities: readonly ParagraphIdentity[];
  ordinalsAligned: boolean;
};

export type ResolveParagraphIdentitiesOptions = {
  /** The paraId written on each `<w:p>` of the source part, in document order. */
  sourceParaIds: readonly (string | undefined)[];
  /** The paraId the model serialized for each `<w:p>`, in the same order. */
  serializedParaIds: readonly (string | undefined)[];
};

/**
 * Decide each serialized paragraph's identity against the source part.
 *
 * A paraId the source writes exactly once is `authored`; the same id written
 * twice stays `authored` and is left for the caller's ambiguity check, so a
 * duplicate is refused rather than silently addressed by position. An id the
 * source does not write at all is `minted`.
 */
export const resolveParagraphIdentities = ({
  sourceParaIds,
  serializedParaIds,
}: ResolveParagraphIdentitiesOptions): ParagraphIdentityPlan => {
  const sourceOrdinalById = new Map<string, number>();
  const duplicated = new Set<string>();
  for (const [ordinal, paraId] of sourceParaIds.entries()) {
    if (paraId === undefined) {
      continue;
    }
    if (sourceOrdinalById.has(paraId)) {
      duplicated.add(paraId);
      continue;
    }
    sourceOrdinalById.set(paraId, ordinal);
  }

  let ordinalsAligned = sourceParaIds.length === serializedParaIds.length;
  const identities: ParagraphIdentity[] = [];
  for (const [ordinal, paraId] of serializedParaIds.entries()) {
    const sourceParaId = sourceParaIds[ordinal];
    if (paraId === undefined) {
      identities.push({ type: "anonymous", ordinal });
      ordinalsAligned &&= sourceParaId === undefined;
      continue;
    }
    if (duplicated.has(paraId)) {
      identities.push({ type: "authored", paraId, ordinal });
      ordinalsAligned = false;
      continue;
    }
    const sourceOrdinal = sourceOrdinalById.get(paraId);
    if (sourceOrdinal === undefined) {
      identities.push({ type: "minted", paraId, ordinal });
      // The source names this position under an id the model dropped: the two
      // sequences disagree about what sits here, so no ordinal is evidence.
      ordinalsAligned &&= sourceParaId === undefined;
      continue;
    }
    identities.push({ type: "authored", paraId, ordinal });
    ordinalsAligned &&= sourceOrdinal === ordinal;
  }

  return { identities, ordinalsAligned };
};
