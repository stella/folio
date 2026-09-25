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
 * What comparing two paragraph sequences by ordinal is worth.
 *
 * An ordinal only locates a paragraph when the two sequences agree on the
 * paragraph order, and the ids they carry are the witnesses to that. The
 * witnesses answer two separate questions, so they are reported separately
 * rather than folded into one boolean: `consistent` says no id contradicts the
 * ordinals, and `sharedOrdinals` counts the ids that actively confirm them. A
 * caller comparing two views of ONE document already knows it is the same
 * document, so absence of contradiction is its whole burden; a caller asking
 * whether two containers of two DIFFERENT documents are the same container
 * needs confirmation as well, or every equal-length container would pair with
 * every other. Deciding that here, once, is what stops the save and the
 * comparison from drifting into two answers about the same paragraphs.
 */
export type ParagraphOrdinalAlignment = {
  /** No id sits at one ordinal on one side and a different one on the other. */
  consistent: boolean;
  /** Ordinals carrying the same id on both sides. */
  sharedOrdinals: number;
};

/** Ordinal per id, or `null` when the sequence writes that id more than once. */
const ordinalsById = (
  identities: readonly ParagraphIdentity[],
): ReadonlyMap<string, number | null> => {
  const byId = new Map<string, number | null>();
  for (const identity of identities) {
    if (identity.type === "anonymous") {
      continue;
    }
    byId.set(identity.paraId, byId.has(identity.paraId) ? null : identity.ordinal);
  }
  return byId;
};

/**
 * Line two paragraph sequences up by ordinal, judged by the ids they carry.
 *
 * An id only testifies when both sides carry it: one side alone says nothing,
 * because a producer's id can be absent from a model that could not read it and
 * a minted id is absent from the file it was minted for. An id both sides carry
 * at the same ordinal confirms the alignment; at different ordinals, or written
 * twice on either side, it refutes it.
 */
export const alignParagraphOrdinals = (
  base: readonly ParagraphIdentity[],
  revised: readonly ParagraphIdentity[],
): ParagraphOrdinalAlignment => {
  if (base.length !== revised.length) {
    return { consistent: false, sharedOrdinals: 0 };
  }
  const baseOrdinals = ordinalsById(base);
  const revisedOrdinals = ordinalsById(revised);
  let sharedOrdinals = 0;
  for (const [paraId, baseOrdinal] of baseOrdinals) {
    const revisedOrdinal = revisedOrdinals.get(paraId);
    if (revisedOrdinal === undefined) {
      continue;
    }
    if (baseOrdinal === null || revisedOrdinal === null || baseOrdinal !== revisedOrdinal) {
      return { consistent: false, sharedOrdinals: 0 };
    }
    sharedOrdinals += 1;
  }
  return { consistent: true, sharedOrdinals };
};

/**
 * Every paragraph's identity, plus whether ordinals may be trusted.
 *
 * `ordinalsAligned` is the evidence a `minted` splice needs. Both sequences
 * describe one document — the source part and the model parsed from it — so
 * {@link alignParagraphOrdinals}'s `consistent` is the whole of the question: a
 * package with ids on some paragraphs proves its own alignment through them,
 * and one with none is aligned exactly when the two sequences have the same
 * length. The save asks this per story (main flow, text boxes), so a text box
 * the model re-reads cannot misalign the main flow.
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
  const sourceIdentities = sourceParaIds.map(
    (paraId, ordinal): ParagraphIdentity =>
      paraId === undefined ? { type: "anonymous", ordinal } : { type: "authored", paraId, ordinal },
  );
  const sourceIds = new Set(sourceParaIds.filter((paraId) => paraId !== undefined));
  const identities = serializedParaIds.map((paraId, ordinal): ParagraphIdentity => {
    if (paraId === undefined) {
      return { type: "anonymous", ordinal };
    }
    return sourceIds.has(paraId)
      ? { type: "authored", paraId, ordinal }
      : { type: "minted", paraId, ordinal };
  });

  return {
    identities,
    ordinalsAligned: alignParagraphOrdinals(sourceIdentities, identities).consistent,
  };
};
