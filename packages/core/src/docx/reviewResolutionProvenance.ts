import {
  panic,
  Result,
  TaggedError,
  type TaggedErrorClass,
} from "better-result";

import { escapeXmlAttribute } from "@stll/docx-core";
import {
  MAX_REVISION_ID,
  MAX_TABLE_GEOMETRY_HISTORY_LENGTH,
  MAX_TABLE_PICTURE_ID,
} from "@stll/docx-core/model";
import {
  type CopyOriginAncestorList,
  isCopyOriginAncestors,
  isKeptCut,
  isKeptCutPiece,
  type KeptCutPiece,
  type VersionOneKept,
  type VersionOneKeptMarkers,
} from "@stll/docx-core/ops";

import { MAX_REVIEW_HISTORY_ATTRIBUTE_LENGTH } from "../reviewHistoryLimits";
import type {
  Insertion,
  KeptCut,
  ParagraphMarkChange,
  ParagraphPropertyChange,
  RecordField,
  RunPropertyChange,
  TableGridChange,
} from "../types/document";
import { FOLIO_REVIEW_HISTORY_NAMESPACE } from "./reviewHistoryNamespace";
import { getAttributeByNamespaceUri, type XmlElement } from "./xmlParser";

const NAMESPACES: ReadonlySet<string> = new Set([
  FOLIO_REVIEW_HISTORY_NAMESPACE,
]);
/**
 * The review history version folio writes. Version 2 adds the revision a
 * paragraph property change belongs to (`folio:resolvesWith`) and records a
 * kept range marker's cut on the marker (`folio:keptCut`). Version 1 files
 * carry no such link and no marker cut: their changes name the kept markers
 * instead, which the reader translates onto the markers
 * ({@link versionOneKeptMarkers}).
 */
export const REVIEW_HISTORY_VERSION = 2;
/** Every version the reader understands. */
const READABLE_VERSIONS: ReadonlySet<unknown> = new Set([
  1,
  REVIEW_HISTORY_VERSION,
]);
type ResolutionJoins = NonNullable<Insertion["resolutionJoins"]>;
type DeferredRemoval = NonNullable<ResolutionJoins["deferredRemove"]>[number];
type BoundaryJoins = NonNullable<RunPropertyChange["boundaryJoins"]>;
type RetainedIdentity = NonNullable<ResolutionJoins["retainedAfter"]>[number];
type IdentitySlot = RetainedIdentity["source"][number];
const SLOT_FIELDS = { space: true, id: true } as const satisfies Record<
  RecordField<IdentitySlot>,
  true
>;
const RETAINED_FIELDS = {
  depth: true,
  source: true,
  target: true,
} as const satisfies Record<RecordField<RetainedIdentity>, true>;
const JOIN_FIELDS = {
  before: true,
  after: true,
  remove: true,
  retainedAfter: true,
  deferredRemove: true,
  acceptance: true,
  plainSides: true,
  emptiedHolder: true,
  holdsParagraph: true,
  holdsControlText: true,
  closedHolder: true,
  retainedEnclosing: true,
} as const satisfies Record<RecordField<ResolutionJoins>, true>;
type EnclosingHandoff = NonNullable<
  ResolutionJoins["retainedEnclosing"]
>[number];
const ENCLOSING_FIELDS = {
  levels: true,
  source: true,
  target: true,
} as const satisfies Record<RecordField<EnclosingHandoff>, true>;
/**
 * Saved as its own attribute: a reader that does not know it keeps the
 * joins and drops only the handoff.
 */
const ENCLOSING_ATTRIBUTE = "retainedEnclosing";

const ACCEPTANCE_POLICIES = {
  "merge-plain-runs": true,
} as const satisfies Record<NonNullable<ResolutionJoins["acceptance"]>, true>;

const HELD_PARAGRAPH_JOINS = {
  "own-joins": true,
  "no-joins": true,
} as const satisfies Record<
  NonNullable<ResolutionJoins["holdsParagraph"]>,
  true
>;

const PARAGRAPH_MARK_FIELDS = {
  kind: { type: "ooxml" },
  info: { type: "ooxml" },
  resolutionJoin: { type: "private", attribute: "resolutionJoin" },
  // Which paragraph rejecting an inserted mark keeps. Its own attribute, not
  // a field of another: a reader that does not know it drops only the
  // marker, and joins the paragraphs as any rejected mark.
  survivor: { type: "private", attribute: "rejectionSurvivor" },
  takenOverMark: { type: "private", attribute: "rejectionSurvivor" },
  // When the split whose cut a deleted mark holds was made. Its own
  // attribute: a reader that does not know it joins at the deletion's date.
  splitDate: { type: "private", attribute: "splitDate" },
  // A later deletion of an added mark is its own w:del beside it.
  deletion: { type: "ooxml" },
} as const satisfies Record<
  RecordField<ParagraphMarkChange>,
  { type: "ooxml" } | { type: "private"; attribute: string }
>;

const DEFERRED_REMOVAL_FIELDS = {
  depth: true,
  blockers: true,
  edge: true,
} as const satisfies Record<RecordField<DeferredRemoval>, true>;

const record = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);
const natural = (value: unknown): value is number =>
  typeof value === "number" && Number.isSafeInteger(value) && value >= 0;
const slots = (value: unknown): value is readonly IdentitySlot[] =>
  Array.isArray(value) &&
  value.every(
    (slot: unknown) =>
      record(slot) &&
      Object.keys(slot).every((key) => Object.hasOwn(SLOT_FIELDS, key)) &&
      (slot["space"] === "revision" || slot["space"] === "control") &&
      natural(slot["id"]) &&
      slot["id"] <= MAX_REVISION_ID,
  );
const retained = (value: unknown): value is readonly RetainedIdentity[] =>
  Array.isArray(value) &&
  value.every((item: unknown) => {
    if (!record(item)) return false;
    const source = item["source"];
    const target = item["target"];
    return (
      Object.keys(item).every((key) => Object.hasOwn(RETAINED_FIELDS, key)) &&
      natural(item["depth"]) &&
      slots(source) &&
      slots(target) &&
      source.length === target.length &&
      source.every((slot, index) => slot.space === target.at(index)?.space)
    );
  });

const enclosingHandoffs = (
  value: unknown,
): value is readonly EnclosingHandoff[] =>
  Array.isArray(value) &&
  value.length > 0 &&
  value.every((item: unknown) => {
    if (!record(item)) return false;
    const levels = item["levels"];
    const source = item["source"];
    const target = item["target"];
    return (
      Object.keys(item).every((key) => Object.hasOwn(ENCLOSING_FIELDS, key)) &&
      natural(levels) &&
      levels >= 1 &&
      levels <= MAX_REVISION_ID &&
      slots(source) &&
      slots(target) &&
      source.length > 0 &&
      source.length === target.length &&
      source.every((slot, index) => slot.space === target.at(index)?.space)
    );
  });

/** What tells two deferred seam groups apart. */
const deferredRemovalKey = ({
  depth,
  blockers,
  edge,
}: DeferredRemoval): string => JSON.stringify({ depth, blockers, edge });

const deferredRemoval = (value: unknown): value is DeferredRemoval => {
  if (
    !record(value) ||
    Object.keys(value).some(
      (key) => !Object.hasOwn(DEFERRED_REMOVAL_FIELDS, key),
    )
  )
    return false;
  const depth = value["depth"];
  const blockers = value["blockers"];
  return (
    (!Object.hasOwn(value, "edge") ||
      value["edge"] === "start" ||
      value["edge"] === "end") &&
    natural(depth) &&
    depth <= MAX_REVISION_ID &&
    Array.isArray(blockers) &&
    blockers.length > 0 &&
    !blockers.includes(undefined) &&
    new Set(blockers).size === blockers.length &&
    blockers.every(
      (blocker: unknown) => natural(blocker) && blocker <= MAX_REVISION_ID,
    )
  );
};

const deferredRemovals = (
  value: unknown,
): value is NonNullable<ResolutionJoins["deferredRemove"]> => {
  if (
    !Array.isArray(value) ||
    value.length === 0 ||
    value.includes(undefined) ||
    !value.every(deferredRemoval)
  )
    return false;
  const groups = value.map(deferredRemovalKey);
  return new Set(groups).size === groups.length;
};

export const isResolutionJoins = (value: unknown): value is ResolutionJoins =>
  record(value) &&
  Object.keys(value).every((key) => Object.hasOwn(JOIN_FIELDS, key)) &&
  natural(value["before"]) &&
  natural(value["after"]) &&
  natural(value["remove"]) &&
  (!Object.hasOwn(value, "acceptance") ||
    (typeof value["acceptance"] === "string" &&
      Object.hasOwn(ACCEPTANCE_POLICIES, value["acceptance"]))) &&
  (value["retainedAfter"] === undefined || retained(value["retainedAfter"])) &&
  (!Object.hasOwn(value, "deferredRemove") ||
    deferredRemovals(value["deferredRemove"])) &&
  (!Object.hasOwn(value, "plainSides") ||
    isBoundaryJoins(value["plainSides"])) &&
  (!Object.hasOwn(value, "emptiedHolder") || value["emptiedHolder"] === true) &&
  (!Object.hasOwn(value, "holdsParagraph") ||
    (typeof value["holdsParagraph"] === "string" &&
      Object.hasOwn(HELD_PARAGRAPH_JOINS, value["holdsParagraph"]))) &&
  (!Object.hasOwn(value, "holdsControlText") ||
    value["holdsControlText"] === true) &&
  (!Object.hasOwn(value, "closedHolder") ||
    isBoundaryJoins(value["closedHolder"])) &&
  (!Object.hasOwn(value, ENCLOSING_ATTRIBUTE) ||
    enclosingHandoffs(value[ENCLOSING_ATTRIBUTE]));

/** The joins as `folio:resolutionJoins` holds them: without the handoff. */
const isStoredResolutionJoins = (value: unknown): value is ResolutionJoins =>
  isResolutionJoins(value) && !Object.hasOwn(value, ENCLOSING_ATTRIBUTE);

const storedJoins = (joins: ResolutionJoins): ResolutionJoins => {
  const { retainedEnclosing: _handoff, ...stored } = joins;
  return stored;
};

const withEnclosing = (
  joins: ResolutionJoins,
  node: XmlElement,
): ResolutionJoins => {
  const handoffs = parseAttribute({
    node,
    attribute: ENCLOSING_ATTRIBUTE,
    valid: enclosingHandoffs,
  });
  return handoffs === undefined
    ? joins
    : { ...joins, retainedEnclosing: handoffs };
};

const serializeEnclosing = (joins: ResolutionJoins | undefined): string =>
  serializeAttribute({
    attribute: ENCLOSING_ATTRIBUTE,
    value: joins?.retainedEnclosing,
    valid: enclosingHandoffs,
  });

export const isParagraphMarkResolutionJoin = (
  value: unknown,
): value is NonNullable<ParagraphMarkChange["resolutionJoin"]> =>
  natural(value) && value <= MAX_REVISION_ID;

const isEncodedParagraphMarkResolutionJoin = (
  value: unknown,
): value is number | null =>
  value === null || isParagraphMarkResolutionJoin(value);

export const isBoundaryJoins = (value: unknown): value is BoundaryJoins =>
  Array.isArray(value) &&
  value.length <= 2 &&
  new Set(value).size === value.length &&
  value.every((side: unknown) => side === "before" || side === "after");

const ReviewResolutionProvenanceErrorBase: TaggedErrorClass<"ReviewResolutionProvenanceError"> =
  TaggedError("ReviewResolutionProvenanceError");

export class ReviewResolutionProvenanceError extends ReviewResolutionProvenanceErrorBase<{
  message: string;
  attribute: string;
  reason: "invalid" | "unsupportedVersion" | "unrepresentable";
}> {}

type ParseAttributeOptions<Value> = {
  node: XmlElement;
  attribute: string;
  valid: (value: unknown) => value is Value;
};

/** Why a provenance attribute's text cannot be read, or the value it holds. */
type DecodedProvenance =
  | { status: "read"; value: unknown; version: unknown }
  | { status: "dropped"; reason: "invalid" | "unsupportedVersion" };

/** Whether `value` has the shape the history `version` writes. */
type ProvenanceShape = (value: unknown, version: unknown) => boolean;

const decodeProvenance = (
  encoded: string,
  valid: ProvenanceShape,
): DecodedProvenance => {
  if (encoded.length > MAX_REVIEW_HISTORY_ATTRIBUTE_LENGTH)
    return { status: "dropped", reason: "invalid" };
  const parsed = Result.try({
    try: (): unknown => JSON.parse(encoded),
    catch: () => undefined,
  });
  if (parsed.isErr()) return { status: "dropped", reason: "invalid" };
  const decoded = parsed.value;
  if (
    !record(decoded) ||
    Object.keys(decoded).some((key) => key !== "version" && key !== "value")
  )
    return { status: "dropped", reason: "invalid" };
  const version = decoded["version"];
  if (!READABLE_VERSIONS.has(version))
    return { status: "dropped", reason: "unsupportedVersion" };
  const value = decoded["value"];
  if (!valid(value, version)) return { status: "dropped", reason: "invalid" };
  return { status: "read", value, version };
};

/**
 * Read one provenance attribute. Provenance is advisory, ignorable metadata:
 * a value folio cannot read (another version, a field it does not know, a
 * malformed value) is dropped and the revisions it would annotate are read
 * as they are; {@link reviewProvenanceIssues} reports what was dropped.
 * What folio writes is validated strictly before it is written.
 */
const parseAttribute = <Value>({
  node,
  attribute,
  valid,
}: ParseAttributeOptions<Value>): Value | undefined => {
  const encoded = getAttributeByNamespaceUri(node, NAMESPACES, attribute);
  if (encoded === null) return undefined;
  const decoded = decodeProvenance(encoded, valid);
  // SAFETY: `valid` narrowed the value before it was read
  return decoded.status === "read" ? (decoded.value as Value) : undefined;
};

type SerializeAttributeOptions = {
  attribute: string;
  value: unknown;
  valid: (value: unknown) => boolean;
};
const encodeAttribute = ({
  attribute,
  value,
  valid,
}: SerializeAttributeOptions): string | undefined => {
  if (value === undefined) return undefined;
  if (!valid(value)) {
    throw new ReviewResolutionProvenanceError({
      attribute,
      reason: "invalid",
      message: `Invalid or oversized ${attribute} review provenance.`,
    });
  }
  const encoded = JSON.stringify({ version: REVIEW_HISTORY_VERSION, value });
  if (encoded.length > MAX_REVIEW_HISTORY_ATTRIBUTE_LENGTH) {
    throw new ReviewResolutionProvenanceError({
      attribute,
      reason: "invalid",
      message: `Oversized ${attribute} review provenance.`,
    });
  }
  return encoded;
};

const serializeAttribute = (options: SerializeAttributeOptions): string => {
  const encoded = encodeAttribute(options);
  return encoded === undefined
    ? ""
    : ` folio:${options.attribute}="${escapeXmlAttribute(encoded)}"`;
};

/**
 * Version 1 history, whose kept markers carry no cut, named them on the
 * change instead (`keptFrom`, `keptCount`, on its joins and on each deferred
 * seam group). This is the one place those fields are read: they are taken
 * off the joins as read, and the paragraph reader marks the markers they
 * name as later versions do ({@link versionOneKeptMarkers}). Later versions
 * never write them.
 */
type VersionOneJoins = { joins: ResolutionJoins; kept: VersionOneKept };

/**
 * `value` without the kept marker fields version 1 wrote on it, and what
 * they say; `undefined` when they are malformed.
 */
const withoutVersionOneKept = (
  value: Record<string, unknown>,
):
  | { rest: Record<string, unknown>; kept: VersionOneKeptMarkers | undefined }
  | undefined => {
  const { keptFrom, keptCount, ...rest } = value;
  if (!Object.hasOwn(value, "keptFrom") && !Object.hasOwn(value, "keptCount"))
    return { rest, kept: undefined };
  if (typeof keptFrom !== "string") return undefined;
  if (!Object.hasOwn(value, "keptCount")) return { rest, kept: { keptFrom } };
  return natural(keptCount) && keptCount >= 1 && keptCount <= MAX_REVISION_ID
    ? { rest, kept: { keptFrom, keptCount } }
    : undefined;
};

/** Version 1 joins as the model holds them, and what they said of kept markers. */
const readVersionOneJoins = (value: unknown): VersionOneJoins | undefined => {
  if (!record(value)) return undefined;
  const change = withoutVersionOneKept(value);
  if (change === undefined) return undefined;
  const { deferredRemove, ...joins } = change.rest;
  const groups: (VersionOneKeptMarkers | undefined)[] = [];
  if (Object.hasOwn(change.rest, "deferredRemove")) {
    if (!Array.isArray(deferredRemove)) return undefined;
    const stripped: unknown[] = [];
    const written = new Set<string>();
    const seen = new Set<string>();
    for (const group of deferredRemove) {
      const read = record(group) ? withoutVersionOneKept(group) : undefined;
      if (read === undefined || !deferredRemoval(read.rest)) return undefined;
      // Version 1 wrote each group once.
      const asWritten = JSON.stringify([
        deferredRemovalKey(read.rest),
        read.kept,
      ]);
      if (written.has(asWritten)) return undefined;
      written.add(asWritten);
      // Groups only kept dates told apart are one seam: the first names its
      // markers.
      const key = deferredRemovalKey(read.rest);
      if (seen.has(key)) continue;
      seen.add(key);
      stripped.push(read.rest);
      groups.push(read.kept);
    }
    joins["deferredRemove"] = stripped;
  }
  if (!isStoredResolutionJoins(joins)) return undefined;
  return {
    joins,
    kept:
      change.kept === undefined ? { groups } : { change: change.kept, groups },
  };
};

/** The joins as `version` of the history writes them. */
const isStoredJoinsOf: ProvenanceShape = (value, version) =>
  version === 1
    ? readVersionOneJoins(value) !== undefined
    : isStoredResolutionJoins(value);

/** What version 1 history said of kept markers, by the joins read from it. */
const VERSION_ONE_KEPT = new WeakMap<ResolutionJoins, VersionOneKept>();

/** `joins` decoded from `version` of the history, as the model holds them. */
const modelJoins = (joins: unknown, version: unknown): ResolutionJoins => {
  if (version !== 1) {
    if (!isStoredResolutionJoins(joins))
      return panic("Decoded resolution joins have their version's shape.");
    return joins;
  }
  const read =
    readVersionOneJoins(joins) ??
    panic("Decoded version 1 joins have their version's shape.");
  VERSION_ONE_KEPT.set(read.joins, read.kept);
  return read.joins;
};

/** `joins` with their handoff, keeping what version 1 said of kept markers. */
const withEnclosingKept = (
  joins: ResolutionJoins,
  node: XmlElement,
): ResolutionJoins => {
  const enclosed = withEnclosing(joins, node);
  const kept = VERSION_ONE_KEPT.get(joins);
  if (kept !== undefined) VERSION_ONE_KEPT.set(enclosed, kept);
  return enclosed;
};

/** Decode a provenance attribute of `node` whose shape depends on its version. */
const decodeAttribute = (
  node: XmlElement,
  attribute: string,
  valid: ProvenanceShape,
): { value: unknown; version: unknown } | undefined => {
  const encoded = getAttributeByNamespaceUri(node, NAMESPACES, attribute);
  if (encoded === null) return undefined;
  const decoded = decodeProvenance(encoded, valid);
  return decoded.status === "read" ? decoded : undefined;
};

export const parseResolutionJoins = (
  node: XmlElement,
): ResolutionJoins | undefined => {
  const decoded = decodeAttribute(node, "resolutionJoins", isStoredJoinsOf);
  return decoded === undefined
    ? undefined
    : withEnclosingKept(modelJoins(decoded.value, decoded.version), node);
};

/**
 * What version 1 history said of the kept markers `joins` name, when they
 * were read from it.
 */
export const versionOneKeptMarkers = (
  joins: ResolutionJoins,
): VersionOneKept | undefined => VERSION_ONE_KEPT.get(joins);
export const parseBoundaryJoins = (
  node: XmlElement,
): BoundaryJoins | undefined =>
  parseAttribute({ node, attribute: "boundaryJoins", valid: isBoundaryJoins });
export const parseCopyOriginAncestors = (
  node: XmlElement,
): CopyOriginAncestorList | undefined =>
  parseAttribute({
    node,
    attribute: "copyOriginAncestors",
    valid: isCopyOriginAncestors,
  });
/**
 * A kept cut a file holds, without the `depth` that files saved before the
 * cut lost it carry, whatever it says.
 */
const storedKeptCut = (value: unknown): unknown => {
  if (!record(value)) return value;
  const { depth: _ignored, ...cut } = value;
  return cut;
};
// Older marker history recorded the holder's side. Read that form and
// bind it once the paragraph is assembled. The writer emits holder IDs only.
const LEGACY_KEPT_CUT_PIECES = new WeakMap<KeptCut, KeptCutPiece>();

export const legacyKeptCutPiece = (cut: KeptCut): KeptCutPiece | undefined =>
  LEGACY_KEPT_CUT_PIECES.get(cut);

const legacyStoredKeptCut = (value: unknown) => {
  if (!record(value)) return undefined;
  const { piece, ...cut } = value;
  return isKeptCutPiece(piece) && isKeptCut(cut) && cut.holders === undefined
    ? { cut, piece }
    : undefined;
};

const isStoredKeptCut = (value: unknown): boolean => {
  const cut = storedKeptCut(value);
  return isKeptCut(cut) || legacyStoredKeptCut(cut) !== undefined;
};

/** The cut a kept range marker stands in (`folio:keptCut`). */
export const parseKeptCut = (node: XmlElement): KeptCut | undefined => {
  const decoded = decodeAttribute(node, "keptCut", isStoredKeptCut);
  const cut = storedKeptCut(decoded?.value);
  if (isKeptCut(cut)) return cut;
  const legacy = legacyStoredKeptCut(cut);
  if (legacy === undefined) return undefined;
  LEGACY_KEPT_CUT_PIECES.set(legacy.cut, legacy.piece);
  return legacy.cut;
};
export const serializeKeptCut = (
  value: KeptCut | undefined,
): string | undefined =>
  encodeAttribute({ attribute: "keptCut", value, valid: isKeptCut });
export const serializeResolutionJoins = (
  value: Insertion["resolutionJoins"],
): string =>
  serializeAttribute({
    attribute: "resolutionJoins",
    value: value === undefined ? undefined : storedJoins(value),
    valid: isStoredResolutionJoins,
  }) + serializeEnclosing(value);
export const serializeBoundaryJoins = (
  value: RunPropertyChange["boundaryJoins"],
): string =>
  serializeAttribute({
    attribute: "boundaryJoins",
    value,
    valid: isBoundaryJoins,
  });
export const serializeCopyOriginAncestors = (
  value: CopyOriginAncestorList | undefined,
): string | undefined =>
  encodeAttribute({
    attribute: "copyOriginAncestors",
    value,
    valid: isCopyOriginAncestors,
  });

type StoredRejectionSurvivor = { takenOverMark: number | null };

const isStoredRejectionSurvivor = (
  value: unknown,
): value is StoredRejectionSurvivor =>
  record(value) &&
  Object.keys(value).length === 1 &&
  (value["takenOverMark"] === null || isRevisionId(value["takenOverMark"]));

const isParagraphMarkSurvivor = (
  value: unknown,
): value is "first" | StoredRejectionSurvivor =>
  value === "first" || isStoredRejectionSurvivor(value);

const isParagraphMarkSplitDate = (
  value: unknown,
): value is NonNullable<ParagraphMarkChange["splitDate"]> =>
  typeof value === "string";

/**
 * A paragraph mark's private review provenance: its rejection join depth
 * (null carries an owned undefined depth), the paragraph rejecting it keeps,
 * and when the split whose cut a deleted mark holds was made. A missing
 * attribute carries no field.
 */
export const parseParagraphMarkResolutionJoin = (
  node: XmlElement,
): Pick<
  ParagraphMarkChange,
  "resolutionJoin" | "survivor" | "takenOverMark" | "splitDate"
> => {
  const value = parseAttribute({
    node,
    attribute: PARAGRAPH_MARK_FIELDS.resolutionJoin.attribute,
    valid: isEncodedParagraphMarkResolutionJoin,
  });
  const survivor = parseAttribute({
    node,
    attribute: PARAGRAPH_MARK_FIELDS.survivor.attribute,
    valid: isParagraphMarkSurvivor,
  });
  const splitDate = parseAttribute({
    node,
    attribute: PARAGRAPH_MARK_FIELDS.splitDate.attribute,
    valid: isParagraphMarkSplitDate,
  });
  return {
    ...(value === undefined
      ? {}
      : { resolutionJoin: value === null ? undefined : value }),
    // Older saved documents use the positional string. All saves use the
    // object envelope, with null for a break without a revision id.
    ...(survivor === undefined ? {} : { survivor: "first" as const }),
    ...(survivor === undefined ||
    survivor === "first" ||
    survivor.takenOverMark === null
      ? {}
      : { takenOverMark: survivor.takenOverMark }),
    ...(splitDate === undefined ? {} : { splitDate }),
  };
};

const paragraphMarkSurvivorOptions = (
  mark: ParagraphMarkChange,
): SerializeAttributeOptions => ({
  attribute: PARAGRAPH_MARK_FIELDS.survivor.attribute,
  value:
    mark.survivor === undefined
      ? undefined
      : { takenOverMark: mark.takenOverMark ?? null },
  valid: isStoredRejectionSurvivor,
});

const paragraphMarkSplitDateOptions = (
  mark: ParagraphMarkChange,
): SerializeAttributeOptions => ({
  attribute: PARAGRAPH_MARK_FIELDS.splitDate.attribute,
  value: mark.splitDate,
  valid: isParagraphMarkSplitDate,
});

const paragraphMarkResolutionJoinOptions = (
  mark: ParagraphMarkChange,
): SerializeAttributeOptions => ({
  attribute: PARAGRAPH_MARK_FIELDS.resolutionJoin.attribute,
  value: Object.hasOwn(mark, "resolutionJoin")
    ? (mark.resolutionJoin ?? null)
    : undefined,
  valid: isEncodedParagraphMarkResolutionJoin,
});

export const serializeParagraphMarkResolutionJoin = (
  mark: ParagraphMarkChange,
): string =>
  serializeAttribute(paragraphMarkResolutionJoinOptions(mark)) +
  serializeAttribute(paragraphMarkSurvivorOptions(mark)) +
  serializeAttribute(paragraphMarkSplitDateOptions(mark));

/** Captured paragraph properties compose the same envelopes into a parsed XML element. */
export const paragraphMarkResolutionJoinAttributes = (
  mark: ParagraphMarkChange,
): Record<string, string> => {
  const attributes: Record<string, string> = {};
  for (const options of [
    paragraphMarkResolutionJoinOptions(mark),
    paragraphMarkSurvivorOptions(mark),
    paragraphMarkSplitDateOptions(mark),
  ]) {
    const encoded = encodeAttribute(options);
    if (encoded !== undefined)
      attributes[`folio:${options.attribute}`] = encoded;
  }
  return attributes;
};

/**
 * One piece of a tracked change written around a lifted hyperlink: a revision that
 * holds a hyperlink is written with the revision inside the link, so a
 * change spanning text and a link becomes one element per stretch. The first
 * piece carries the whole change's resolution joins; every piece says which
 * one it is, so a reader rebuilds the one change, and a change nested in it
 * can be rebuilt in turn. It is its own attribute: a reader that does not know
 * it drops the joins and keeps the revisions.
 * A single piece records lifted ancestry even when no element was split.
 * Older readers reject that positive-count extension and keep ordinary OOXML.
 */
export type SplitResolutionPiece = {
  piece: number;
  pieces: number;
  joins?: ResolutionJoins;
};

const SPLIT_PIECE_FIELDS = {
  piece: true,
  pieces: true,
  joins: true,
} as const satisfies Record<keyof SplitResolutionPiece, true>;

export const isSplitResolutionPiece = (
  value: unknown,
): value is SplitResolutionPiece =>
  record(value) &&
  Object.keys(value).every((key) => Object.hasOwn(SPLIT_PIECE_FIELDS, key)) &&
  natural(value["piece"]) &&
  natural(value["pieces"]) &&
  value["pieces"] >= 1 &&
  value["piece"] < value["pieces"] &&
  // Only the first piece carries joins, when the change has them.
  (!Object.hasOwn(value, "joins") ||
    (value["piece"] === 0 && isStoredResolutionJoins(value["joins"])));

/** A split piece as `version` of the history writes it. */
const isSplitPieceOf: ProvenanceShape = (value, version) => {
  if (version !== 1 || !record(value) || !Object.hasOwn(value, "joins"))
    return isSplitResolutionPiece(value);
  return isSplitResolutionPiece({
    ...value,
    joins: readVersionOneJoins(value["joins"])?.joins,
  });
};

export const parseSplitResolutionPiece = (
  node: XmlElement,
): SplitResolutionPiece | undefined => {
  const decoded = decodeAttribute(node, "splitResolutionJoins", isSplitPieceOf);
  if (decoded === undefined) return undefined;
  const { value, version } = decoded;
  if (!record(value) || !Object.hasOwn(value, "joins")) {
    if (!isSplitResolutionPiece(value))
      return panic("A decoded split piece has its version's shape.");
    return value;
  }
  const piece = { ...value, joins: modelJoins(value["joins"], version) };
  if (!isSplitResolutionPiece(piece) || piece.joins === undefined)
    return panic("A decoded split piece has its version's shape.");
  return { ...piece, joins: withEnclosingKept(piece.joins, node) };
};

export const serializeSplitResolutionPiece = (
  value: SplitResolutionPiece,
): string =>
  serializeAttribute({
    attribute: "splitResolutionJoins",
    value:
      value.joins === undefined
        ? value
        : { ...value, joins: storedJoins(value.joins) },
    valid: isSplitResolutionPiece,
  }) + serializeEnclosing(value.joins);

const isRevisionId = (value: unknown): value is number =>
  natural(value) && value <= MAX_REVISION_ID;

/**
 * The revision a paragraph property change belongs to (`folio:resolvesWith`
 * on `w:pPrChange`): a tracked join's paragraph-mark deletion, decided with
 * it. A reader that does not know it resolves the two records separately.
 */
export const parseParagraphPropertyChangeOwner = (
  node: XmlElement,
): Pick<ParagraphPropertyChange, "resolvesWith"> => {
  const resolvesWith = parseAttribute({
    node,
    attribute: "resolvesWith",
    valid: isRevisionId,
  });
  return resolvesWith === undefined ? {} : { resolvesWith };
};
export const serializeParagraphPropertyChangeOwner = (
  change: ParagraphPropertyChange,
): string =>
  serializeAttribute({
    attribute: "resolvesWith",
    value: change.resolvesWith,
    valid: isRevisionId,
  });

const TABLE_GEOMETRY_CHANGE = "tableGeometryChange";
type GeometryChange = NonNullable<TableGridChange["geometryChange"]>;
type TableMove = NonNullable<
  Extract<GeometryChange, { type: "mergeCells" }>["moves"]
>[number];
type NestedTableMove = NonNullable<
  Extract<GeometryChange, { type: "mergeCells" }>["nestedMoves"]
>[number];
type ValidHandoff = (source: number, destination: number) => boolean;

const validIdentityHandoffs = (
  value: unknown,
  valid: ValidHandoff,
): boolean => {
  if (value === undefined) return true;
  if (!Array.isArray(value) || value.length === 0) return false;
  const authored = new Set<number>();
  const temporary = new Set<number>();
  for (const pair of value) {
    if (!Array.isArray(pair) || pair.length !== 2) return false;
    const [source, destination] = pair;
    if (
      typeof source !== "number" ||
      typeof destination !== "number" ||
      !valid(source, destination) ||
      source === destination ||
      authored.has(source) ||
      temporary.has(destination)
    )
      return false;
    authored.add(source);
    temporary.add(destination);
  }
  return [...authored].every((source) => !temporary.has(source));
};
const validControlHandoff: ValidHandoff = (source, destination) =>
  Number.isInteger(source) &&
  source >= -MAX_REVISION_ID - 1 &&
  source <= MAX_REVISION_ID &&
  isRevisionId(destination);
const validPictureId = (id: number) =>
  Number.isInteger(id) && id > 0 && id <= MAX_TABLE_PICTURE_ID;
const validPictureHandoff: ValidHandoff = (source, destination) =>
  validPictureId(source) && validPictureId(destination);
const validRevisionHandoff: ValidHandoff = (source, destination) =>
  isRevisionId(source) && isRevisionId(destination);

const isTableMove = (value: unknown): value is TableMove =>
  record(value) &&
  typeof value["source"] === "string" &&
  typeof value["destination"] === "string" &&
  isRevisionId(value["destinationMark"]) &&
  ((value["kind"] === "empty" &&
    value["controls"] === undefined &&
    value["pictures"] === undefined) ||
    (value["kind"] === undefined &&
      isRevisionId(value["fromRevision"]) &&
      isRevisionId(value["toRevision"]) &&
      isRevisionId(value["fromRange"]) &&
      isRevisionId(value["toRange"]) &&
      validIdentityHandoffs(value["controls"], validControlHandoff) &&
      validIdentityHandoffs(value["pictures"], validPictureHandoff)));
const isNestedTableMove = (value: unknown): value is NestedTableMove => {
  if (!record(value)) return false;
  const rows = value["rows"];
  const paragraphs = value["paragraphs"];
  const omittedTable = value["omittedTable"];
  return (
    Array.isArray(rows) &&
    rows.every(isRevisionId) &&
    new Set(rows).size === rows.length &&
    Array.isArray(paragraphs) &&
    (omittedTable === undefined
      ? rows.length > 0 || paragraphs.length > 0
      : rows.length === 0 &&
        paragraphs.length === 0 &&
        Array.isArray(omittedTable) &&
        omittedTable.length === 2 &&
        omittedTable.every(isRevisionId) &&
        omittedTable[0] !== omittedTable[1]) &&
    paragraphs.every(
      (pair) =>
        Array.isArray(pair) &&
        pair.length === 2 &&
        pair.every((id) => typeof id === "string" && /^[\da-f]{8}$/iu.test(id)),
    ) &&
    validIdentityHandoffs(value["revisions"], validRevisionHandoff) &&
    validIdentityHandoffs(value["controls"], validControlHandoff) &&
    validIdentityHandoffs(value["pictures"], validPictureHandoff)
  );
};
const isTableGeometryChange = (value: unknown): value is GeometryChange => {
  if (
    !record(value) ||
    (value["type"] !== "deleteColumn" && value["type"] !== "mergeCells") ||
    Object.keys(value).some(
      (key) =>
        key !== "type" &&
        key !== "revisionIds" &&
        (value["type"] !== "mergeCells" ||
          (key !== "moves" && key !== "nestedMoves" && key !== "emptyOwner")),
    )
  )
    return false;
  const ids = value["revisionIds"];
  const moves = value["moves"];
  const nestedMoves = value["nestedMoves"];
  const emptyOwner = value["emptyOwner"];
  return (
    Array.isArray(ids) &&
    ids.length > 0 &&
    ids.every(isRevisionId) &&
    new Set(ids).size === ids.length &&
    (moves === undefined ||
      (value["type"] === "mergeCells" &&
        Array.isArray(moves) &&
        moves.every(isTableMove))) &&
    (nestedMoves === undefined ||
      (value["type"] === "mergeCells" &&
        Array.isArray(nestedMoves) &&
        nestedMoves.every(isNestedTableMove) &&
        JSON.stringify(value).length <= MAX_TABLE_GEOMETRY_HISTORY_LENGTH)) &&
    (emptyOwner === undefined ||
      (value["type"] === "mergeCells" && typeof emptyOwner === "string"))
  );
};

export const parseTableGeometryChange = (
  node: XmlElement,
): Pick<TableGridChange, "geometryChange"> => {
  const geometryChange = parseAttribute({
    node,
    attribute: TABLE_GEOMETRY_CHANGE,
    valid: isTableGeometryChange,
  });
  return geometryChange === undefined ? {} : { geometryChange };
};

export const serializeTableGeometryChange = (change: TableGridChange): string =>
  serializeAttribute({
    attribute: TABLE_GEOMETRY_CHANGE,
    value: change.geometryChange,
    valid: isTableGeometryChange,
  });

/** Every provenance attribute folio reads, with the shape its value must have. */
const PROVENANCE_VALUE_SHAPES: Readonly<Record<string, ProvenanceShape>> = {
  resolutionJoins: isStoredJoinsOf,
  boundaryJoins: isBoundaryJoins,
  [PARAGRAPH_MARK_FIELDS.resolutionJoin.attribute]:
    isEncodedParagraphMarkResolutionJoin,
  copyOriginAncestors: isCopyOriginAncestors,
  keptCut: isStoredKeptCut,
  splitResolutionJoins: isSplitPieceOf,
  [ENCLOSING_ATTRIBUTE]: enclosingHandoffs,
  [PARAGRAPH_MARK_FIELDS.survivor.attribute]: isParagraphMarkSurvivor,
  [PARAGRAPH_MARK_FIELDS.splitDate.attribute]: isParagraphMarkSplitDate,
  resolvesWith: isRevisionId,
  [TABLE_GEOMETRY_CHANGE]: isTableGeometryChange,
};

/**
 * Why the reader drops a provenance attribute's value, or `undefined` when it
 * reads it (or does not read that attribute at all).
 */
export const reviewProvenanceIssue = (
  attribute: string,
  encoded: string,
): "invalid" | "unsupportedVersion" | undefined => {
  const shape = PROVENANCE_VALUE_SHAPES[attribute];
  if (shape === undefined) return undefined;
  const decoded = decodeProvenance(encoded, shape);
  return decoded.status === "dropped" ? decoded.reason : undefined;
};
