import { Result, TaggedError } from "better-result";
import { escapeXmlAttribute } from "@stll/docx-core";
import { MAX_REVISION_ID } from "@stll/docx-core/model";
import type { Insertion, ParagraphMarkChange, RunPropertyChange } from "../types/document";
import { getAttributeByNamespaceUri, type XmlElement } from "./xmlParser";
import { FOLIO_REVIEW_HISTORY_NAMESPACE } from "./reviewHistoryNamespace";

const NAMESPACES: ReadonlySet<string> = new Set([FOLIO_REVIEW_HISTORY_NAMESPACE]);
const VERSION = 1;
const MAX_ATTRIBUTE_LENGTH = 32_768;
type ResolutionJoins = NonNullable<Insertion["resolutionJoins"]>;
type DeferredRemoval = NonNullable<ResolutionJoins["deferredRemove"]>[number];
type BoundaryJoins = NonNullable<RunPropertyChange["boundaryJoins"]>;
type RetainedIdentity = NonNullable<ResolutionJoins["retainedAfter"]>[number];
type IdentitySlot = RetainedIdentity["source"][number];
const SLOT_FIELDS = { space: true, id: true } as const satisfies Record<keyof IdentitySlot, true>;
const RETAINED_FIELDS = { depth: true, source: true, target: true } as const satisfies Record<
  keyof RetainedIdentity,
  true
>;
const JOIN_FIELDS = {
  before: true,
  after: true,
  remove: true,
  retainedAfter: true,
  deferredRemove: true,
  acceptance: true,
} as const satisfies Record<keyof ResolutionJoins, true>;

const ACCEPTANCE_POLICIES = { "merge-plain-runs": true } as const satisfies Record<
  NonNullable<ResolutionJoins["acceptance"]>,
  true
>;

const PARAGRAPH_MARK_FIELDS = {
  kind: { type: "ooxml" },
  info: { type: "ooxml" },
  resolutionJoin: { type: "private", attribute: "resolutionJoin" },
} as const satisfies Record<
  keyof ParagraphMarkChange,
  { type: "ooxml" } | { type: "private"; attribute: string }
>;

const DEFERRED_REMOVAL_FIELDS = { depth: true, blockers: true } as const satisfies Record<
  keyof DeferredRemoval,
  true
>;

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

const deferredRemoval = (value: unknown): value is DeferredRemoval => {
  if (
    !record(value) ||
    Object.keys(value).some((key) => !Object.hasOwn(DEFERRED_REMOVAL_FIELDS, key))
  )
    return false;
  const depth = value["depth"];
  const blockers = value["blockers"];
  return (
    natural(depth) &&
    depth <= MAX_REVISION_ID &&
    Array.isArray(blockers) &&
    blockers.length > 0 &&
    !blockers.includes(undefined) &&
    new Set(blockers).size === blockers.length &&
    blockers.every((blocker: unknown) => natural(blocker) && blocker <= MAX_REVISION_ID)
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
  const groups = value.map(({ depth, blockers }) => JSON.stringify({ depth, blockers }));
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
  (!Object.hasOwn(value, "deferredRemove") || deferredRemovals(value["deferredRemove"]));

export const isParagraphMarkResolutionJoin = (
  value: unknown,
): value is NonNullable<ParagraphMarkChange["resolutionJoin"]> =>
  natural(value) && value <= MAX_REVISION_ID;

const isEncodedParagraphMarkResolutionJoin = (value: unknown): value is number | null =>
  value === null || isParagraphMarkResolutionJoin(value);

export const isBoundaryJoins = (value: unknown): value is BoundaryJoins =>
  Array.isArray(value) &&
  value.length <= 2 &&
  new Set(value).size === value.length &&
  value.every((side: unknown) => side === "before" || side === "after");

export class ReviewResolutionProvenanceError extends TaggedError(
  "ReviewResolutionProvenanceError",
)<{
  message: string;
  attribute: string;
  reason: "invalid" | "unsupportedVersion" | "unrepresentable";
}> {}

type ParseAttributeOptions<Value> = {
  node: XmlElement;
  attribute: string;
  valid: (value: unknown) => value is Value;
};
const parseAttribute = <Value>({
  node,
  attribute,
  valid,
}: ParseAttributeOptions<Value>): Value | undefined => {
  const encoded = getAttributeByNamespaceUri(node, NAMESPACES, attribute);
  if (encoded === null) return undefined;
  const failed = (reason: "invalid" | "unsupportedVersion" = "invalid") =>
    new ReviewResolutionProvenanceError({
      attribute,
      message: `Invalid or unsupported ${attribute} review provenance.`,
      reason,
    });
  if (encoded.length > MAX_ATTRIBUTE_LENGTH) throw failed();
  const parsed = Result.try({
    try: (): unknown => JSON.parse(encoded),
    catch: () => failed(),
  });
  if (parsed.isErr()) throw parsed.error;
  const decoded = parsed.value;
  if (!record(decoded) || Object.keys(decoded).some((key) => key !== "version" && key !== "value"))
    throw failed();
  const value = decoded["value"];
  if (!valid(value)) throw failed();
  if (decoded["version"] !== VERSION) throw failed("unsupportedVersion");
  return value;
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
  const encoded = JSON.stringify({ version: VERSION, value });
  if (encoded.length > MAX_ATTRIBUTE_LENGTH) {
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

export const parseResolutionJoins = (node: XmlElement): ResolutionJoins | undefined =>
  parseAttribute({ node, attribute: "resolutionJoins", valid: isResolutionJoins });
export const parseBoundaryJoins = (node: XmlElement): BoundaryJoins | undefined =>
  parseAttribute({ node, attribute: "boundaryJoins", valid: isBoundaryJoins });
export const serializeResolutionJoins = (value: Insertion["resolutionJoins"]): string =>
  serializeAttribute({ attribute: "resolutionJoins", value, valid: isResolutionJoins });
export const serializeBoundaryJoins = (value: RunPropertyChange["boundaryJoins"]): string =>
  serializeAttribute({ attribute: "boundaryJoins", value, valid: isBoundaryJoins });

/** Null carries an owned undefined depth; a missing attribute carries no field. */
export const parseParagraphMarkResolutionJoin = (
  node: XmlElement,
): Pick<ParagraphMarkChange, "resolutionJoin"> => {
  const value = parseAttribute({
    node,
    attribute: PARAGRAPH_MARK_FIELDS.resolutionJoin.attribute,
    valid: isEncodedParagraphMarkResolutionJoin,
  });
  return value === undefined ? {} : { resolutionJoin: value === null ? undefined : value };
};

const paragraphMarkResolutionJoinOptions = (
  mark: ParagraphMarkChange,
): SerializeAttributeOptions => ({
  attribute: PARAGRAPH_MARK_FIELDS.resolutionJoin.attribute,
  value: Object.hasOwn(mark, "resolutionJoin") ? (mark.resolutionJoin ?? null) : undefined,
  valid: isEncodedParagraphMarkResolutionJoin,
});

export const serializeParagraphMarkResolutionJoin = (mark: ParagraphMarkChange): string =>
  serializeAttribute(paragraphMarkResolutionJoinOptions(mark));

/** Captured paragraph properties compose the same envelope into a parsed XML element. */
export const paragraphMarkResolutionJoinAttributes = (
  mark: ParagraphMarkChange,
): Record<string, string> => {
  const options = paragraphMarkResolutionJoinOptions(mark);
  const encoded = encodeAttribute(options);
  return encoded === undefined ? {} : { [`folio:${options.attribute}`]: encoded };
};
