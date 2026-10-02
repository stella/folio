import { Result, TaggedError } from "better-result";
import { escapeXmlAttribute } from "@stll/docx-core";
import { MAX_REVISION_ID } from "@stll/docx-core/model";
import type { Insertion, RunPropertyChange } from "../types/document";
import { getAttributeByNamespaceUri, type XmlElement } from "./xmlParser";
import { FOLIO_REVIEW_HISTORY_NAMESPACE } from "./reviewHistoryNamespace";

const NAMESPACES: ReadonlySet<string> = new Set([FOLIO_REVIEW_HISTORY_NAMESPACE]);
const VERSION = 1;
const MAX_ATTRIBUTE_LENGTH = 32_768;
type ResolutionJoins = NonNullable<Insertion["resolutionJoins"]>;
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
} as const satisfies Record<keyof ResolutionJoins, true>;

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
      (slot.space === "revision" || slot.space === "control") &&
      natural(slot.id) &&
      slot.id <= MAX_REVISION_ID,
  );
const retained = (value: unknown): value is readonly RetainedIdentity[] =>
  Array.isArray(value) &&
  value.every(
    (item: unknown) =>
      record(item) &&
      Object.keys(item).every((key) => Object.hasOwn(RETAINED_FIELDS, key)) &&
      natural(item.depth) &&
      slots(item.source) &&
      slots(item.target) &&
      item.source.length === item.target.length &&
      item.source.every((slot, index) => slot.space === item.target.at(index)?.space),
  );

export const isResolutionJoins = (value: unknown): value is ResolutionJoins =>
  record(value) &&
  Object.keys(value).every((key) => Object.hasOwn(JOIN_FIELDS, key)) &&
  natural(value.before) &&
  natural(value.after) &&
  natural(value.remove) &&
  (value.retainedAfter === undefined || retained(value.retainedAfter));

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
  const decoded = Result.try({
    try: (): unknown => JSON.parse(encoded),
    catch: () => failed(),
  }).unwrap();
  if (
    !record(decoded) ||
    Object.keys(decoded).some((key) => key !== "version" && key !== "value") ||
    !valid(decoded.value)
  )
    throw failed();
  if (decoded.version !== VERSION) throw failed("unsupportedVersion");
  return decoded.value;
};

type SerializeAttributeOptions = {
  attribute: string;
  value: unknown;
  valid: (value: unknown) => boolean;
};
const serializeAttribute = ({ attribute, value, valid }: SerializeAttributeOptions): string => {
  if (value === undefined) return "";
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
  return ` folio:${attribute}="${escapeXmlAttribute(encoded)}"`;
};

export const parseResolutionJoins = (node: XmlElement): ResolutionJoins | undefined =>
  parseAttribute({ node, attribute: "resolutionJoins", valid: isResolutionJoins });
export const parseBoundaryJoins = (node: XmlElement): BoundaryJoins | undefined =>
  parseAttribute({ node, attribute: "boundaryJoins", valid: isBoundaryJoins });
export const serializeResolutionJoins = (value: Insertion["resolutionJoins"]): string =>
  serializeAttribute({ attribute: "resolutionJoins", value, valid: isResolutionJoins });
export const serializeBoundaryJoins = (value: RunPropertyChange["boundaryJoins"]): string =>
  serializeAttribute({ attribute: "boundaryJoins", value, valid: isBoundaryJoins });
