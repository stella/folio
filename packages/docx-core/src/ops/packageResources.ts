import { cloneModel } from "./modelClone";
import { Result } from "better-result";

import { MAX_REVISION_ID, type Document } from "../model/document";
import type { MediaFile, Relationship } from "../model/styles";
import type { DocumentEdit } from "./edits";
import { structurallyEqual } from "./equality";
import { DOCUMENT_OP_REFUSAL_REASONS, DocumentOpRefusal } from "./refusal";
import {
  DOCUMENT_OP_TYPES,
  type PackageResourceEntry,
  type PackageResourceEntryChange,
  type PackageResourceMapChange,
  type PackageResourceMedia,
  type PackageResourcePart,
  type PackageResourcePresence,
  type PackageResources,
  type SetPackageResourcesOp,
} from "./types";

const partOf = <Value>(present: boolean, value: Value | undefined): PackageResourcePart<Value> => {
  if (!present) return { type: "omitted" };
  return value === undefined
    ? { type: "undefined" }
    : { type: "present", value: cloneModel(value) };
};

/** Capture style and numbering definitions as ordinary JSON data, preserving field presence. */
export const packageResourcesOf = ({ package: pkg }: Document): PackageResources => ({
  styles: partOf(Object.hasOwn(pkg, "styles"), pkg.styles),
  numbering: partOf(Object.hasOwn(pkg, "numbering"), pkg.numbering),
});

const valueOf = <Value>(part: PackageResourcePart<Value>): Value | undefined => {
  switch (part.type) {
    case "omitted":
    case "undefined":
      return undefined;
    case "present":
      return part.value;
    default: {
      const unreachable: never = part;
      return unreachable;
    }
  }
};

/** Binary media travels as a byte array, rather than an ArrayBuffer that JSON discards. */
const mediaRecordOf = (media: MediaFile): PackageResourceMedia =>
  Object.assign({}, media, { data: [...new Uint8Array(media.data)] });

const mediaFileOf = (record: PackageResourceMedia): MediaFile =>
  Object.assign({}, record, { data: new Uint8Array(record.data).buffer });

type PackageMap<Stored> = {
  presence: PackageResourcePresence;
  entries: ReadonlyMap<string, Stored>;
};

const packageMapOf = <Stored>(
  present: boolean,
  map: ReadonlyMap<string, Stored> | undefined,
): PackageMap<Stored> => ({
  presence: !present ? "omitted" : map === undefined ? "undefined" : "present",
  entries: map ?? new Map(),
});

const relationshipsOf = ({ package: pkg }: Document) =>
  packageMapOf(Object.hasOwn(pkg, "relationships"), pkg.relationships);

const mediaOf = ({ package: pkg }: Document) =>
  packageMapOf(Object.hasOwn(pkg, "media"), pkg.media);

const entryOf = <Stored, Value>(
  stored: Stored | undefined,
  record: (stored: Stored) => Value,
): PackageResourceEntry<Value> =>
  stored === undefined ? { type: "absent" } : { type: "present", value: record(stored) };

type MapChangeOptions<Stored, Value> = {
  before: PackageMap<Stored>;
  after: PackageMap<Stored>;
  record: (stored: Stored) => Value;
};

const mapChangeOf = <Stored, Value>({
  before,
  after,
  record,
}: MapChangeOptions<Stored, Value>): PackageResourceMapChange<Value> => {
  const entries: PackageResourceEntryChange<Value>[] = [];
  for (const key of new Set([...before.entries.keys(), ...after.entries.keys()])) {
    const previous = before.entries.get(key);
    const next = after.entries.get(key);
    // Untouched entries keep their record, so an import never copies media it leaves alone.
    if (previous === next) continue;
    const change = { key, expected: entryOf(previous, record), next: entryOf(next, record) };
    if (!structurallyEqual(change.expected, change.next)) entries.push(change);
  }
  return { expected: before.presence, next: after.presence, entries };
};

type PackageResourcesOpOptions = { before: Document; after: Document };

/** Describe the package resource difference as one operation carrying only changed entries. */
export const packageResourcesOpOf = ({
  before,
  after,
}: PackageResourcesOpOptions): SetPackageResourcesOp => ({
  type: DOCUMENT_OP_TYPES.SET_PACKAGE_RESOURCES,
  expected: packageResourcesOf(before),
  resources: packageResourcesOf(after),
  relationships: mapChangeOf({
    before: relationshipsOf(before),
    after: relationshipsOf(after),
    record: (relationship: Relationship) => cloneModel(relationship),
  }),
  media: mapChangeOf({ before: mediaOf(before), after: mediaOf(after), record: mediaRecordOf }),
});

type MapApplication<Stored> =
  | { type: "applied"; map: PackageMap<Stored> }
  | { type: "refused"; reason: "stale" | "structureMismatch"; message: string };

type ApplyMapChangeOptions<Stored, Value> = {
  current: PackageMap<Stored>;
  change: PackageResourceMapChange<Value>;
  record: (stored: Stored) => Value;
  stored: (value: Value) => Stored;
};

const applyMapChange = <Stored, Value>({
  current,
  change,
  record,
  stored,
}: ApplyMapChangeOptions<Stored, Value>): MapApplication<Stored> => {
  if (current.presence !== change.expected)
    return {
      type: "refused",
      reason: DOCUMENT_OP_REFUSAL_REASONS.STALE,
      message: "The package resources changed.",
    };
  if (duplicates(change.entries.map(({ key }) => key)))
    return {
      type: "refused",
      reason: DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
      message: "Package resource changes repeat a key.",
    };
  const entries = new Map(current.entries);
  for (const { key, expected, next } of change.entries) {
    if (!structurallyEqual(entryOf(current.entries.get(key), record), expected))
      return {
        type: "refused",
        reason: DOCUMENT_OP_REFUSAL_REASONS.STALE,
        message: "The package resources changed.",
      };
    if (next.type === "present") entries.set(key, stored(next.value));
    else entries.delete(key);
  }
  if (change.next !== "present" && entries.size > 0)
    return {
      type: "refused",
      reason: DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
      message: "A removed package map still has entries.",
    };
  return { type: "applied", map: { presence: change.next, entries } };
};

const inverseMapChange = <Value>(
  change: PackageResourceMapChange<Value>,
): PackageResourceMapChange<Value> => ({
  expected: change.next,
  next: change.expected,
  entries: change.entries.map(({ key, expected, next }) => ({
    key,
    expected: next,
    next: expected,
  })),
});

const presentEntries = <Value>(change: PackageResourceMapChange<Value>) =>
  change.entries.flatMap(({ key, next }) =>
    next.type === "present" ? [[key, next.value] as const] : [],
  );

const duplicates = (ids: readonly (string | number)[]): boolean => new Set(ids).size !== ids.length;

type ResourceViolationOptions = {
  op: SetPackageResourcesOp;
  expected: PackageResources;
  relationships: ReadonlyMap<string, Relationship>;
  media: ReadonlyMap<string, MediaFile>;
};

/** Validate the definitions and entries this operation adds or changes. */
const resourceViolation = ({
  op: { resources, relationships: relationshipChange, media: mediaChange },
  expected,
  relationships,
  media,
}: ResourceViolationOptions): string | undefined => {
  const styles = valueOf(resources.styles);
  if (styles !== undefined && duplicates(styles.styles.map(({ styleId }) => styleId)))
    return "Imported styles contain duplicate ids.";
  const numbering = valueOf(resources.numbering);
  if (numbering !== undefined) {
    const abstracts = numbering.abstractNums.map(({ abstractNumId }) => abstractNumId);
    const nums = numbering.nums.map(({ numId }) => numId);
    if (
      duplicates(abstracts) ||
      duplicates(nums) ||
      abstracts.some((id) => !Number.isSafeInteger(id) || id < 0 || id > MAX_REVISION_ID) ||
      nums.some((id) => !Number.isSafeInteger(id) || id <= 0 || id > MAX_REVISION_ID)
    )
      return "Imported numbering contains invalid or duplicate ids.";
    if (numbering.nums.some(({ abstractNumId }) => !abstracts.includes(abstractNumId)))
      return "An imported numbering instance has no abstract definition.";
  }
  const changedRelationships = presentEntries(relationshipChange);
  if (
    changedRelationships.some(([key, relationship]) => key.length === 0 || key !== relationship.id)
  )
    return "Imported relationships contain invalid ids.";
  if (
    presentEntries(mediaChange).some(
      ([key, file]) =>
        key !== file.path ||
        !key.startsWith("word/media/") ||
        key.split("/").includes("..") ||
        file.data.some((byte) => !Number.isInteger(byte) || byte < 0 || byte > 255),
    )
  )
    return "Imported media contains invalid paths or bytes.";
  for (const [, relationship] of changedRelationships) {
    if (!relationship.type.endsWith("/image") || relationship.targetMode === "External") continue;
    const target = relationship.target.startsWith("/")
      ? relationship.target.slice(1)
      : `word/${relationship.target}`;
    if (!media.has(target)) return "An imported image relationship has no media part.";
  }
  const available = {
    style: new Set<string | number>(styles?.styles.map(({ styleId }) => styleId) ?? []),
    numbering: new Set<string | number>(numbering?.nums.map(({ numId }) => numId) ?? []),
    relationship: new Set<string | number>(relationships.keys()),
  } satisfies Record<ResourceReferenceKind, ReadonlySet<string | number>>;
  const unavailable = (kind: ResourceReferenceKind, id: string | number) =>
    !available[kind].has(id);
  const previousStyles = new Map(
    valueOf(expected.styles)?.styles.map((style) => [style.styleId, style]),
  );
  for (const style of styles?.styles ?? []) {
    if (structurallyEqual(style, previousStyles.get(style.styleId))) continue;
    if (resourceReferencesMatch({ value: style, matches: unavailable, relationshipScope: "main" }))
      return "An imported style references an unavailable package resource.";
  }
  const previousAbstracts = new Map(
    valueOf(expected.numbering)?.abstractNums.map((abstract) => [abstract.abstractNumId, abstract]),
  );
  for (const abstract of numbering?.abstractNums ?? []) {
    if (structurallyEqual(abstract, previousAbstracts.get(abstract.abstractNumId))) continue;
    if (
      resourceReferencesMatch({ value: abstract, matches: unavailable, relationshipScope: "main" })
    )
      return "An imported numbering definition references an unavailable style.";
  }
  const previousNums = new Map(valueOf(expected.numbering)?.nums.map((num) => [num.numId, num]));
  for (const num of numbering?.nums ?? []) {
    if (structurallyEqual(num, previousNums.get(num.numId))) continue;
    if (resourceReferencesMatch({ value: num, matches: unavailable, relationshipScope: "main" }))
      return "An imported numbering instance references an unavailable style.";
  }
  return undefined;
};

const RESOURCE_REFERENCE_KINDS = {
  styleId: "style",
  basedOn: "style",
  next: "style",
  link: "style",
  pStyle: "style",
  numStyleLink: "style",
  styleLink: "style",
  numId: "numbering",
  rId: "relationship",
  hlinkRId: "relationship",
  imageRId: "relationship",
} as const;

type ResourceReferenceKind =
  (typeof RESOURCE_REFERENCE_KINDS)[keyof typeof RESOURCE_REFERENCE_KINDS];

const referenceKinds = new Map<string, ResourceReferenceKind>(
  Object.entries(RESOURCE_REFERENCE_KINDS),
);

type ResourceReferencesOptions = {
  value: unknown;
  matches: (kind: ResourceReferenceKind, id: string | number) => boolean;
  relationshipScope: "main" | "part";
};

const resourceReferencesMatch = ({
  value,
  matches,
  relationshipScope,
}: ResourceReferencesOptions): boolean => {
  if (typeof value !== "object" || value === null) return false;
  if (value instanceof Map)
    return [...value.values()].some((item) =>
      resourceReferencesMatch({ value: item, matches, relationshipScope }),
    );
  if (Array.isArray(value))
    return value.some((item) =>
      resourceReferencesMatch({ value: item, matches, relationshipScope }),
    );
  return Object.entries(value).some(([key, item]) => {
    // Rendering caches state computed display facts, not package dependencies.
    if (key === "listRendering") return false;
    const kind = referenceKinds.get(key);
    if (
      kind !== undefined &&
      (kind !== "relationship" || relationshipScope === "main") &&
      (typeof item === "string" || typeof item === "number") &&
      matches(kind, item)
    )
      return true;
    // Header/footer rIds belong to their own relationship parts; styles and numbering remain global.
    const scope = key === "headers" || key === "footers" ? "part" : relationshipScope;
    return resourceReferencesMatch({ value: item, matches, relationshipScope: scope });
  });
};

const missingIds = (before: readonly (string | number)[], after: readonly (string | number)[]) => {
  const kept = new Set(after);
  return new Set(before.filter((id) => !kept.has(id)));
};

export const applyPackageResourcesOp = (
  document: Document,
  op: SetPackageResourcesOp,
): Result<DocumentEdit, DocumentOpRefusal> => {
  const expected = packageResourcesOf(document);
  const failed = (reason: "stale" | "structureMismatch", message: string) =>
    Result.err(new DocumentOpRefusal({ reason, message, opType: op.type }));
  if (!structurallyEqual(expected, op.expected))
    return failed(DOCUMENT_OP_REFUSAL_REASONS.STALE, "The package resources changed.");
  // The journal and resulting document must not share mutable resource records.
  const owned = cloneModel(op);
  const relationships = applyMapChange({
    current: relationshipsOf(document),
    change: owned.relationships,
    record: (relationship: Relationship) => cloneModel(relationship),
    stored: (relationship: Relationship) => cloneModel(relationship),
  });
  if (relationships.type === "refused") return failed(relationships.reason, relationships.message);
  const media = applyMapChange({
    current: mediaOf(document),
    change: owned.media,
    record: mediaRecordOf,
    stored: mediaFileOf,
  });
  if (media.type === "refused") return failed(media.reason, media.message);
  const violation = resourceViolation({
    op: owned,
    expected,
    relationships: relationships.map.entries,
    media: media.map.entries,
  });
  if (violation !== undefined)
    return failed(DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH, violation);
  const { resources } = owned;
  const pkg = { ...document.package };
  delete pkg.styles;
  delete pkg.numbering;
  delete pkg.relationships;
  delete pkg.media;
  if (resources.styles.type !== "omitted") pkg.styles = valueOf(resources.styles);
  if (resources.numbering.type !== "omitted") pkg.numbering = valueOf(resources.numbering);
  if (relationships.map.presence !== "omitted")
    pkg.relationships =
      relationships.map.presence === "present" ? new Map(relationships.map.entries) : undefined;
  if (media.map.presence !== "omitted")
    pkg.media = media.map.presence === "present" ? new Map(media.map.entries) : undefined;
  const removed = {
    style: missingIds(
      document.package.styles?.styles.map(({ styleId }) => styleId) ?? [],
      pkg.styles?.styles.map(({ styleId }) => styleId) ?? [],
    ),
    numbering: missingIds(
      document.package.numbering?.nums.map(({ numId }) => numId) ?? [],
      pkg.numbering?.nums.map(({ numId }) => numId) ?? [],
    ),
    relationship: missingIds(
      [...(document.package.relationships?.keys() ?? [])],
      [...(pkg.relationships?.keys() ?? [])],
    ),
  } satisfies Record<ResourceReferenceKind, ReadonlySet<string | number>>;
  if (
    resourceReferencesMatch({
      value: pkg,
      matches: (kind, id) => removed[kind].has(id),
      relationshipScope: "main",
    })
  )
    return failed(
      DOCUMENT_OP_REFUSAL_REASONS.STALE,
      "A removed package resource is still referenced.",
    );
  if (
    structurallyEqual(expected, resources) &&
    owned.relationships.expected === owned.relationships.next &&
    owned.relationships.entries.length === 0 &&
    owned.media.expected === owned.media.next &&
    owned.media.entries.length === 0
  )
    return Result.ok({
      document,
      inverse: [],
      touched: { modified: [], inserted: [], removed: [] },
    });
  const updated = { ...document, package: pkg };
  return Result.ok({
    document: updated,
    inverse: [
      {
        type: DOCUMENT_OP_TYPES.SET_PACKAGE_RESOURCES,
        expected: packageResourcesOf(updated),
        resources: expected,
        relationships: inverseMapChange(owned.relationships),
        media: inverseMapChange(owned.media),
      },
    ],
    touched: { modified: [], inserted: [], removed: [] },
  });
};
