import { cloneModel } from "./modelClone";
import { Result } from "better-result";

import { MAX_REVISION_ID, type Document } from "../model/document";
import type { DocumentEdit } from "./edits";
import { structurallyEqual } from "./equality";
import { DOCUMENT_OP_REFUSAL_REASONS, DocumentOpRefusal } from "./refusal";
import {
  DOCUMENT_OP_TYPES,
  type PackageResourcePart,
  type PackageResources,
  type SetPackageResourcesOp,
} from "./types";

const partOf = <Value>(present: boolean, value: Value | undefined): PackageResourcePart<Value> => {
  if (!present) return { type: "omitted" };
  return value === undefined
    ? { type: "undefined" }
    : { type: "present", value: cloneModel(value) };
};

/** Capture maps and binary media as ordinary JSON data, preserving field presence. */
export const packageResourcesOf = ({ package: pkg }: Document): PackageResources => ({
  styles: partOf(Object.hasOwn(pkg, "styles"), pkg.styles),
  numbering: partOf(Object.hasOwn(pkg, "numbering"), pkg.numbering),
  relationships: partOf(
    Object.hasOwn(pkg, "relationships"),
    pkg.relationships === undefined ? undefined : [...pkg.relationships.entries()],
  ),
  media: partOf(
    Object.hasOwn(pkg, "media"),
    pkg.media === undefined
      ? undefined
      : [...pkg.media.entries()].map(
          ([key, media]) =>
            [key, Object.assign({}, media, { data: [...new Uint8Array(media.data)] })] as const,
        ),
  ),
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

const duplicates = (ids: readonly (string | number)[]): boolean => new Set(ids).size !== ids.length;

const resourceViolation = (
  resources: PackageResources,
  expected: PackageResources,
): string | undefined => {
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
  const relationships = valueOf(resources.relationships) ?? [];
  const media = valueOf(resources.media) ?? [];
  if (
    duplicates(relationships.map(([key]) => key)) ||
    relationships.some(([key, relationship]) => key.length === 0 || key !== relationship.id)
  )
    return "Imported relationships contain invalid or duplicate ids.";
  if (
    duplicates(media.map(([key]) => key)) ||
    media.some(
      ([key, file]) =>
        key !== file.path ||
        !key.startsWith("word/media/") ||
        key.split("/").includes("..") ||
        file.data.some((byte) => !Number.isInteger(byte) || byte < 0 || byte > 255),
    )
  )
    return "Imported media contains invalid paths or bytes.";
  const paths = new Set(media.map(([key]) => key));
  for (const [, relationship] of relationships) {
    if (!relationship.type.endsWith("/image") || relationship.targetMode === "External") continue;
    const target = relationship.target.startsWith("/")
      ? relationship.target.slice(1)
      : `word/${relationship.target}`;
    if (!paths.has(target)) return "An imported image relationship has no media part.";
  }
  const available = {
    style: new Set<string | number>(styles?.styles.map(({ styleId }) => styleId) ?? []),
    numbering: new Set<string | number>(numbering?.nums.map(({ numId }) => numId) ?? []),
    relationship: new Set<string | number>(relationships.map(([key]) => key)),
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
  const violation = resourceViolation(op.resources, expected);
  if (violation !== undefined)
    return failed(DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH, violation);
  // The journal and resulting document must not share mutable resource records.
  const resources = cloneModel(op.resources);
  const pkg = { ...document.package };
  delete pkg.styles;
  delete pkg.numbering;
  delete pkg.relationships;
  delete pkg.media;
  if (resources.styles.type !== "omitted") pkg.styles = valueOf(resources.styles);
  if (resources.numbering.type !== "omitted") pkg.numbering = valueOf(resources.numbering);
  if (resources.relationships.type !== "omitted") {
    const relationships = valueOf(resources.relationships);
    pkg.relationships = relationships === undefined ? undefined : new Map(relationships);
  }
  if (resources.media.type !== "omitted") {
    const media = valueOf(resources.media);
    pkg.media =
      media === undefined
        ? undefined
        : new Map(
            media.map(([key, file]) => [
              key,
              Object.assign({}, file, { data: new Uint8Array(file.data).buffer }),
            ]),
          );
  }
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
  if (structurallyEqual(expected, op.resources))
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
      },
    ],
    touched: { modified: [], inserted: [], removed: [] },
  });
};
