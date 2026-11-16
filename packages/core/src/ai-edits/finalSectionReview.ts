import type { FormattingPatch } from "@stll/docx-core/ops";

import { sectionReferenceHistory } from "../docx/sectionReferenceHistory";
import { sectionRejectProperties } from "../prosemirror/commands/propertyChangeScope";
import type { SectionProperties } from "../types/document";
import { canonicalJson } from "../utils/canonicalJson";
import type { FolioRevisionStamp } from "./apply";

type ResolveFinalSectionChangeOptions = {
  current: SectionProperties;
  mode: "accept" | "reject";
  id: number;
};

/** Rebase the remaining changes through the legacy section-resolution calculation. */
export const resolveFinalSectionChange = ({
  current,
  mode,
  id,
}: ResolveFinalSectionChangeOptions): SectionProperties | null => {
  const changes = current.propertyChanges;
  const targetIndex = changes?.findIndex(({ info }) => info.id === id);
  if (
    !changes ||
    changes.length === 0 ||
    targetIndex === undefined ||
    targetIndex < 0
  )
    return null;
  const { propertyChanges: _propertyChanges, ...liveProperties } = current;
  const previousStates = changes.map((change) =>
    sectionRejectProperties({
      live: current,
      previousProperties: change.previousProperties,
      previousReferences: change.previousReferences,
    }),
  );
  let resolvedProperties: SectionProperties =
    previousStates[0] ?? liveProperties;
  const remaining = [];
  for (const [index, change] of changes.entries()) {
    const nextProperties = previousStates[index + 1] ?? liveProperties;
    if (index === targetIndex) {
      if (mode === "accept") resolvedProperties = nextProperties;
      continue;
    }
    remaining.push({
      ...change,
      previousProperties: resolvedProperties,
      ...(change.previousReferences !== undefined && {
        previousReferences: {
          ...(resolvedProperties.headerReferences !== undefined && {
            headerReferences: resolvedProperties.headerReferences,
          }),
          ...(resolvedProperties.footerReferences !== undefined && {
            footerReferences: resolvedProperties.footerReferences,
          }),
        },
      }),
    });
    resolvedProperties = nextProperties;
  }
  return {
    ...resolvedProperties,
    ...(remaining.length > 0 && { propertyChanges: remaining }),
  };
};

/** Resolve the whole final-section history through the same baseline as the legacy reviewer. */
export const resolveFinalSectionChanges = (
  current: SectionProperties,
  mode: "accept" | "reject",
): SectionProperties | null => {
  const firstChange = current.propertyChanges?.at(0);
  if (!firstChange) return null;
  const { propertyChanges: _propertyChanges, ...liveProperties } = current;
  return mode === "accept"
    ? liveProperties
    : sectionRejectProperties({
        live: current,
        previousProperties: firstChange.previousProperties,
        previousReferences: firstChange.previousReferences,
      });
};

type StageFinalSectionPropertiesOptions = {
  target: SectionProperties;
  previous: SectionProperties;
  revision: FolioRevisionStamp;
  author: string;
};

/** The tracked final-section state both reviewer engines stage. */
export const stagedFinalSectionProperties = ({
  target,
  previous,
  revision,
  author,
}: StageFinalSectionPropertiesOptions): SectionProperties | null => {
  if (canonicalJson(target) === canonicalJson(previous)) return null;
  const previousReferences = sectionReferenceHistory({ previous, target });
  return {
    ...target,
    propertyChanges: [
      {
        type: "sectionPropertyChange",
        info: { id: revision.idSeed, author, date: revision.date },
        previousProperties: previous,
        ...(previousReferences !== undefined && { previousReferences }),
        currentProperties: target,
      },
    ],
  };
};

type FinalSectionResolutionPatchOptions = {
  current: SectionProperties;
  next: SectionProperties;
};
/** Explicit nulls remove fields the resolved section no longer carries. */
export const finalSectionResolutionPatch = ({
  current,
  next,
}: FinalSectionResolutionPatchOptions): FormattingPatch<SectionProperties> =>
  Object.fromEntries(
    [...new Set([...Object.keys(current), ...Object.keys(next)])].map((key) => [
      key,
      Object.hasOwn(next, key) ? Reflect.get(next, key) : null,
    ]),
  );
