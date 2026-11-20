/**
 * Resolving a section's tracked property changes (`w:sectPrChange`): the one
 * calculation every reviewer and the revision resolution apply.
 */

import type {
  SectionProperties,
  SectionPropertyChange,
} from "../model/document";
import {
  type FormattingPatch,
  REVISION_DECISIONS,
  type RevisionDecision,
} from "./types";

type SectionRejectPropertiesOptions = {
  live: SectionProperties;
  previousProperties: SectionProperties | undefined;
  previousReferences?: SectionPropertyChange["previousReferences"];
};

/**
 * Rejected section properties: the stored old sectPr wholesale, preserving
 * the live header/footer references — `EG_HdrFtrReferences` is not part of
 * the `CT_SectPrBase` payload a `w:sectPrChange` stores (ECMA-376
 * §17.13.5.32), so those children survive a native reject. An explicit
 * reference history replaces that selection, including restoring its
 * absence. The caller re-attaches whatever `propertyChanges` remain
 * unresolved.
 */
export const sectionRejectProperties = ({
  live,
  previousProperties,
  previousReferences,
}: SectionRejectPropertiesOptions): SectionProperties => {
  const restored: SectionProperties = { ...previousProperties };
  delete restored.propertyChanges;
  const references = previousReferences ?? live;
  if (references.headerReferences)
    restored.headerReferences = references.headerReferences;
  else delete restored.headerReferences;
  if (references.footerReferences)
    restored.footerReferences = references.footerReferences;
  else delete restored.footerReferences;
  return restored;
};

type ResolveSectionChangeOptions = {
  current: SectionProperties;
  decision: RevisionDecision;
  id: number;
};

/**
 * `current` with its change `id` resolved, the changes after it rebased onto
 * what that resolution leaves; `null` when no change of `current` is `id`.
 */
export const resolveSectionChange = ({
  current,
  decision,
  id,
}: ResolveSectionChangeOptions): SectionProperties | null => {
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
  const remaining: SectionPropertyChange[] = [];
  for (const [index, change] of changes.entries()) {
    const nextProperties = previousStates[index + 1] ?? liveProperties;
    if (index === targetIndex) {
      if (decision === REVISION_DECISIONS.ACCEPT)
        resolvedProperties = nextProperties;
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

/** `current` with its whole change history resolved; `null` when it has none. */
export const resolveSectionChanges = (
  current: SectionProperties,
  decision: RevisionDecision,
): SectionProperties | null => {
  const firstChange = current.propertyChanges?.at(0);
  if (!firstChange) return null;
  const { propertyChanges: _propertyChanges, ...liveProperties } = current;
  return decision === REVISION_DECISIONS.ACCEPT
    ? liveProperties
    : sectionRejectProperties({
        live: current,
        previousProperties: firstChange.previousProperties,
        previousReferences: firstChange.previousReferences,
      });
};

type SectionResolutionPatchOptions = {
  current: SectionProperties;
  next: SectionProperties;
};

/** The patch from `current` to `next`: explicit nulls remove what `next` no longer carries. */
export const sectionResolutionPatch = ({
  current,
  next,
}: SectionResolutionPatchOptions): FormattingPatch<SectionProperties> =>
  Object.fromEntries(
    [...new Set([...Object.keys(current), ...Object.keys(next)])].map((key) => [
      key,
      Object.hasOwn(next, key) ? Reflect.get(next, key) : null,
    ]),
  );
