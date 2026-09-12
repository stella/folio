import { panic } from "better-result";

import type {
  FolioContentPropertyChange,
  FolioContentPropertyPresence,
  FolioContentPropertySet,
  FolioContentPropertyValue,
} from "./content-types";

/** Exact structural equality over canonical, sorted neutral property values. @internal */
export const sameFolioContentPropertyValue = (
  base: FolioContentPropertyValue,
  revised: FolioContentPropertyValue,
): boolean => {
  if (base === revised) return true;
  if (
    typeof base !== "object" ||
    base === null ||
    typeof revised !== "object" ||
    revised === null
  ) {
    return false;
  }
  if (base.type !== revised.type) return false;
  if (base.type === "array" && revised.type === "array") {
    return (
      base.items.length === revised.items.length &&
      base.items.every((value, index) => {
        const counterpart = revised.items[index];
        return counterpart !== undefined && sameFolioContentPropertyValue(value, counterpart);
      })
    );
  }
  if (base.type === "object" && revised.type === "object") {
    return sameFolioContentPropertySet(base.entries, revised.entries);
  }
  return panic("Canonical property values with one discriminator failed to narrow");
};

/** Exact structural equality over canonical property sets. @internal */
export const sameFolioContentPropertySet = (
  base: FolioContentPropertySet,
  revised: FolioContentPropertySet,
): boolean =>
  base.length === revised.length &&
  base.every((entry, index) => {
    const counterpart = revised[index];
    return (
      counterpart !== undefined &&
      entry.key === counterpart.key &&
      sameFolioContentPropertyValue(entry.value, counterpart.value)
    );
  });

const ABSENT_PROPERTY = Object.freeze({ type: "absent" } as const);

const presentProperty = (value: FolioContentPropertyValue): FolioContentPropertyPresence =>
  Object.freeze({ type: "present", value });

/** One exhaustive, ordered delta over two canonical property sets. @internal */
export const changedFolioContentProperties = <Key extends string>(
  base: FolioContentPropertySet<Key>,
  revised: FolioContentPropertySet<Key>,
): readonly FolioContentPropertyChange<Key>[] => {
  const changes: FolioContentPropertyChange<Key>[] = [];
  let baseIndex = 0;
  let revisedIndex = 0;
  while (baseIndex < base.length || revisedIndex < revised.length) {
    const baseEntry = base[baseIndex];
    const revisedEntry = revised[revisedIndex];
    if (baseEntry && (!revisedEntry || baseEntry.key < revisedEntry.key)) {
      changes.push(
        Object.freeze({
          key: baseEntry.key,
          base: presentProperty(baseEntry.value),
          revised: ABSENT_PROPERTY,
        }),
      );
      baseIndex++;
      continue;
    }
    if (revisedEntry && (!baseEntry || revisedEntry.key < baseEntry.key)) {
      changes.push(
        Object.freeze({
          key: revisedEntry.key,
          base: ABSENT_PROPERTY,
          revised: presentProperty(revisedEntry.value),
        }),
      );
      revisedIndex++;
      continue;
    }
    if (!baseEntry || !revisedEntry) {
      return panic("Canonical property merge lost an entry");
    }
    if (!sameFolioContentPropertyValue(baseEntry.value, revisedEntry.value)) {
      changes.push(
        Object.freeze({
          key: baseEntry.key,
          base: presentProperty(baseEntry.value),
          revised: presentProperty(revisedEntry.value),
        }),
      );
    }
    baseIndex++;
    revisedIndex++;
  }
  return Object.freeze(changes);
};

/** Structural equality for already canonical property deltas. @internal */
export const sameFolioContentPropertyChanges = (
  left: readonly FolioContentPropertyChange[],
  right: readonly FolioContentPropertyChange[],
): boolean =>
  left.length === right.length &&
  left.every((change, index) => {
    const counterpart = right[index];
    if (
      counterpart === undefined ||
      change.key !== counterpart.key ||
      change.base.type !== counterpart.base.type ||
      change.revised.type !== counterpart.revised.type
    ) {
      return false;
    }
    return (
      (change.base.type === "absent" ||
        (counterpart.base.type === "present" &&
          sameFolioContentPropertyValue(change.base.value, counterpart.base.value))) &&
      (change.revised.type === "absent" ||
        (counterpart.revised.type === "present" &&
          sameFolioContentPropertyValue(change.revised.value, counterpart.revised.value)))
    );
  });
