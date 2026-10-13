import { panic } from "better-result";
import type { RunPropertyChange } from "./content";

/** A run-property container has one pending formatting revision. */
export const getSingularRunPropertyChange = (
  propertyChanges: readonly RunPropertyChange[] | undefined,
): RunPropertyChange | undefined => {
  const propertyChangeCount = propertyChanges?.length ?? 0;
  if (propertyChangeCount > 1) {
    panic("A run-property container cannot serialize more than one w:rPrChange", {
      elementName: "w:rPrChange",
      propertyChangeCount,
    });
  }
  return propertyChanges?.at(0);
};
