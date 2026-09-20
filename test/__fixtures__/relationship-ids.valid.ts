/** Absence spelled once, and references that stay references. */

import { relationshipIdOf } from "@stll/docx-core/model";

type Drawing = { rId?: string; hlinkRId?: string };

export const noReference: Drawing = { rId: undefined };

export const narrowedReference = (attrs: { rId?: string }): Drawing => {
  const rId = relationshipIdOf(attrs.rId);
  return rId === undefined ? {} : { rId };
};

export const preferredReference = (primary?: string, secondary?: string): Drawing => ({
  rId: primary ?? secondary,
});

/** A relationship id is not the only thing in this repository named `rId`. */
export const headerSlot = (rId: string): string => `[data-rid="${rId}"]`;
