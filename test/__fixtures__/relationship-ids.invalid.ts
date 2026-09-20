/** Deliberate violations: absence written as a relationship reference. */

type Drawing = { rId?: string; hlinkRId?: string };

export const literalReference: Drawing = { rId: "" };

export const fallbackReference = (attrs: { hlinkRId?: string }): Drawing => ({
  hlinkRId: attrs.hlinkRId || "",
});

export const assignedReference = (drawing: Drawing): void => {
  drawing.rId = "";
};
