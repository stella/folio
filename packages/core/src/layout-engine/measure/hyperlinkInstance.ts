import type { HyperlinkInfo } from "../types";

// Imported hyperlink wrappers carry a conversion-wide index in ProseMirror.
// Keep it out of the public Flow shape while letting the painter distinguish
// adjacent, separate links that happen to share the same visible metadata.
const hyperlinkInstanceIndexes = new WeakMap<HyperlinkInfo, number>();

export const setHyperlinkInstanceIndex = (
  hyperlink: HyperlinkInfo,
  instanceIndex: number,
): void => {
  hyperlinkInstanceIndexes.set(hyperlink, instanceIndex);
};

export const getHyperlinkInstanceIndex = (hyperlink: HyperlinkInfo): number | undefined =>
  hyperlinkInstanceIndexes.get(hyperlink);
