/** Deliberate violations: widening a model type with a local intersection,
 * and reading a widened field back with an `in` check instead of a typed
 * property access (issue #845). */

import type { ListRendering, Paragraph } from "../../packages/core/src/types/document";

const widened = (rendering: ListRendering): ListRendering & { levelStarts: number[] } => ({
  ...rendering,
  levelStarts: [1],
});

export const readStarts = (paragraph: Paragraph): number[] | undefined =>
  paragraph.listRendering && "levelStarts" in paragraph.listRendering ? [1] : undefined;

export { widened };
