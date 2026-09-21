/** Deliberate violations: widening a model type with a local intersection,
 * carrying a data field alongside a phantom brand, and reading a widened
 * field back with an `in` check instead of a typed property access
 * (issue #845). */

import type { ListRendering, Paragraph } from "../../packages/core/src/types/document";

const widened = (rendering: ListRendering): ListRendering & { levelStarts: number[] } => ({
  ...rendering,
  levelStarts: [1],
});

declare const RENDERING_ATTR: unique symbol;

export type BrandedAndWidened = ListRendering & {
  readonly [RENDERING_ATTR]: true;
  levelStarts: number[];
};

const DATA_KEY = "levelStarts" as const;

export type ComputedDataWidening = ListRendering & { [DATA_KEY]: number[] };

export const readStarts = (paragraph: Paragraph): number[] | undefined =>
  paragraph.listRendering && "levelStarts" in paragraph.listRendering ? [1] : undefined;

export { widened };
