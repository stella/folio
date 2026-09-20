/** Accepted patterns: a type predicate narrows an existing optional field, a
 * phantom brand adds a `unique symbol` key and no data, an `in` check guards
 * an `unknown` value (not a model-typed binding), and a widened field is read
 * directly instead of probed with `in`. */

import type { ListRendering, Paragraph } from "../../packages/core/src/types/document";

export const hasParaId = (block: Paragraph): block is Paragraph & { paraId: string } =>
  typeof block.paraId === "string";

declare const RENDERING_ATTR: unique symbol;

export type ListRenderingAttr = ListRendering & { readonly [RENDERING_ATTR]: true };

export const isNamed = (value: unknown): boolean =>
  typeof value === "object" && value !== null && "name" in value;

export const starts = (paragraph: Paragraph) => paragraph.listRendering?.levelStarts;
