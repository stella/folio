// Fixture for `folio-read-model/no-read-only-attr-cast`. Each cast hands a
// writer an attr the style cascade filled with shared values.

import type { Node as PMNode } from "prosemirror-model";

import type { ParagraphAttrs } from "@stll/folio-core/prosemirror/schema/nodes";
import type {
  ParagraphFormatting,
  TabStop,
  TextFormatting,
} from "@stll/folio-core/types/document";

export const fromNode = (node: PMNode): TextFormatting | undefined =>
  node.attrs["defaultTextFormatting"] as TextFormatting | undefined;

export const fromAttrs = (attrs: ParagraphAttrs): TextFormatting | undefined =>
  attrs.defaultTextFormatting as TextFormatting | undefined;

export const resolved = (attrs: ParagraphAttrs): ParagraphFormatting =>
  attrs._resolvedFormatting as ParagraphFormatting;

export const tabs = (node: PMNode): TabStop[] => node.attrs["tabs"] as TabStop[];
