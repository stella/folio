// Fixture for `folio-read-model/no-read-only-attr-cast`. None of these drops
// the read-only typing of a style-filled attr.

import type { Node as PMNode } from "prosemirror-model";

import type { ParagraphAttrs } from "@stll/folio-core/prosemirror/schema/nodes";
import type { TextFormatting } from "@stll/folio-core/types/document";
import {
  type DeepReadonly,
  writableCopy,
} from "@stll/folio-core/types/documentModel";

export const read = (attrs: ParagraphAttrs) => attrs.defaultTextFormatting;

export const readOnlyCast = (
  node: PMNode,
): DeepReadonly<TextFormatting> | undefined =>
  node.attrs["defaultTextFormatting"] as
    | DeepReadonly<TextFormatting>
    | undefined;

export const copied = (attrs: ParagraphAttrs): TextFormatting =>
  writableCopy<TextFormatting>(attrs.defaultTextFormatting ?? {});

// A table's own `borders` field that is not a node's attrs.
export const tableBorders = (formatting: { borders?: object }) =>
  formatting.borders as object;
