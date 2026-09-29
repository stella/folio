/**
 * The document's own mark types.
 *
 * An extension sees the schema its runtime was built with, but the editor it
 * runs in may hold a document of another schema instance with the same specs:
 * header, footer and note editors each build a runtime of their own, while
 * every document is parsed into the shared schema. Marks made from the
 * runtime's types do not belong to such a document, and a runtime type
 * compared against the document's marks never matches. Commands therefore
 * resolve the types they build with, and compare against, through the state
 * they act on.
 */

import { panic } from "better-result";
import type { MarkType, Schema } from "prosemirror-model";

/** `type`'s counterpart in the schema of `state`. */
export const documentMarkType = (state: { readonly schema: Schema }, type: MarkType): MarkType =>
  type.schema === state.schema
    ? type
    : (state.schema.marks[type.name] ?? panic(`The document schema has no mark ${type.name}`));
