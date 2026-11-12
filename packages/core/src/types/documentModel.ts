import type { DeepReadonly } from "@stll/docx-core/model";

import type { Document } from "./document";

export type { DeepReadonly };

/**
 * The editor's document model: in a canonical session the session's own
 * document, which the session never writes again once it has handed it out,
 * so a host may hold it as long as it likes; it must never write to it.
 */
export type ReadonlyDocument = DeepReadonly<Document>;

/**
 * A record a function builds itself out of read-only values: its own fields
 * are the builder's to set, while the values in them stay read-only, shared
 * with whatever they were read from.
 */
export type OwnRecord<Value> = {
  -readonly [Key in keyof DeepReadonly<Value>]: DeepReadonly<Value>[Key];
};

/**
 * The model as a `Document`, for a reader that takes one: the numbering map,
 * the hidden editor's state and the comment, find and table-selection readers.
 * The style resolver takes the model's styles as they are and hands out
 * read-only results, so it needs no `readModel`. Readers never write to the
 * document they are given, and a deep-freeze test runs each one a host calls
 * over frozen documents, so the model is passed to them as it is, without a
 * copy. This is the only place a model is read as a plain `Document`; a writer
 * copies the model first (`cloneDocumentWithParagraphPropertySources`) and
 * writes the copy.
 */
export const readModel = (model: ReadonlyDocument): Document =>
  // SAFETY: the callers are readers (see above); the type only drops `readonly`.
  model as Document;

/**
 * A writable copy of a read-only value: the one way a writer gets something it
 * may change out of a value a reader handed out, such as a style resolver's
 * results or the style-derived attributes of a paragraph. A structured clone
 * shares nothing with its source, so writing the copy never reaches the
 * document or a cache. A whole document is copied with
 * `cloneDocumentWithParagraphPropertySources`, which also keeps the paragraph
 * identities a structured clone drops.
 */
export const writableCopy = <Value>(value: DeepReadonly<Value>): Value =>
  // SAFETY: the clone shares nothing with `value`, so the writable type
  // cannot reach the value it was copied from.
  structuredClone(value) as Value;
