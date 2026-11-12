import type { Document } from "./document";

/** `Value` with every reachable field, element and entry read-only. */
export type DeepReadonly<Value> = Value extends (...args: never[]) => unknown
  ? Value
  : Value extends ArrayBuffer | Date
    ? Value
    : Value extends ReadonlyMap<infer Key, infer Entry>
      ? ReadonlyMap<DeepReadonly<Key>, DeepReadonly<Entry>>
      : Value extends ReadonlySet<infer Entry>
        ? ReadonlySet<DeepReadonly<Entry>>
        : Value extends readonly (infer Element)[]
          ? readonly DeepReadonly<Element>[]
          : Value extends object
            ? { readonly [Key in keyof Value]: DeepReadonly<Value[Key]> }
            : Value;

/**
 * The editor's document model: in a canonical session the session's own
 * document, which the session never writes again once it has handed it out,
 * so a host may hold it as long as it likes; it must never write to it.
 */
export type ReadonlyDocument = DeepReadonly<Document>;

/**
 * The model as a `Document`, for a reader that takes one: the style resolver
 * and its `resolve*` results, the numbering map, the hidden editor's state and
 * the comment, find and table-selection readers. Readers never write to the
 * document they are given, and a deep-freeze test runs each one a host calls
 * over frozen documents, so the model is passed to them as it is, without a
 * copy. This is the only place a model is read as a plain `Document`; a writer
 * copies the model first (`cloneDocumentWithParagraphPropertySources`) and
 * writes the copy.
 */
export const readModel = (model: ReadonlyDocument): Document =>
  // SAFETY: the callers are readers (see above); the type only drops `readonly`.
  model as Document;
