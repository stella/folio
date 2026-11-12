// Fixture for `folio-read-model/no-model-document-cast`. The rule must flag
// every cast to the model `Document`, however it is imported or spelled.

import type { Document as ModelDocument } from "../../packages/core/src/types/document";
import type * as Model from "@stll/folio-core/types/document";
import type { Document } from "@stll/folio-core/types/document";
import type { ReadonlyDocument } from "@stll/folio-core/types/documentModel";

export const direct = (model: ReadonlyDocument): Document => model as Document;

export const throughUnknown = (model: ReadonlyDocument): Document =>
  model as unknown as Document;

export const asUnion = (model: ReadonlyDocument | null): Document | null =>
  model as Document | null;

export const renamed = (model: ReadonlyDocument): ModelDocument =>
  model as ModelDocument;

export const asIntersection = (
  model: ReadonlyDocument,
): Document & { marker?: string } => model as Document & { marker?: string };

export const throughNamespace = (model: ReadonlyDocument): Model.Document =>
  model as Model.Document;

type WritableDocument = Document;
type MaybeWritable = WritableDocument | null;

export const throughAlias = (model: ReadonlyDocument): WritableDocument =>
  model as WritableDocument;

export const throughAliasChain = (
  model: ReadonlyDocument | null,
): MaybeWritable => model as MaybeWritable;
