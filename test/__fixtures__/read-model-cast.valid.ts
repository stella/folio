// Fixture for `folio-read-model/no-model-document-cast`. None of these casts
// a model to the model `Document`.

import type { Document } from "@stll/folio-core/types/document";
import {
  type ReadonlyDocument,
  readModel,
} from "@stll/folio-core/types/documentModel";

// The sanctioned reader cast.
export const read = (model: ReadonlyDocument): Document => readModel(model);

// A checked literal, not a cast.
export const literal = {
  package: { document: { content: [] } },
} satisfies Document;

// The DOM `Document` is not the model type.
export const domDocument = (value: object): globalThis.Document =>
  value as unknown as globalThis.Document;
