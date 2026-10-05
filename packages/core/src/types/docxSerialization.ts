import type { SaveDiagnosticOptions } from "../docx/saveDiagnostics";

export const FOLIO_DOCX_SERIALIZATION_MODE = Object.freeze({
  full: "full",
  preferSelective: "prefer-selective",
} as const);

export type FolioDocxSerializationMode =
  (typeof FOLIO_DOCX_SERIALIZATION_MODE)[keyof typeof FOLIO_DOCX_SERIALIZATION_MODE];

export type FolioGetDocxOptions = {
  /** Fidelity diagnostics for the exact snapshot being serialized. */
  onDiagnostic?: SaveDiagnosticOptions["onDiagnostic"];
  mode?: FolioDocxSerializationMode;
};
