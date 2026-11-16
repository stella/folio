import type { SaveDiagnosticOptions } from "../docx/saveDiagnostics";

export const FOLIO_DOCX_SERIALIZATION_MODE: Readonly<{
  readonly full: "full";
  readonly preferSelective: "prefer-selective";
}> = Object.freeze({
  full: "full",
  preferSelective: "prefer-selective",
} as const);

export type FolioDocxSerializationMode =
  (typeof FOLIO_DOCX_SERIALIZATION_MODE)[keyof typeof FOLIO_DOCX_SERIALIZATION_MODE];

/** How the serializer produced the returned package bytes. */
export type FolioDocxSerializationPath = "selective" | "full-repack";

export type FolioGetDocxOptions = {
  /** Fidelity diagnostics for the exact snapshot being serialized. */
  onDiagnostic?: SaveDiagnosticOptions["onDiagnostic"];
  /** The condition that declined an attempted selective write. */
  onSelectiveDecline?: SaveDiagnosticOptions["onSelectiveDecline"];
  /** The actual path that produced these bytes, after successful serialization. */
  onSerializationPath?:
    ((path: FolioDocxSerializationPath) => void) | undefined;
  mode?: FolioDocxSerializationMode;
};

/** Host-facing save options share the serializer's callbacks. */
export type FolioSaveOptions = Omit<FolioGetDocxOptions, "mode"> & {
  selective?: boolean;
};
