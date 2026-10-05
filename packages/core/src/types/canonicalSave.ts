import type { Document } from "./document";

/** Committed model and conservative changes relative to its parsed source package. */
export type CanonicalSaveSnapshot = {
  document: Document;
  version: number;
  changedBlockIds: readonly string[];
  structure: "stable" | "changed";
};
