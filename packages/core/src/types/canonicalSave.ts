import type { Document } from "./document";

/** Committed model and conservative changes relative to its parsed source package. */
export type CanonicalSaveSnapshot = {
  document: Document;
  version: number;
  changedBlockIds: readonly string[];
  structure: "stable" | "changed";
  /**
   * The notes the opened package held without a reference. They are its
   * content and survive the save; only a note an edit left without one goes.
   */
  heldUnreferencedNotes: ReadonlySet<string>;
  /**
   * The package the session last saved, when it is the one this save
   * patches instead of the opened package.
   */
  saveBaseline?: CanonicalSaveSnapshotBaseline;
};

/** A saved package's buffer and source captures, without its model. */
export type CanonicalSaveSnapshotBaseline = {
  buffer: ArrayBuffer;
  document: Document;
};
