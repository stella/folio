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
   * Note paragraphs `touched` since the save baseline (or since the session
   * opened, without one), and those whose content differs from the opened
   * document. A note part is written against the opened package, so a note
   * an edit and its undo left as it was is written as the source held it.
   */
  noteParagraphs: {
    touched: readonly string[];
    changedSinceSource: readonly string[];
  };
  /**
   * Paragraphs touched since the session opened whose content is again what
   * the opened document held. A save against the opened package writes them
   * as the source held them.
   */
  restoredSinceSource: readonly string[];
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
