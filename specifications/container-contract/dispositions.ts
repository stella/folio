/**
 * What folio promises to do with each thing the schema allows in a container.
 *
 * The contract is the declaration; the survival census is the execution that
 * checks it. A pair declared `modelled` or `captured-verbatim` that the census
 * shows lost fails CI, and a pair declared `dropped` that the census shows
 * surviving fails CI too, so the declaration cannot drift away from the code in
 * either direction.
 *
 * These maps decide something about `@stll/folio-core`'s parsers, but nothing
 * that imports the package needs them: they are read by the compiler here, by
 * the contract generator and by the agreement check. They therefore live in a
 * tsconfig project of their own, budgeted like any other, the way
 * `specifications/reserved-values` does.
 */

export const DISPOSITIONS = {
  /**
   * A parser reads it into the typed model and a serializer writes it back.
   *
   * The strongest promise: the value survives an edit, a round trip through
   * the editor, and a rebuild of the container from the model alone.
   */
  modelled: "modelled",
  /**
   * The markup is kept as bytes and replayed.
   *
   * Nothing understands it, so nothing can edit it, but nothing loses it
   * either. This is the right disposition for a historical record such as a
   * tracked-change snapshot, and the wrong one for anything a user edits.
   */
  capturedVerbatim: "captured-verbatim",
  /** folio does not keep it. Every entry names why, from {@link DROP_REASONS}. */
  dropped: "dropped",
} as const;

export type Disposition = (typeof DISPOSITIONS)[keyof typeof DISPOSITIONS];

/**
 * Why a dropped pair is dropped, and what would have to change to keep it.
 *
 * A reason is a class rather than prose per pair: 1500 bespoke sentences would
 * be 1500 places to let a stale one sit. Each class names one mechanism and one
 * kind of fix, and the generator refuses a `dropped` entry whose reason is not
 * one of these.
 */
export const DROP_REASONS = {
  containerNotKept: {
    summary: "The container is not kept, so nothing inside it can be.",
    fix: "Give the container a disposition first; every pair under it follows.",
  },
  neverParsed: {
    summary: "No parser reads it, and no capture keeps its markup.",
    fix: "Route the container's children through the shared dispatcher, whose default is an ordered verbatim sink.",
  },
  parsedNotSerialized: {
    summary: "A parser reads it into the model and no serializer writes it back.",
    fix: "Write it from the model in the container's serializer.",
  },
  replayOnly: {
    summary:
      "It survives only while verbatim replay hands the container's markup back, so an edit loses it.",
    fix: "Give it a capture slot of its own, or model it, and write it on the rebuild path.",
  },
  replayRejected: {
    summary:
      "A capture holds it and a replay gate refuses that capture, forcing a rebuild that cannot write it.",
    fix: "Model the construct and let the gate accept the capture, as `w:numberingChange` now does.",
  },
  editorProjection: {
    summary:
      "It survives a save but not the ProseMirror projection, so an edited document loses it.",
    fix: "Carry it through `toProseDoc`/`fromProseDoc`, on the node or as a projection-only attribute.",
  },
  repeatTruncated: {
    summary:
      "The slot comes back and some of the instances written into it do not, so a repeated particle is silently shortened.",
    fix: "Read and write the particle as the list the schema declares it to be, rather than as its first instance.",
  },
  respelled: {
    summary: "It comes back with a different value.",
    fix: "Trace the rewrite to the reader or writer that performs it; a spelling change folio makes by design belongs in the law's value equivalence instead.",
  },
  parserThrows: {
    summary: "The construct makes the parser throw, so the document does not open at all.",
    fix: "Make the parser tolerant of it; a refusal is a worse loss than a drop.",
  },
} as const;

export type DropReason = keyof typeof DROP_REASONS;

export type ContractEntry =
  | { disposition: typeof DISPOSITIONS.modelled }
  | { disposition: typeof DISPOSITIONS.capturedVerbatim }
  | { disposition: typeof DISPOSITIONS.dropped; reason: DropReason };

/** The committed contract, keyed by the pair key `schemaSpace` derives. */
export type ContainerContract = {
  schemaVersion: 1;
  entries: Record<string, ContractEntry>;
};
