import {
  finalizeParagraphFormatting,
  type Paragraph,
  type ParagraphFormatting,
  type ParagraphPropertyChange,
} from "../model/document";
import { structurallyEqual } from "./equality";

/**
 * Where a registered snapshot comes from: a projector that put effective
 * values over the stated ones, or an edit that gave the paragraph exactly
 * the properties it states.
 */
type SnapshotProvenance = "projected" | "stated";

const snapshots = new WeakMap<
  object,
  {
    effective: ParagraphFormatting | undefined;
    snapshot: ParagraphFormatting | undefined;
    provenance: SnapshotProvenance;
  }
>();

type RegisterParagraphPropertySnapshotOptions = {
  /** A paragraph, or review fields an operation gives one. */
  paragraph: Pick<Paragraph, "formatting">;
  snapshot: ParagraphFormatting | undefined;
};

/** Bind effective formatting to the authored properties a review must record. */
export const registerParagraphPropertySnapshot = ({
  paragraph,
  snapshot,
}: RegisterParagraphPropertySnapshotOptions): void => {
  const current = snapshots.get(paragraph);
  if (
    current?.provenance === "projected" &&
    structurallyEqual(current.effective, paragraph.formatting) &&
    structurallyEqual(current.snapshot, snapshot)
  )
    return;
  snapshots.set(paragraph, {
    effective: paragraph.formatting,
    snapshot,
    provenance: "projected",
  });
};

/**
 * Register `paragraph`'s formatting as exactly what it states, as a rejected
 * property change restores the recorded properties: no list level or style
 * projects a value onto it.
 */
export const registerStatedParagraphFormatting = (
  paragraph: Pick<Paragraph, "formatting">,
): void => {
  snapshots.set(paragraph, {
    effective: paragraph.formatting,
    snapshot: paragraph.formatting,
    provenance: "stated",
  });
};

/**
 * The registration a paragraph's snapshot stands on, as an opaque token: the
 * same object while nothing registers or transfers another one for it. A pass
 * that registered a snapshot can tell from it that registering again would
 * change nothing.
 */
export const paragraphPropertySnapshotToken = (
  paragraph: Paragraph,
): object | undefined => snapshots.get(paragraph);

type TransferParagraphPropertySnapshotOptions = {
  source: object;
  target: object;
};

/** Retain paragraph-owned snapshot bindings through sanctioned model copies. */
export const transferParagraphPropertySnapshot = ({
  source,
  target,
}: TransferParagraphPropertySnapshotOptions): void => {
  const registered = snapshots.get(source);
  if (registered !== undefined) snapshots.set(target, registered);
};

/** Unprojected model formatting already states its authored properties. */
export const getParagraphPropertySnapshot = (
  paragraph: Paragraph,
): ParagraphFormatting | undefined => {
  const registered = snapshots.get(paragraph);
  return registered !== undefined &&
    structurallyEqual(paragraph.formatting, registered.effective)
    ? registered.snapshot
    : paragraph.formatting;
};

/**
 * Whether `paragraph`'s formatting is still the formatting an edit registered
 * as stated. A projector that infers stated properties from the effective
 * ones it projected before keeps every value such formatting holds.
 */
export const paragraphStatesItsFormatting = (
  paragraph: Pick<Paragraph, "formatting">,
): boolean => {
  const registered = snapshots.get(paragraph);
  return (
    registered?.provenance === "stated" &&
    structurallyEqual(paragraph.formatting, registered.effective)
  );
};

/**
 * The properties a paragraph states in its own `w:pPr`, as a reopened file
 * reads them: never what it inherits or what list numbering projects onto it.
 * A pending property change records these as the paragraph's current ones,
 * whichever edit, tracked or direct, last gave the paragraph properties.
 */
export const authoredParagraphFormatting = (
  paragraph: Paragraph,
): ParagraphFormatting | undefined =>
  finalizeParagraphFormatting(getParagraphPropertySnapshot(paragraph));

/**
 * `changes`, pending on `paragraph`, each recording what the paragraph now
 * states as its current properties.
 */
export const changesRecordingAuthored = (
  changes: readonly ParagraphPropertyChange[],
  paragraph: Paragraph,
): ParagraphPropertyChange[] => {
  const authored = authoredParagraphFormatting(paragraph);
  return changes.map((change) => {
    const refreshed = { ...change };
    if (authored === undefined) delete refreshed.currentFormatting;
    else refreshed.currentFormatting = authored;
    return refreshed;
  });
};
