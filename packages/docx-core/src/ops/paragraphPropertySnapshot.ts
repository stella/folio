import {
  finalizeParagraphFormatting,
  type Paragraph,
  type ParagraphFormatting,
  type ParagraphPropertyChange,
} from "../model/document";
import { structurallyEqual } from "./equality";

const snapshots = new WeakMap<
  object,
  {
    effective: ParagraphFormatting | undefined;
    snapshot: ParagraphFormatting | undefined;
  }
>();

type RegisterParagraphPropertySnapshotOptions = {
  paragraph: Paragraph;
  snapshot: ParagraphFormatting | undefined;
};

/** Bind effective formatting to the authored properties a review must record. */
export const registerParagraphPropertySnapshot = ({
  paragraph,
  snapshot,
}: RegisterParagraphPropertySnapshotOptions): void => {
  const current = snapshots.get(paragraph);
  if (
    current !== undefined &&
    structurallyEqual(current.effective, paragraph.formatting) &&
    structurallyEqual(current.snapshot, snapshot)
  )
    return;
  snapshots.set(paragraph, { effective: paragraph.formatting, snapshot });
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
