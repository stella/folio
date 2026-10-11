import type { Paragraph, ParagraphFormatting } from "../model/document";
import { structurallyEqual } from "./equality";

const snapshots = new WeakMap<
  object,
  { effective: ParagraphFormatting | undefined; snapshot: ParagraphFormatting | undefined }
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

type TransferParagraphPropertySnapshotOptions = { source: object; target: object };

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
  return registered !== undefined && structurallyEqual(paragraph.formatting, registered.effective)
    ? registered.snapshot
    : paragraph.formatting;
};
