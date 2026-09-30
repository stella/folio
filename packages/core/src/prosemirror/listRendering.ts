/**
 * The list rendering (marker, indentation bookkeeping) a paragraph's
 * numbering gives it, as a paragraph-property change records and restores it.
 */

import type { NumberingMap } from "../docx/numberingParser";
import {
  paragraphNumberingLevel,
  paragraphNumberingReferenceId,
  sameStatedParagraphNumbering,
} from "../docx/numberingReference";
import { CLEARED_LIST_RENDERING_ATTRS, LIST_RENDERING_ATTR_KEYS } from "./listMarker";
import type { ParagraphAttrs, ParagraphPropertyChangeAttrs } from "./schema/nodes";
import { listAttrsFromNumbering } from "./styles/resolvedStyleAttrs";

type PreviousFormatting = NonNullable<ParagraphPropertyChangeAttrs["previousFormatting"]>;

/** Whether a `w:pPrChange` record states the list rendering it had (a list command's does). */
export const recordsListRendering = (record: PreviousFormatting): boolean =>
  LIST_RENDERING_ATTR_KEYS.some((key) => Object.hasOwn(record, key));

/** Every list-rendering attr for `numPr`: its level's, or all cleared when it numbers nothing. */
export const listRenderingFor = (
  numPr: PreviousFormatting["numPr"],
  numbering: NumberingMap | null,
): Record<string, unknown> => {
  const numId = paragraphNumberingReferenceId(numPr ?? undefined);
  const rendering =
    numId === undefined
      ? CLEARED_LIST_RENDERING_ATTRS
      : listAttrsFromNumbering(
          { numId, ilvl: paragraphNumberingLevel(numPr ?? undefined) ?? 0 },
          numbering,
        );
  const attrs: Record<string, unknown> = {};
  for (const key of LIST_RENDERING_ATTR_KEYS) {
    attrs[key] = rendering[key] ?? null;
  }
  return attrs;
};

/**
 * The list rendering a paragraph takes when a `w:pPrChange` is rejected. A
 * reject restores the recorded numbering, but the rendering attrs sit outside
 * the scope it restores wholesale, so a record that does not state them (one
 * written by an operation, or read from a file) would leave the rendering of
 * the numbering it undoes: a paragraph restored to no numbering still showed
 * the marker it was given. When the reject changes the numbering, the
 * rendering is recomputed from the restored numbering; otherwise it stays.
 */
export type RejectedListRenderingOptions = {
  /** The paragraph's attrs before the reject. */
  current: Pick<ParagraphAttrs, "numPr">;
  /** The previous state the rejected record restores. */
  previousFormatting: PreviousFormatting | null | undefined;
  numbering: NumberingMap | null;
};

export const rejectedListRenderingPatch = ({
  current,
  previousFormatting,
  numbering,
}: RejectedListRenderingOptions): Record<string, unknown> => {
  const record = previousFormatting ?? {};
  if (recordsListRendering(record)) {
    return {};
  }
  const restored = record.numPr ?? undefined;
  if (sameStatedParagraphNumbering(restored, current.numPr ?? undefined)) {
    return {};
  }
  return listRenderingFor(restored, numbering);
};
