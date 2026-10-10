import { effectiveParagraphNumbering } from "./numberingAttr";
/**
 * The list rendering (marker, indentation bookkeeping) a paragraph's
 * numbering gives it, as a paragraph-property change records and restores it.
 */

import type { NumberingMap } from "../docx/numberingParser";
import {
  type ParagraphNumberingOverride,
  paragraphNumberingLevel,
  paragraphNumberingReferenceId,
  sameStatedParagraphNumbering,
} from "../docx/numberingReference";
import { CLEARED_LIST_RENDERING_ATTRS, LIST_RENDERING_ATTR_KEYS } from "./listMarker";
import type { ListRenderingAttrKey } from "./listRenderingAttrs";
import type { ParagraphAttrs, ParagraphPropertyChangeAttrs } from "./schema/nodes";
import { listAttrsFromNumbering } from "./styles/resolvedStyleAttrs";

type PreviousFormatting = NonNullable<ParagraphPropertyChangeAttrs["previousFormatting"]>;

/** These attrs come from inline content, not the numbering definition. */
const CONTENT_DERIVED_LIST_RENDERING_ATTRS = new Set<ListRenderingAttrKey>([
  "listImplicitChildLevelAdvances",
  "listMarkerSecondSlotOffsetTwips",
]);

/** Whether a `w:pPrChange` record states the list rendering it had (a list command's does). */
export const recordsListRendering = (record: PreviousFormatting): boolean =>
  LIST_RENDERING_ATTR_KEYS.some((key) => Object.hasOwn(record, key));

/** Every list-rendering attr for `numPr`: its level's, or all cleared when it numbers nothing. */
export const listRenderingFor = (
  numPr: ParagraphNumberingOverride | null | undefined,
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
 * the marker it was given. Reconcile the rendering attrs with the restored
 * numbering even when the effective reference itself did not change.
 */
export type RejectedListRenderingOptions = {
  /** The paragraph's attrs before the reject. */
  current: Pick<ParagraphAttrs, "numPr" | "numPrFromStyle" | ListRenderingAttrKey>;
  /** The previous state the rejected record restores. */
  previousFormatting: PreviousFormatting | null | undefined;
  numbering: NumberingMap | null;
  restoredNumbering: ParagraphAttrs["numPr"] | null;
};

export const rejectedListRenderingPatch = ({
  current,
  previousFormatting,
  numbering,
  restoredNumbering,
}: RejectedListRenderingOptions): Record<string, unknown> => {
  const record = previousFormatting ?? {};
  if (recordsListRendering(record)) {
    return {};
  }
  const restored = restoredNumbering ?? undefined;
  const sameEffectiveNumbering = sameStatedParagraphNumbering(
    restored,
    effectiveParagraphNumbering(current),
  );
  if (!sameEffectiveNumbering) return listRenderingFor(restored, numbering);

  const expected = listRenderingFor(restored, numbering);
  const patch: Record<string, unknown> = {};
  for (const key of LIST_RENDERING_ATTR_KEYS) {
    if (CONTENT_DERIVED_LIST_RENDERING_ATTRS.has(key)) continue;
    const currentValue = current[key] ?? null;
    const expectedValue = expected[key] ?? null;
    if (JSON.stringify(currentValue) !== JSON.stringify(expectedValue)) {
      patch[key] = expectedValue;
    }
  }
  return Object.keys(patch).length === 0 ? {} : patch;
};
