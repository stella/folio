import { panic } from "better-result";
import type { Node as PMNode } from "prosemirror-model";

export type TrackedRunInlineAtomDisposition =
  | "carry"
  | "field-carrier"
  | "outside-wrapper"
  | "page-break-carrier"
  | "text-carrier"
  | "transparent";

/**
 * How every inline atom in the editor schema participates in an OOXML
 * run-level tracked-change wrapper.
 *
 * The schema is assembled at runtime, so the fixed-point test binds this map
 * to its complete inline-atom set. The lookup also panics on a missing entry:
 * a new atom cannot silently lose revision ownership before that test runs.
 */
export const TRACKED_RUN_INLINE_ATOM_DISPOSITIONS: Readonly<{
  readonly bookmarkBoundary: "carry";
  // The paragraph-level record stands for its w:r carrier, which can remain
  // in the CT_RunTrackChange that owns its authored position.
  readonly commentReference: "carry";
  readonly commentRangeBoundary: "carry";
  readonly field: "field-carrier";
  readonly hardBreak: "carry";
  readonly image: "carry";
  readonly math: "carry";
  readonly moveRangeBoundary: "carry";
  readonly noteMarker: "carry";
  readonly pageBreakRun: "page-break-carrier";
  readonly preservedXml: "carry";
  // A zero-width range owns its authored position inside the revision.
  readonly rangeAnchor: "carry";
  readonly renderedPageBreak: "carry";
  readonly shape: "carry";
  readonly structuredField: "carry";
  readonly symbol: "carry";
  readonly tab: "carry";
  readonly text: "text-carrier";
  readonly textBoxAnchor: "carry";
}> = Object.freeze({
  bookmarkBoundary: "carry",
  // The paragraph-level record stands for its w:r carrier, which can remain
  // in the CT_RunTrackChange that owns its authored position.
  commentReference: "carry",
  commentRangeBoundary: "carry",
  field: "field-carrier",
  hardBreak: "carry",
  image: "carry",
  math: "carry",
  moveRangeBoundary: "carry",
  noteMarker: "carry",
  pageBreakRun: "page-break-carrier",
  preservedXml: "carry",
  // A zero-width range owns its authored position inside the revision.
  rangeAnchor: "carry",
  renderedPageBreak: "carry",
  shape: "carry",
  structuredField: "carry",
  symbol: "carry",
  tab: "carry",
  text: "text-carrier",
  textBoxAnchor: "carry",
} as const satisfies Readonly<Record<string, TrackedRunInlineAtomDisposition>>);

const dispositionsByName = new Map<string, TrackedRunInlineAtomDisposition>(
  Object.entries(TRACKED_RUN_INLINE_ATOM_DISPOSITIONS),
);

export const trackedRunInlineAtomDisposition = (
  node: PMNode,
): TrackedRunInlineAtomDisposition | null => {
  if (!node.isInline || !node.isAtom) {
    return null;
  }

  const disposition = dispositionsByName.get(node.type.name);
  if (!disposition) {
    panic(
      `Inline atom ${JSON.stringify(node.type.name)} has no tracked-run disposition`,
    );
  }
  return disposition;
};

/** Whether an editor node can own an insertion/deletion mark as one run carrier. */
export const canCarryTrackedRunMark = (node: PMNode): boolean => {
  const disposition = trackedRunInlineAtomDisposition(node);
  return (
    disposition === "carry" ||
    disposition === "field-carrier" ||
    disposition === "text-carrier" ||
    disposition === "page-break-carrier"
  );
};
