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
export const TRACKED_RUN_INLINE_ATOM_DISPOSITIONS = Object.freeze({
  bookmarkBoundary: "carry",
  // The model has no slot for a comment reference inside a tracked wrapper:
  // `w:commentReference` is paragraph content, so a revision cannot own it.
  commentReference: "transparent",
  field: "field-carrier",
  hardBreak: "carry",
  image: "carry",
  math: "carry",
  pageBreakRun: "page-break-carrier",
  preservedXml: "carry",
  // A range marker is not content a revision owns: `TRACKED_CHANGE_WRAPPER_CONTENT`
  // admits neither half, so the model keeps an empty range beside the wrapper.
  rangeAnchor: "transparent",
  renderedPageBreak: "transparent",
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
    panic(`Inline atom ${JSON.stringify(node.type.name)} has no tracked-run disposition`);
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
