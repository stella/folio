import { panic } from "better-result";
import type { Node as PMNode } from "prosemirror-model";

import { decodeOoxmlSymbolCharacter } from "../utils/ooxmlSymbol";
import { expectFieldAttrs, expectSymbolAttrs } from "./attrs";

export type RunFormattingInlineAtomDisposition =
  | "break-run"
  | "field-run"
  | "not-a-run"
  | "page-break-carrier"
  | "structured-field"
  | "symbol-run"
  | "tab-run"
  | "text-run";

/**
 * How every inline atom maps to a serialized run-formatting carrier.
 *
 * This is intentionally separate from tracked wrapper ownership: drawings and
 * bookmarks can sit inside an insertion/deletion wrapper, but their editor
 * marks do not serialize as run properties. The runtime fixed-point test binds
 * this map to the complete schema atom set.
 */
export const RUN_FORMATTING_INLINE_ATOM_DISPOSITIONS = Object.freeze({
  bookmarkBoundary: "not-a-run",
  field: "field-run",
  hardBreak: "break-run",
  image: "not-a-run",
  math: "not-a-run",
  pageBreakRun: "page-break-carrier",
  renderedPageBreak: "not-a-run",
  shape: "not-a-run",
  structuredField: "structured-field",
  symbol: "symbol-run",
  tab: "tab-run",
  text: "text-run",
  textBoxAnchor: "not-a-run",
} as const satisfies Readonly<Record<string, RunFormattingInlineAtomDisposition>>);

const dispositionsByName = new Map<string, RunFormattingInlineAtomDisposition>(
  Object.entries(RUN_FORMATTING_INLINE_ATOM_DISPOSITIONS),
);

export const runFormattingInlineAtomDisposition = (
  node: PMNode,
): RunFormattingInlineAtomDisposition | null => {
  if (!node.isInline || !node.isAtom) {
    return null;
  }

  const disposition = dispositionsByName.get(node.type.name);
  if (!disposition) {
    return panic(`Inline atom ${JSON.stringify(node.type.name)} has no run-formatting disposition`);
  }
  return disposition;
};

export type RunFormattingCarrierRepresentation = {
  node: PMNode;
  position: number;
  role: "owner" | "serialized-result";
};

export type RunFormattingCarrier = {
  disposition: Exclude<RunFormattingInlineAtomDisposition, "not-a-run">;
  node: PMNode;
  position: number;
  representations: readonly RunFormattingCarrierRepresentation[];
};

export type SelectedRunFormattingCarrierRepresentation = RunFormattingCarrierRepresentation & {
  carrierDisposition: RunFormattingCarrier["disposition"];
  from: number;
  to: number;
};

/**
 * Expand one logical carrier into every editor node that materializes its run
 * properties. A structured field is owned by its outer atom, while its result
 * runs are derived representations needed by the package serializer.
 */
export const expandRunFormattingCarrier = (
  node: PMNode,
  position: number,
): RunFormattingCarrier | null => {
  const disposition = runFormattingInlineAtomDisposition(node);
  if (disposition === null || disposition === "not-a-run") {
    return null;
  }
  if (disposition !== "structured-field") {
    return {
      disposition,
      node,
      position,
      representations: [{ node, position, role: "owner" }],
    };
  }

  const representations: RunFormattingCarrierRepresentation[] = [{ node, position, role: "owner" }];
  node.descendants((child, relativePosition) => {
    const childDisposition = runFormattingInlineAtomDisposition(child);
    if (childDisposition === null) {
      return true;
    }
    if (childDisposition === "structured-field") {
      return panic("A structured field cannot contain another structured field");
    }
    if (childDisposition !== "not-a-run") {
      representations.push({
        node: child,
        position: position + 1 + relativePosition,
        role: "serialized-result",
      });
    }
    return false;
  });
  return { disposition, node, position, representations };
};

type SelectRunFormattingCarrierRepresentationsOptions = {
  doc: PMNode;
  from: number;
  to: number;
};

/**
 * Resolve a document range to the exact editor nodes that serialize run
 * properties. The result is shared by every formatting mutation so preflight,
 * direct edits, and tracked edits cannot disagree about ownership.
 */
export const selectRunFormattingCarrierRepresentations = ({
  doc,
  from,
  to,
}: SelectRunFormattingCarrierRepresentationsOptions): SelectedRunFormattingCarrierRepresentation[] => {
  const selected: SelectedRunFormattingCarrierRepresentation[] = [];
  doc.nodesBetween(from, to, (node, position) => {
    if (!node.isInline) {
      return true;
    }
    const carrier = expandRunFormattingCarrier(node, position);
    if (!carrier) {
      return !node.isAtom;
    }
    if (carrier.disposition === "page-break-carrier") {
      return false;
    }

    for (const representation of carrier.representations) {
      if (runFormattingInlineAtomDisposition(representation.node) === "page-break-carrier") {
        continue;
      }
      const representationFrom = Math.max(from, representation.position);
      const representationTo = Math.min(to, representation.position + representation.node.nodeSize);
      if (representationFrom >= representationTo) {
        continue;
      }
      if (
        carrier.disposition === "structured-field" &&
        representation.role === "owner" &&
        (representationFrom !== representation.position ||
          representationTo !== representation.position + representation.node.nodeSize)
      ) {
        continue;
      }
      selected.push({
        ...representation,
        carrierDisposition: carrier.disposition,
        from: representationFrom,
        to: representationTo,
      });
    }
    return false;
  });
  return selected;
};

/** Deterministic visible text for one logical formatting revision carrier. */
export const runFormattingCarrierReviewText = (carrier: RunFormattingCarrier): string => {
  switch (carrier.disposition) {
    case "text-run":
      return carrier.node.text ?? "";
    case "tab-run":
      return "\t";
    case "break-run":
      return "\n";
    case "page-break-carrier":
      return "";
    case "symbol-run": {
      const { char } = expectSymbolAttrs(carrier.node);
      return decodeOoxmlSymbolCharacter(char) ?? "\uFFFD";
    }
    case "field-run":
    case "structured-field":
      return expectFieldAttrs(carrier.node).displayText ?? carrier.node.textContent;
  }
};
