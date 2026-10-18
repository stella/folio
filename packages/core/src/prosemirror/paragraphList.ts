import { levelShowsMarker, resolveNumberingLevel } from "@stll/docx-core/model";
import { panic } from "better-result";
import { computeListRendering, type NumberingMap } from "../docx/numberingParser";
import type { ListRendering } from "../types/document";
import { effectiveParagraphNumbering, type ParagraphNumberingSources } from "./numberingAttr";

type ParagraphListReference = { kind: "reference"; numId: number; ilvl: number };

export type ParagraphList =
  | { kind: "none" }
  | {
      kind: "markerless";
      reference: ParagraphListReference;
      rendering?: ListRendering;
      reason:
        | "noNumbering"
        | "levelOutOfRange"
        | "noInstance"
        | "noAbstract"
        | "unresolvedStyleLink"
        | "levelNotDefined"
        | "markerHiddenOrBlank";
    }
  | { kind: "marker"; reference: ParagraphListReference; rendering: ListRendering };

/**
 * Resolves paragraph list membership, level, and marker visibility once for
 * every reader. Authored numbering attrs select the reference; package
 * definitions alone decide whether that reference paints a marker.
 */
export const resolveParagraphList = (
  sources: ParagraphNumberingSources,
  numbering: NumberingMap | null,
): ParagraphList => {
  const effective = effectiveParagraphNumbering(sources);
  switch (effective.kind) {
    case "none":
      return { kind: "none" };
    case "reference": {
      const reference: ParagraphListReference = {
        kind: "reference",
        numId: effective.numId,
        ilvl: effective.ilvl,
      };
      if (numbering === null) return { kind: "markerless", reference, reason: "noNumbering" };
      const resolution = resolveNumberingLevel(
        numbering.definitions,
        reference.numId,
        reference.ilvl,
      );
      if (resolution.kind === "undefined") {
        return { kind: "markerless", reference, reason: resolution.reason };
      }
      if (!levelShowsMarker(numbering.definitions, reference.numId, reference.ilvl)) {
        const rendering = computeListRendering(reference, numbering);
        return {
          kind: "markerless",
          reference,
          reason: "markerHiddenOrBlank",
          ...(rendering === null ? {} : { rendering }),
        };
      }
      const rendering = computeListRendering(reference, numbering);
      return rendering === null
        ? { kind: "markerless", reference, reason: "markerHiddenOrBlank" }
        : { kind: "marker", reference, rendering };
    }
    default: {
      const unhandled: never = effective;
      return panic(`Unhandled paragraph numbering ${JSON.stringify(unhandled)}`);
    }
  }
};
