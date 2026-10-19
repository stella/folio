import {
  levelShowsMarker,
  paragraphNumberingReference,
  resolveNumberingLevel,
  type ResolvedParagraphNumbering,
} from "@stll/docx-core/model";
import { panic } from "better-result";
import { computeListRendering, type NumberingMap } from "../docx/numberingParser";
import type { ListRendering } from "../types/document";
import { effectiveParagraphNumbering, type ParagraphNumberingSources } from "./numberingAttr";

type ParagraphListReference = Extract<ResolvedParagraphNumbering, { kind: "reference" }>;

const standaloneList = (
  reference: ParagraphListReference,
  rendering: ListRendering,
): ParagraphList =>
  rendering.markerHidden === true || rendering.marker.trim().length === 0
    ? {
        kind: "markerless",
        reference,
        rendering,
        reason: "markerHiddenOrBlank",
      }
    : { kind: "marker", reference, rendering };

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
  sources: ParagraphNumberingSources & {
    /** `null` means the current style was resolved and supplies no numbering. */
    inheritedReference?: ParagraphListReference | null;
  },
  numbering: NumberingMap | null,
  standaloneRendering?: ListRendering,
): ParagraphList => {
  const resolved = effectiveParagraphNumbering({
    numPr: sources.numPr,
    numPrFromStyle:
      sources.inheritedReference === undefined
        ? sources.numPrFromStyle
        : (sources.inheritedReference ?? undefined),
  });
  const effective =
    resolved.kind === "none" &&
    sources.numPr === undefined &&
    sources.numPrFromStyle === undefined &&
    standaloneRendering !== undefined
      ? paragraphNumberingReference({
          numId: standaloneRendering.numId,
          ilvl: standaloneRendering.level,
        })
      : resolved;
  switch (effective.kind) {
    case "none": {
      // A package-free canonical model can carry its list projection directly.
      if (numbering === null && standaloneRendering !== undefined) {
        const reference = paragraphNumberingReference({
          numId: standaloneRendering.numId,
          ilvl: standaloneRendering.level,
        });
        return standaloneList(reference, standaloneRendering);
      }
      return { kind: "none" };
    }
    case "reference": {
      const reference = effective;
      if (numbering === null) {
        return standaloneRendering === undefined
          ? { kind: "markerless", reference, reason: "noNumbering" }
          : standaloneList(reference, standaloneRendering);
      }
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
