import { panic } from "better-result";

import {
  levelShowsMarker,
  paragraphNumberingReference,
  type ResolvedParagraphNumbering,
  resolveNumberingLevel,
} from "@stll/docx-core/model";

import {
  computeListRendering,
  type NumberingMap,
} from "../docx/numberingParser";
import type { ListRendering } from "../types/document";
import {
  effectiveParagraphNumbering,
  type ParagraphNumberingSources,
} from "./numberingAttr";

type ParagraphListReference = Extract<
  ResolvedParagraphNumbering,
  { kind: "reference" }
>;

const standaloneList = (
  reference: ParagraphListReference,
  rendering: ListRendering,
  source: ListReferenceSource,
): ParagraphList =>
  rendering.markerHidden === true || rendering.marker.trim().length === 0
    ? {
        kind: "markerless",
        reference,
        source,
        rendering,
        reason: "markerHiddenOrBlank",
      }
    : { kind: "marker", reference, source, rendering };

/**
 * `projection`: only the paragraph's cached list projection made it a list
 * (a package-free model), so its reference names no authored numbering
 * instance and readers must not report one.
 */
type ListReferenceSource = "authored" | "projection";

export type ParagraphList =
  | { kind: "none" }
  | {
      kind: "markerless";
      reference: ParagraphListReference;
      source: ListReferenceSource;
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
  | {
      kind: "marker";
      reference: ParagraphListReference;
      source: ListReferenceSource;
      rendering: ListRendering;
    };

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
  // Only a package-free canonical model carries its list projection directly,
  // and only while the paragraph's own numbering does not cancel the list.
  // With a package, the definitions alone decide.
  const projectionStandsIn =
    numbering === null &&
    standaloneRendering !== undefined &&
    sources.numPr?.kind !== "none";
  switch (resolved.kind) {
    case "none": {
      if (projectionStandsIn) {
        const reference = paragraphNumberingReference({
          numId: standaloneRendering.numId,
          ilvl: standaloneRendering.level,
        });
        return standaloneList(reference, standaloneRendering, "projection");
      }
      return { kind: "none" };
    }
    case "reference": {
      const reference = resolved;
      if (numbering === null) {
        return standaloneRendering === undefined
          ? {
              kind: "markerless",
              reference,
              source: "authored",
              reason: "noNumbering",
            }
          : standaloneList(reference, standaloneRendering, "authored");
      }
      const resolution = resolveNumberingLevel(
        numbering.definitions,
        reference.numId,
        reference.ilvl,
      );
      if (resolution.kind === "undefined") {
        return {
          kind: "markerless",
          reference,
          source: "authored",
          reason: resolution.reason,
        };
      }
      if (
        !levelShowsMarker(
          numbering.definitions,
          reference.numId,
          reference.ilvl,
        )
      ) {
        const rendering = computeListRendering(reference, numbering);
        return {
          kind: "markerless",
          reference,
          source: "authored",
          reason: "markerHiddenOrBlank",
          ...(rendering === null ? {} : { rendering }),
        };
      }
      const rendering = computeListRendering(reference, numbering);
      return rendering === null
        ? {
            kind: "markerless",
            reference,
            source: "authored",
            reason: "markerHiddenOrBlank",
          }
        : { kind: "marker", reference, source: "authored", rendering };
    }
    default: {
      const unhandled: never = resolved;
      return panic(
        `Unhandled paragraph numbering ${JSON.stringify(unhandled)}`,
      );
    }
  }
};
