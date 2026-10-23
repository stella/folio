/**
 * The number or bullet a paragraph shows beside its text, as a reader reports
 * it: counted in document order by the counter the page paints its markers
 * with (`advanceVisibleListMarker`), so a reader's labels follow every edit.
 * The `listMarker` attr is the marker the parser resolved when the document
 * was opened; an item inserted, deleted, moved or restarted since leaves it
 * stale for that item and every item after it.
 *
 * A reader counts inside the walk it already makes over the document (the AI
 * snapshot's, Markdown's), one step per paragraph, so an edit costs nothing
 * until something reads the document.
 */

import { panic } from "better-result";

import type { ResolvedParagraphNumbering } from "@stll/docx-core/model";

import {
  bulletMarkerFontName,
  convertBulletToUnicode,
} from "../docx/bulletMarkers";
import type { NumberingMap } from "../docx/numberingParser";
import type { ListRendering } from "../types/document";
import {
  advanceVisibleListMarker,
  CLEARED_LIST_RENDERING_ATTRS,
  createListCounterState,
  type ListCounterStreams,
} from "./listMarker";
import {
  listRenderingAttrPatch,
  listRenderingFromAttrs,
} from "./listRenderingAttrs";
import {
  effectiveParagraphNumberingReference,
  paragraphNumberingAttr,
  type ParagraphNumberingSources,
  paragraphNumberingSourcesFromBoundary,
} from "./numberingAttr";
import { resolveParagraphList } from "./paragraphList";
import type { ParagraphAttrs } from "./schema/nodes";

/**
 * The cached text of a `LISTNUM` field the parser folded into the marker
 * (`7.1\t(a)`). The level's template does not carry it, so it is kept as read.
 */
const foldedFieldSuffix = (attrs: Readonly<ParagraphAttrs>): string => {
  const marker = attrs.listMarker;
  const template = attrs.listMarkerTemplate;
  // Without a template the marker itself is resolved, suffix and all.
  if (!marker || !template || template.includes("\t")) {
    return "";
  }
  const tab = marker.indexOf("\t");
  return tab === -1 ? "" : marker.slice(tab);
};

const staticMarker = (attrs: Readonly<ParagraphAttrs>): string | undefined => {
  const marker = attrs.listMarker;
  if (!marker) {
    return undefined;
  }
  const text = (
    attrs.listIsBullet
      ? convertBulletToUnicode(
          marker,
          bulletMarkerFontName(attrs.listMarkerFormatting),
        )
      : marker
  ).trim();
  return text || undefined;
};

const visibleLabel = (
  counterAttrs: Readonly<ParagraphAttrs>,
  marker: string | null,
): string | undefined => {
  if (counterAttrs.listMarkerHidden === true) return undefined;
  const counted = marker?.trim();
  if (counted) {
    return `${counted}${foldedFieldSuffix(counterAttrs)}`;
  }
  // What the page paints when the counter resolves nothing: the marker as
  // the paragraph carries it.
  return staticMarker(counterAttrs);
};

const inheritedNumberingReference = (
  attrs: Readonly<ParagraphAttrs>,
  sources: ParagraphNumberingSources,
) => {
  // A tracked previous reference is only the missing numId for a current
  // level-only override. A current style reference already supplies that id
  // and must remain the owner of the current marker.
  if (
    sources.numPr?.kind !== "levelOnly" ||
    sources.numPrFromStyle !== undefined
  ) {
    return undefined;
  }
  // Only the predecessor's own `w:numPr` reference stands in, as the painter
  // (`advanceVisibleListMarker`) shows it: numbering its previous style
  // supplied left with that style.
  const previous = attrs._propertyChanges?.at(-1)?.previousFormatting;
  if (previous?.numPr?.kind !== "reference") return undefined;
  return effectiveParagraphNumberingReference(
    paragraphNumberingSourcesFromBoundary({
      numPr: previous.numPr,
      numPrFromStyle: undefined,
    }),
  );
};

/**
 * The attrs the counter reads for a paragraph of the document model, as
 * `toProseDoc` projects them (`listRenderingAttrPatch`), without the layout
 * attrs no label depends on: a reader that walks the model (Markdown) counts
 * every list paragraph through here.
 */
export const listLabelAttrsFromRendering = (
  rendering: ListRendering | undefined,
  { numPr, numPrFromStyle }: ParagraphNumberingSources,
): ParagraphAttrs => {
  const attrs: ParagraphAttrs = {
    ...(numPr === undefined ? {} : { numPr: paragraphNumberingAttr(numPr) }),
    ...(numPrFromStyle === undefined
      ? {}
      : { numPrFromStyle: paragraphNumberingAttr(numPrFromStyle) }),
  };
  if (rendering === undefined) return attrs;
  if (rendering.marker) attrs.listMarker = rendering.marker;
  if (rendering.markerTemplate)
    attrs.listMarkerTemplate = rendering.markerTemplate;
  if (rendering.isBullet) attrs.listIsBullet = true;
  if (rendering.isLegal) attrs.listIsLegal = true;
  if (rendering.numFmt) attrs.listNumFmt = rendering.numFmt;
  if (rendering.markerHidden) attrs.listMarkerHidden = true;
  if (rendering.markerFormatting)
    attrs.listMarkerFormatting = rendering.markerFormatting;
  if (rendering.levelNumFmts) attrs.listLevelNumFmts = rendering.levelNumFmts;
  if (rendering.levelStarts) attrs.listLevelStarts = rendering.levelStarts;
  if (rendering.abstractNumId !== undefined)
    attrs.listAbstractNumId = rendering.abstractNumId;
  if (rendering.startOverride !== undefined)
    attrs.listStartOverride = rendering.startOverride;
  if (rendering.implicitChildLevelAdvances !== undefined) {
    attrs.listImplicitChildLevelAdvances = rendering.implicitChildLevelAdvances;
  }
  return attrs;
};

type VisibleListLabel = {
  kind: "none" | "markerless" | "marker";
  label: string | undefined;
  /** The revision-selected reference whose marker was counted for display. */
  numbering:
    Extract<ResolvedParagraphNumbering, { kind: "reference" }> | undefined;
};

/**
 * Advances past one paragraph and answers the label it shows, or `undefined`
 * when it shows none, with the numbering that label is counted in.
 */
export type ListLabelCounter = (
  attrs: Readonly<ParagraphAttrs>,
  standaloneRendering?: ListRendering,
) => VisibleListLabel;

/**
 * A counter over paragraph attrs in document order. Call it for every
 * paragraph, numbered or not: whether a list continues depends on what stands
 * between its items.
 */
export const createListLabelCounter = (
  numberingOrStreams: NumberingMap | ListCounterStreams | null = null,
  suppliedStreams?: ListCounterStreams,
): ListLabelCounter => {
  const numbering =
    numberingOrStreams && "definitions" in numberingOrStreams
      ? numberingOrStreams
      : null;
  const readsProjectedStreams =
    numberingOrStreams !== null && !("definitions" in numberingOrStreams);
  const streams =
    suppliedStreams ??
    (numberingOrStreams && !("definitions" in numberingOrStreams)
      ? numberingOrStreams
      : {
          final: createListCounterState(),
          original: createListCounterState(),
        });
  return (attrs, standaloneRendering) => {
    const sources = paragraphNumberingSourcesFromBoundary({
      numPr: attrs.numPr,
      numPrFromStyle: attrs.numPrFromStyle,
    });
    const effective = effectiveParagraphNumberingReference(sources);
    const projectedRendering =
      standaloneRendering ??
      (attrs.listMarker != null ||
      attrs.listIsBullet === true ||
      attrs.listNumFmt != null
        ? listRenderingFromAttrs({ attrs, numId: effective?.numId ?? 1 })
        : undefined);
    const inheritedReference = inheritedNumberingReference(attrs, sources);
    const list = resolveParagraphList(
      { ...sources, ...(inheritedReference ? { inheritedReference } : {}) },
      numbering,
      projectedRendering,
    );
    const counterInput: ParagraphAttrs = { ...attrs };
    Object.assign(counterInput, CLEARED_LIST_RENDERING_ATTRS);
    const contentDerived = {
      listImplicitChildLevelAdvances: attrs.listImplicitChildLevelAdvances,
      listMarkerSecondSlotOffsetTwips: attrs.listMarkerSecondSlotOffsetTwips,
      listFoldedMarkerSuffix: attrs.listFoldedMarkerSuffix,
    };
    // When the list owner resolved a level-only override against its tracked
    // predecessor, the counter gets that same owned reference: leaving the
    // unresolved authored override would classify a marker the counter cannot
    // advance. Every other paragraph keeps its own sources, so style numbering
    // still counts as style numbering.
    const countWithInheritedReference = (
      reference: Extract<ResolvedParagraphNumbering, { kind: "reference" }>,
    ): void => {
      if (!inheritedReference) return;
      counterInput.numPr = paragraphNumberingAttr(reference);
      delete counterInput.numPrFromStyle;
    };
    if (list.kind === "marker") {
      Object.assign(counterInput, listRenderingAttrPatch(list.rendering));
      countWithInheritedReference(list.reference);
    } else if (list.kind === "markerless" && list.rendering !== undefined) {
      Object.assign(counterInput, listRenderingAttrPatch(list.rendering));
      countWithInheritedReference(list.reference);
    } else if (
      list.kind === "markerless" &&
      projectedRendering?.markerHidden === true
    ) {
      Object.assign(counterInput, listRenderingAttrPatch(projectedRendering));
    }
    Object.assign(
      counterInput,
      Object.fromEntries(
        Object.entries(contentDerived).filter(
          ([, value]) => value !== undefined,
        ),
      ),
    );
    if (counterInput._propertyChanges) {
      counterInput._propertyChanges = counterInput._propertyChanges.map(
        (change) => {
          const previous = change.previousFormatting;
          if (!previous) return change;
          const previousSources = paragraphNumberingSourcesFromBoundary({
            numPr: previous.numPr,
            numPrFromStyle: previous.numPrFromStyle,
          });
          const previousReference =
            effectiveParagraphNumberingReference(previousSources);
          const previousProjected =
            previous.listMarker != null ||
            previous.listIsBullet === true ||
            previous.listNumFmt != null
              ? listRenderingFromAttrs({
                  attrs: previous,
                  numId: previousReference?.numId ?? 1,
                })
              : undefined;
          const previousList = resolveParagraphList(
            previousSources,
            numbering,
            previousProjected,
          );
          const previousFormatting = { ...previous };
          Object.assign(previousFormatting, CLEARED_LIST_RENDERING_ATTRS);
          if (
            previousList.kind === "marker" ||
            (previousList.kind === "markerless" &&
              previousList.rendering !== undefined)
          ) {
            const rendering = previousList.rendering;
            if (rendering)
              Object.assign(
                previousFormatting,
                listRenderingAttrPatch(rendering),
              );
          } else if (previousProjected?.markerHidden === true) {
            Object.assign(
              previousFormatting,
              listRenderingAttrPatch(previousProjected),
            );
          }
          return { ...change, previousFormatting };
        },
      );
    }
    const { counterAttrs, marker } = advanceVisibleListMarker(
      counterInput,
      streams,
    );
    const projectedLabel = visibleLabel(counterAttrs, marker);
    if (readsProjectedStreams && projectedLabel !== undefined) {
      return {
        kind: "marker",
        label: projectedLabel,
        numbering: effective,
      };
    }
    const countedSources = paragraphNumberingSourcesFromBoundary({
      numPr: counterAttrs.numPr,
      numPrFromStyle: counterAttrs.numPrFromStyle,
    });
    const countedNumbering =
      effectiveParagraphNumberingReference(countedSources);
    const revisionSelectedRendering =
      counterInput._propertyChanges?.length &&
      counterAttrs !== counterInput &&
      numbering === null
        ? listRenderingFromAttrs({
            attrs: counterAttrs,
            numId: countedNumbering?.numId ?? 1,
          })
        : undefined;
    const countedInheritedReference = inheritedNumberingReference(
      counterAttrs,
      countedSources,
    );
    const countedList = resolveParagraphList(
      {
        ...countedSources,
        ...(countedInheritedReference
          ? { inheritedReference: countedInheritedReference }
          : {}),
      },
      numbering,
      revisionSelectedRendering ?? projectedRendering,
    );
    switch (countedList.kind) {
      case "none":
        return {
          kind: "none",
          label: undefined,
          numbering: undefined,
        };
      case "markerless":
        return {
          kind: "markerless",
          label: undefined,
          numbering:
            countedList.source === "authored"
              ? countedList.reference
              : undefined,
        };
      case "marker":
        return {
          kind: "marker",
          label: projectedLabel,
          numbering:
            countedList.source === "authored"
              ? countedList.reference
              : undefined,
        };
      default: {
        const unhandled: never = countedList;
        return panic(`Unhandled paragraph list ${JSON.stringify(unhandled)}`);
      }
    }
  };
};
