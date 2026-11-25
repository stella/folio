/**
 * The automatic list label each paragraph of a story renders, counted with
 * the same marker code the document parser uses.
 */
import type { ParagraphNumberingOverride } from "@stll/docx-core/model";

import type { Paragraph } from "../../../types/document";
import {
  computeListMarker,
  type ComputeListMarkerOptions,
} from "../../listMarkerComputation";
import { resolveListTemplateWithComponents } from "../../listTemplate";
import { computeListRendering, type NumberingMap } from "../../numberingParser";
import {
  mergeParagraphNumbering,
  paragraphNumberingReference,
  resolveParagraphNumbering,
} from "../../numberingReference";
import type { StyleMap } from "../../styleParser";
import type { DocxBlockNumbering } from "./types";

export type ParagraphLabel = {
  label: string | null;
  numbering: DocxBlockNumbering | null;
  /**
   * The label as a `REF \n \t` field shows it: number parts and delimiters,
   * without the words of the level text and without a trailing period.
   */
  numberOnly: string | null;
  /** The label as a `REF \n` field shows it: without a trailing period. */
  paragraphNumber: string | null;
};

const NO_LABEL: ParagraphLabel = {
  label: null,
  numbering: null,
  numberOnly: null,
  paragraphNumber: null,
};

/**
 * The numbering a style binds its paragraphs to. A style that names a list
 * without a level takes the level of that list which names the style back
 * (`w:pStyle`), as the consumer that wrote it does.
 */
export const styleNumbering = (
  styleId: string | null,
  styles: StyleMap,
  numbering: NumberingMap | null,
): ParagraphNumberingOverride | undefined => {
  if (styleId === null) {
    return undefined;
  }
  const stated = styles.get(styleId)?.pPr?.numPr;
  if (
    stated === undefined ||
    stated.kind !== "reference" ||
    stated.ilvl !== undefined ||
    numbering === null
  ) {
    return stated;
  }
  const abstractNumId = numbering.getAbstractNumId(stated.numId);
  const abstract =
    abstractNumId === null ? null : numbering.getAbstract(abstractNumId);
  const levelByStyle = new Map(
    (abstract?.levels ?? []).flatMap(({ pStyle, ilvl }) =>
      pStyle === undefined ? [] : [[pStyle, ilvl] as const],
    ),
  );
  let cursor: string | undefined = styleId;
  const seen = new Set<string>();
  while (cursor !== undefined && !seen.has(cursor)) {
    seen.add(cursor);
    const ilvl = levelByStyle.get(cursor);
    if (ilvl !== undefined) {
      return paragraphNumberingReference({ numId: stated.numId, ilvl });
    }
    cursor = styles.get(cursor)?.basedOn;
  }
  return stated;
};

const stripTrailingPeriods = (value: string): string =>
  value.replace(/\.+$/u, "");

type NumberOnlyOptions = {
  template: string;
  counters: readonly number[];
  levelFormats: Parameters<
    typeof resolveListTemplateWithComponents
  >[0]["levelFormats"];
  forceDecimal: boolean;
};

/** Drop the literal runs of a level text that hold a letter; keep delimiters. */
const numberOnlyLabel = ({
  template,
  counters,
  levelFormats,
  forceDecimal,
}: NumberOnlyOptions): string => {
  const { value, components } = resolveListTemplateWithComponents({
    template,
    counters,
    levelFormats,
    forceDecimal,
  });
  let out = "";
  let cursor = 0;
  const literal = (text: string): string => (/\p{L}/u.test(text) ? "" : text);
  for (const { start, end } of components) {
    out += literal(value.slice(cursor, start)) + value.slice(start, end);
    cursor = end;
  }
  out += literal(value.slice(cursor));
  return stripTrailingPeriods(out.trim());
};

export type LabelCounter = {
  next: (paragraph: {
    styleId: string | null;
    numPr: ParagraphNumberingOverride | undefined;
  }) => ParagraphLabel;
};

type LabelCounterOptions = {
  numbering: NumberingMap | null;
  styles: StyleMap;
  /** The style a paragraph without `w:pStyle` takes. */
  defaultStyleId: string | null;
};

/** A counter over one story; call `next` for every paragraph in document order. */
export const createLabelCounter = ({
  numbering,
  styles,
  defaultStyleId,
}: LabelCounterOptions): LabelCounter => {
  const counters: ComputeListMarkerOptions = {
    numbering,
    listCounters: new Map(),
    abstractCounters: new Map(),
    restartedNumIds: new Set(),
    previousList: { abstractNumId: null, fromStyle: false, numId: null },
  };
  return {
    next: ({ styleId, numPr }) => {
      const fromStyle = styleNumbering(
        styleId ?? defaultStyleId,
        styles,
        numbering,
      );
      const inherits =
        fromStyle !== undefined &&
        (numPr === undefined || numPr.kind === "levelOnly");
      const effective = inherits
        ? mergeParagraphNumbering(fromStyle, numPr)
        : numPr;
      const resolved = resolveParagraphNumbering(effective);
      const listRendering =
        numbering !== null && resolved.kind === "reference"
          ? computeListRendering(
              { numId: resolved.numId, ilvl: resolved.ilvl },
              numbering,
            )
          : null;
      const paragraph: Paragraph = {
        type: "paragraph",
        content: [],
        formatting: {
          ...(numPr === undefined ? {} : { numPr }),
          ...(inherits ? { numPrFromStyle: fromStyle } : {}),
        },
        ...(listRendering === null ? {} : { listRendering }),
      };
      const template = listRendering?.marker ?? "";
      const advance = computeListMarker(paragraph, counters);
      if (
        advance === null ||
        listRendering === null ||
        numbering === null ||
        resolved.kind !== "reference"
      ) {
        return NO_LABEL;
      }
      const abstractNumId = numbering.getAbstractNumId(resolved.numId);
      const label = listRendering.marker;
      const numberingInfo: DocxBlockNumbering | null =
        abstractNumId === null
          ? null
          : { numId: resolved.numId, ilvl: resolved.ilvl, abstractNumId };
      if (listRendering.isBullet) {
        return {
          label,
          numbering: numberingInfo,
          numberOnly: null,
          paragraphNumber: null,
        };
      }
      return {
        label,
        numbering: numberingInfo,
        paragraphNumber: stripTrailingPeriods(label.trim()),
        numberOnly: numberOnlyLabel({
          template,
          counters: advance.counters,
          levelFormats: listRendering.levelNumFmts,
          forceDecimal: listRendering.isLegal === true,
        }),
      };
    },
  };
};
