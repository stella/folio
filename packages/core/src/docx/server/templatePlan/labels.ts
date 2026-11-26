/**
 * The automatic list label each paragraph of a story renders, counted with
 * the same marker code the document parser uses.
 */
import { finalizeParagraphFormatting } from "@stll/docx-core";
import type {
  LevelSuffix,
  ParagraphNumberingOverride,
} from "@stll/docx-core/model";

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
  /**
   * The label and what its list puts after it, when the label can stand as
   * text: it reads the same wherever the file is opened. Null for a bullet,
   * whose glyph belongs to a font, for a number a language spells out, and
   * for an ordinal in a language whose way of writing it is not known.
   */
  asText: { label: string; suffix: LevelSuffix } | null;
};

const NO_LABEL: ParagraphLabel = {
  label: null,
  numbering: null,
  numberOnly: null,
  paragraphNumber: null,
  asText: null,
};

/** Counter formats every consumer and language writes with the same characters. */
const PLAIN_FORMATS: ReadonlySet<string> = new Set([
  "decimal",
  "decimalZero",
  "upperRoman",
  "lowerRoman",
  "upperLetter",
  "lowerLetter",
]);

const LEVEL_PLACEHOLDER = /%([1-9])/gu;

type AsTextOptions = {
  label: string;
  /** The level text the label was rendered from. */
  template: string;
  levelFormats: readonly string[] | undefined;
  suffix: LevelSuffix | undefined;
  /** Whether the paragraph's language writes `ordinal` a known way. */
  ordinalKnown: boolean;
};

const labelAsText = ({
  label,
  template,
  levelFormats,
  suffix,
  ordinalKnown,
}: AsTextOptions): ParagraphLabel["asText"] => {
  if (label.trim() === "") {
    return null;
  }
  for (const [, digit] of template.matchAll(LEVEL_PLACEHOLDER)) {
    const format = levelFormats?.[Number(digit) - 1];
    const plain =
      format !== undefined &&
      (PLAIN_FORMATS.has(format) || (format === "ordinal" && ordinalKnown));
    if (!plain) {
      return null;
    }
  }
  return { label, suffix: suffix ?? "tab" };
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

/**
 * The `ordinal` format by the language of the paragraph mark, as the
 * reference consumer prints it: Czech and German write the number with a
 * period, French `1er` and `2e`, English `1st`. A language not listed here
 * has no entry: its ordinals are not known to be any of these.
 */
const ORDINALS: Readonly<Record<string, (value: number) => string>> = {
  cs: (value) => `${String(value)}.`,
  de: (value) => `${String(value)}.`,
  fr: (value) => (value === 1 ? "1er" : `${String(value)}e`),
};

const primaryLanguage = (language: string | null): string | null =>
  language === null
    ? null
    : (language.split(/[-_]/u).at(0)?.toLowerCase() ?? null);

type RenderOptions = {
  template: string;
  counters: readonly number[];
  levelFormats: Parameters<
    typeof resolveListTemplateWithComponents
  >[0]["levelFormats"];
  forceDecimal: boolean;
  /** The paragraph mark's language. */
  language: string | null;
};

type RenderedLabel = ReturnType<typeof resolveListTemplateWithComponents> & {
  /** Whether a counter of the label is written by the language. */
  byLanguage: boolean;
  /** Whether the language's way of writing it is known. */
  known: boolean;
};

/** A level text over its counters, ordinals written by the language. */
const renderLabel = ({
  template,
  counters,
  levelFormats,
  forceDecimal,
  language,
}: RenderOptions): RenderedLabel => {
  const rendered = resolveListTemplateWithComponents({
    template,
    counters,
    levelFormats,
    forceDecimal,
  });
  const ordinals = rendered.components.filter(
    ({ level }) => !forceDecimal && levelFormats?.[level] === "ordinal",
  );
  const primary = primaryLanguage(language);
  const ordinal = primary === null ? undefined : ORDINALS[primary];
  if (ordinals.length === 0 || ordinal === undefined) {
    return {
      ...rendered,
      byLanguage: ordinals.length > 0,
      known: ordinals.length === 0 || primary === "en",
    };
  }
  let value = "";
  let cursor = 0;
  const components = rendered.components.map((component) => {
    const { level, start, end } = component;
    value += rendered.value.slice(cursor, start);
    cursor = end;
    const written = ordinals.includes(component)
      ? ordinal(counters[level] ?? 0)
      : rendered.value.slice(start, end);
    const moved = { ...component, start: value.length };
    value += written;
    return { ...moved, end: value.length };
  });
  value += rendered.value.slice(cursor);
  return { value, components, byLanguage: true, known: true };
};

/**
 * Drop the literal runs of a level text that hold a letter. Of the others
 * a `\t` field keeps the delimiters and drops the blanks: `Part %1 -` gives
 * `B-`.
 */
const numberOnlyLabel = ({ value, components }: RenderedLabel): string => {
  let out = "";
  let cursor = 0;
  const literal = (text: string): string =>
    /\p{L}/u.test(text) ? "" : text.replaceAll(/\s+/gu, "");
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
    /** The language the paragraph mark states itself, or null. */
    language?: string | null;
  }) => ParagraphLabel;
};

type LabelCounterOptions = {
  numbering: NumberingMap | null;
  styles: StyleMap;
  /** The style a paragraph without `w:pStyle` takes. */
  defaultStyleId: string | null;
  /** The language the document defaults give every run. */
  defaultLanguage?: string | null;
};

/** The language of a paragraph mark: its own, its style's, the default. */
const markLanguage = (
  stated: string | null,
  styleId: string | null,
  styles: StyleMap,
  defaultLanguage: string | null,
): string | null => {
  if (stated !== null) {
    return stated;
  }
  let cursor = styleId ?? undefined;
  const seen = new Set<string>();
  while (cursor !== undefined && !seen.has(cursor)) {
    seen.add(cursor);
    const style = styles.get(cursor);
    const language = style?.rPr?.language?.val;
    if (language !== undefined) {
      return language;
    }
    cursor = style?.basedOn;
  }
  return defaultLanguage;
};

/** A counter over one story; call `next` for every paragraph in document order. */
export const createLabelCounter = ({
  numbering,
  styles,
  defaultStyleId,
  defaultLanguage = null,
}: LabelCounterOptions): LabelCounter => {
  const counters: ComputeListMarkerOptions = {
    numbering,
    listCounters: new Map(),
    abstractCounters: new Map(),
    restartedNumIds: new Set(),
    previousList: { abstractNumId: null, fromStyle: false, numId: null },
  };
  return {
    next: ({ styleId, numPr, language = null }) => {
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
      const formatting = finalizeParagraphFormatting({
        ...(numPr === undefined ? {} : { numPr }),
        ...(inherits ? { numPrFromStyle: fromStyle } : {}),
      });
      const paragraph: Paragraph = {
        type: "paragraph",
        content: [],
        ...(formatting === undefined ? {} : { formatting }),
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
      const rendered = renderLabel({
        template,
        counters: advance.counters,
        levelFormats: listRendering.levelNumFmts,
        forceDecimal: listRendering.isLegal === true,
        language: markLanguage(
          language,
          styleId ?? defaultStyleId,
          styles,
          defaultLanguage,
        ),
      });
      // The shared marker code writes every ordinal the English way.
      const label =
        rendered.byLanguage && rendered.known
          ? rendered.value
          : listRendering.marker;
      const numberingInfo: DocxBlockNumbering | null =
        abstractNumId === null
          ? null
          : { numId: resolved.numId, ilvl: resolved.ilvl, abstractNumId };
      // A level whose run properties hide it counts and shows nothing.
      if (listRendering.markerHidden === true) {
        return { ...NO_LABEL, numbering: numberingInfo };
      }
      if (listRendering.isBullet) {
        return {
          label,
          numbering: numberingInfo,
          numberOnly: null,
          paragraphNumber: null,
          asText: null,
        };
      }
      return {
        label,
        numbering: numberingInfo,
        asText: labelAsText({
          label,
          template,
          levelFormats: listRendering.levelNumFmts,
          suffix: listRendering.markerSuffix,
          ordinalKnown: rendered.known,
        }),
        paragraphNumber: stripTrailingPeriods(label.trim()),
        numberOnly: numberOnlyLabel(rendered),
      };
    },
  };
};
