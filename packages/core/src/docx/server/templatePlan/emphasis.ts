/**
 * Character emphasis as a reader sees it, and the part of it a run states over
 * its paragraph's style. Only that part travels onto a template: fonts, sizes
 * and colours are the template's to decide.
 */
import type { Style } from "../../../types/document";
import { parseOnOffChild, type XmlElement } from "../../xmlParser";
import { wChild, wChildVal } from "./xml";

const TOGGLES = [
  "b",
  "i",
  "caps",
  "smallCaps",
  "strike",
  "dstrike",
  "vanish",
] as const;
type Toggle = (typeof TOGGLES)[number];

export type Emphasis = Record<Toggle, boolean> & {
  u: string | null;
  vertAlign: "superscript" | "subscript" | null;
  highlight: string | null;
};

/** What a run states: a toggle on or off, a value, or `null` for "cleared". */
export type EmphasisOverride = Partial<Record<Toggle, boolean>> & {
  u?: string | null;
  vertAlign?: "superscript" | "subscript" | null;
  highlight?: string | null;
};

export const NO_EMPHASIS: Emphasis = Object.freeze({
  b: false,
  i: false,
  caps: false,
  smallCaps: false,
  strike: false,
  dstrike: false,
  vanish: false,
  u: null,
  vertAlign: null,
  highlight: null,
});

const styleToggle = (
  style: Style | undefined,
  toggle: Toggle,
): boolean | undefined => {
  const rPr = style?.rPr;
  switch (toggle) {
    case "b":
      return rPr?.bold;
    case "i":
      return rPr?.italic;
    case "caps":
      return rPr?.allCaps;
    case "smallCaps":
      return rPr?.smallCaps;
    case "strike":
      return rPr?.strike;
    case "dstrike":
      return rPr?.doubleStrike;
    case "vanish":
      return rPr?.hidden;
    default:
      return toggle satisfies never;
  }
};

const underlineOf = (value: string | null | undefined): string | null =>
  value === undefined || value === null || value === "none" ? null : value;

const vertAlignOf = (
  value: string | null | undefined,
): Emphasis["vertAlign"] =>
  value === "superscript" || value === "subscript" ? value : null;

const highlightOf = (value: string | null | undefined): string | null =>
  value === undefined || value === null || value === "none" ? null : value;

/** The emphasis a paragraph style gives every run of its paragraphs. */
export const paragraphBaseline = (style: Style | undefined): Emphasis => {
  const baseline: Emphasis = { ...NO_EMPHASIS };
  for (const toggle of TOGGLES) {
    baseline[toggle] = styleToggle(style, toggle) === true;
  }
  baseline.u = underlineOf(style?.rPr?.underline?.style);
  baseline.vertAlign = vertAlignOf(style?.rPr?.vertAlign);
  baseline.highlight = highlightOf(style?.rPr?.highlight);
  return baseline;
};

/**
 * `base` under a character style. A toggle the style turns on inverts the
 * paragraph style's value (ECMA-376 §17.7.3); other properties replace it.
 */
export const withCharacterStyle = (
  base: Emphasis,
  style: Style | undefined,
): Emphasis => {
  if (style === undefined) {
    return base;
  }
  const next: Emphasis = { ...base };
  for (const toggle of TOGGLES) {
    if (styleToggle(style, toggle) === true) {
      next[toggle] = !base[toggle];
    }
  }
  const rPr = style.rPr;
  if (rPr?.underline !== undefined) {
    next.u = underlineOf(rPr.underline.style);
  }
  if (rPr?.vertAlign !== undefined) {
    next.vertAlign = vertAlignOf(rPr.vertAlign);
  }
  if (rPr?.highlight !== undefined) {
    next.highlight = highlightOf(rPr.highlight);
  }
  return next;
};

/** What a `w:rPr` states directly. */
export const readDirectEmphasis = (
  rPr: XmlElement | null | undefined,
): EmphasisOverride => {
  const override: EmphasisOverride = {};
  if (!rPr) {
    return override;
  }
  for (const toggle of TOGGLES) {
    const stated = parseOnOffChild(rPr, "w", toggle);
    if (stated !== undefined) {
      override[toggle] = stated;
    }
  }
  if (wChild(rPr, "u") !== null) {
    override.u = underlineOf(wChildVal(rPr, "u") ?? "single");
  }
  if (wChild(rPr, "vertAlign") !== null) {
    override.vertAlign = vertAlignOf(wChildVal(rPr, "vertAlign"));
  }
  if (wChild(rPr, "highlight") !== null) {
    override.highlight = highlightOf(wChildVal(rPr, "highlight"));
  }
  return override;
};

export const withOverride = (
  base: Emphasis,
  override: EmphasisOverride,
): Emphasis => ({ ...base, ...override });

/** The part of `effective` that differs from `baseline`, stated explicitly. */
export const emphasisDeviation = (
  effective: Emphasis,
  baseline: Emphasis,
): EmphasisOverride => {
  const deviation: EmphasisOverride = {};
  for (const toggle of TOGGLES) {
    if (effective[toggle] !== baseline[toggle]) {
      deviation[toggle] = effective[toggle];
    }
  }
  if (effective.u !== baseline.u) {
    deviation.u = effective.u;
  }
  if (effective.vertAlign !== baseline.vertAlign) {
    deviation.vertAlign = effective.vertAlign;
  }
  if (effective.highlight !== baseline.highlight) {
    deviation.highlight = effective.highlight;
  }
  return deviation;
};

const SIGNATURE_SEPARATOR = ";";

/** One canonical string per override; the empty string states nothing. */
export const emphasisSignature = (override: EmphasisOverride): string => {
  const parts: string[] = [];
  for (const toggle of TOGGLES) {
    const stated = override[toggle];
    if (stated !== undefined) {
      parts.push(`${toggle}=${stated ? "1" : "0"}`);
    }
  }
  if (override.highlight !== undefined) {
    parts.push(`highlight=${override.highlight ?? "none"}`);
  }
  if (override.u !== undefined) {
    parts.push(`u=${override.u ?? "none"}`);
  }
  if (override.vertAlign !== undefined) {
    parts.push(`vertAlign=${override.vertAlign ?? "baseline"}`);
  }
  return parts.join(SIGNATURE_SEPARATOR);
};

/** The highlight a signature states, or null when it states none. */
export const signatureHighlight = (signature: string): string | null => {
  const part = signature
    .split(SIGNATURE_SEPARATOR)
    .find((entry) => entry.startsWith("highlight="));
  const color = part?.slice("highlight=".length);
  return color === undefined || color === "none" ? null : color;
};

/** A signature without its highlight entry. */
export const signatureWithoutHighlight = (signature: string): string =>
  signature
    .split(SIGNATURE_SEPARATOR)
    .filter((entry) => entry !== "" && !entry.startsWith("highlight="))
    .join(SIGNATURE_SEPARATOR);

/** Whether a signature states anything besides a highlight. */
export const signatureHasEmphasis = (signature: string): boolean =>
  signature
    .split(SIGNATURE_SEPARATOR)
    .some((entry) => entry !== "" && !entry.startsWith("highlight="));

export const withSignatureHighlight = (
  signature: string,
  color: string,
): string => {
  const parts = signature
    .split(SIGNATURE_SEPARATOR)
    .filter((entry) => entry !== "" && !entry.startsWith("highlight="));
  const toggles = parts.filter((entry) => !/^(?:u|vertAlign)=/u.test(entry));
  const rest = parts.filter((entry) => /^(?:u|vertAlign)=/u.test(entry));
  return [...toggles, `highlight=${color}`, ...rest].join(SIGNATURE_SEPARATOR);
};

const onOff = (prefix: string, name: string, value: boolean): string =>
  value ? `<${prefix}:${name}/>` : `<${prefix}:${name} ${prefix}:val="0"/>`;

export type RunPropertiesXmlOptions = {
  prefix: string;
  signature: string;
  /** A character style of the template, written first. */
  rStyle?: string | null;
  escapeAttribute: (value: string) => string;
};

/**
 * The `w:rPr` for a signature, children in `CT_RPr` order; the empty string
 * when there is nothing to state.
 */
export const runPropertiesXml = ({
  prefix,
  signature,
  rStyle,
  escapeAttribute,
}: RunPropertiesXmlOptions): string => {
  const stated = new Map(
    signature
      .split(SIGNATURE_SEPARATOR)
      .filter((entry) => entry !== "")
      .map((entry) => {
        const equals = entry.indexOf("=");
        return [entry.slice(0, equals), entry.slice(equals + 1)] as const;
      }),
  );
  const out: string[] = [];
  if (rStyle) {
    out.push(`<${prefix}:rStyle ${prefix}:val="${escapeAttribute(rStyle)}"/>`);
  }
  const toggle = (name: Toggle, complexScript?: string): void => {
    const value = stated.get(name);
    if (value === undefined) {
      return;
    }
    out.push(onOff(prefix, name, value === "1"));
    if (complexScript !== undefined) {
      out.push(onOff(prefix, complexScript, value === "1"));
    }
  };
  toggle("b", "bCs");
  toggle("i", "iCs");
  toggle("caps");
  toggle("smallCaps");
  toggle("strike");
  toggle("dstrike");
  toggle("vanish");
  for (const name of ["highlight", "u", "vertAlign"] as const) {
    const value = stated.get(name);
    if (value !== undefined) {
      out.push(
        `<${prefix}:${name} ${prefix}:val="${escapeAttribute(value)}"/>`,
      );
    }
  }
  return out.length === 0
    ? ""
    : `<${prefix}:rPr>${out.join("")}</${prefix}:rPr>`;
};
