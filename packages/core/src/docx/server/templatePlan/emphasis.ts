/**
 * Character emphasis as a reader sees it, and the part of it a run states over
 * its paragraph's style. Only that part travels onto a template: fonts, sizes
 * and colours are the template's to decide.
 */
import type { Style } from "../../../types/document";
import type { XmlElement } from "../../xmlParser";
import { wAttr, wChild, wChildren, wChildVal } from "./xml";

/**
 * In `CT_RPr` order. Bold and italic for complex scripts are properties of
 * their own: a run may state one and not the other.
 */
const TOGGLES = [
  "b",
  "bCs",
  "i",
  "iCs",
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
  /** The text colour as RGB hex; null for automatic. */
  color: string | null;
  /** The run's shading fill as RGB hex; null for none. */
  shd: string | null;
};

/** What a run states: a toggle on or off, a value, or `null` for "cleared". */
export type EmphasisOverride = Partial<Record<Toggle, boolean>> & {
  u?: string | null;
  vertAlign?: "superscript" | "subscript" | null;
  highlight?: string | null;
  color?: string | null;
  shd?: string | null;
};

export const NO_EMPHASIS: Emphasis = Object.freeze({
  b: false,
  bCs: false,
  i: false,
  iCs: false,
  caps: false,
  smallCaps: false,
  strike: false,
  dstrike: false,
  vanish: false,
  u: null,
  vertAlign: null,
  highlight: null,
  color: null,
  shd: null,
});

type RunFormatting = Style["rPr"];

/** Toggle properties: stated on at two style tiers, they cancel (ECMA-376 §17.7.3). */
const XOR_TOGGLES: ReadonlySet<Toggle> = new Set([
  "b",
  "bCs",
  "i",
  "iCs",
  "caps",
  "smallCaps",
  "strike",
  "vanish",
]);

const formattingToggle = (
  rPr: RunFormatting,
  toggle: Toggle,
): boolean | undefined => {
  switch (toggle) {
    case "b":
      return rPr?.bold;
    case "bCs":
      return rPr?.boldCs;
    case "i":
      return rPr?.italic;
    case "iCs":
      return rPr?.italicCs;
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

const HEX_COLOR = /^[0-9A-Fa-f]{6}$/u;

/** A colour as RGB hex in capitals; null for automatic or unreadable. */
const colorOf = (value: string | null | undefined): string | null =>
  value !== undefined && value !== null && HEX_COLOR.test(value)
    ? value.toUpperCase()
    : null;

/** `base` under one more style tier's resolved run formatting. */
const withStyleTier = (base: Emphasis, rPr: RunFormatting): Emphasis => {
  if (rPr === undefined) {
    return base;
  }
  const next: Emphasis = { ...base };
  for (const toggle of TOGGLES) {
    const stated = formattingToggle(rPr, toggle);
    if (XOR_TOGGLES.has(toggle)) {
      if (stated === true) {
        next[toggle] = !base[toggle];
      }
    } else if (stated !== undefined) {
      // Double strike inherits like any other property: the last tier wins.
      next[toggle] = stated;
    }
  }
  if (rPr.underline !== undefined) {
    next.u = underlineOf(rPr.underline.style);
  }
  if (rPr.vertAlign !== undefined) {
    next.vertAlign = vertAlignOf(rPr.vertAlign);
  }
  if (rPr.highlight !== undefined) {
    next.highlight = highlightOf(rPr.highlight);
  }
  if (rPr.color !== undefined) {
    next.color = rPr.color.auto === true ? null : colorOf(rPr.color.rgb);
  }
  if (rPr.shading !== undefined) {
    next.shd = colorOf(rPr.shading.fill?.rgb);
  }
  return next;
};

/** The emphasis the document defaults give every run. */
export const defaultsBaseline = (rPr: RunFormatting): Emphasis =>
  withStyleTier(NO_EMPHASIS, rPr);

/**
 * The emphasis every run of a paragraph starts from: the document defaults
 * under the paragraph's style.
 */
export const paragraphBaseline = (
  style: Style | undefined,
  defaults: Emphasis,
): Emphasis => withStyleTier(defaults, style?.rPr);

/** `base` under a character style. */
export const withCharacterStyle = (
  base: Emphasis,
  style: Style | undefined,
): Emphasis => withStyleTier(base, style?.rPr);

const OFF_VALUES: ReadonlySet<string> = new Set(["0", "false", "off"]);

/** What a `w:rPr` states directly. */
export const readDirectEmphasis = (
  rPr: XmlElement | null | undefined,
): EmphasisOverride => {
  const override: EmphasisOverride = {};
  if (!rPr) {
    return override;
  }
  for (const toggle of TOGGLES) {
    const stated = wChildren(rPr, toggle).at(-1);
    if (stated !== undefined) {
      override[toggle] = !OFF_VALUES.has(wAttr(stated, "val") ?? "");
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
  if (wChild(rPr, "color") !== null) {
    override.color = colorOf(wChildVal(rPr, "color"));
  }
  const shading = wChild(rPr, "shd");
  if (shading !== null) {
    override.shd = colorOf(wAttr(shading, "fill"));
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
  if (effective.color !== baseline.color) {
    deviation.color = effective.color;
  }
  if (effective.shd !== baseline.shd) {
    deviation.shd = effective.shd;
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
  if (override.color !== undefined) {
    parts.push(`color=${override.color ?? "auto"}`);
  }
  if (override.shd !== undefined) {
    parts.push(`shd=${override.shd ?? "auto"}`);
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
  const toggles = parts.filter(
    (entry) => !/^(?:u|vertAlign|color|shd)=/u.test(entry),
  );
  const rest = parts.filter((entry) =>
    /^(?:u|vertAlign|color|shd)=/u.test(entry),
  );
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
  for (const name of TOGGLES) {
    const value = stated.get(name);
    if (value !== undefined) {
      out.push(onOff(prefix, name, value === "1"));
    }
  }
  // `CT_RPr` order: colour, highlight, underline, shading, vertical place.
  const color = stated.get("color");
  if (color !== undefined) {
    out.push(`<${prefix}:color ${prefix}:val="${escapeAttribute(color)}"/>`);
  }
  for (const name of ["highlight", "u"] as const) {
    const value = stated.get(name);
    if (value !== undefined) {
      out.push(
        `<${prefix}:${name} ${prefix}:val="${escapeAttribute(value)}"/>`,
      );
    }
  }
  const shd = stated.get("shd");
  if (shd !== undefined) {
    out.push(
      shd === "auto"
        ? `<${prefix}:shd ${prefix}:val="clear" ${prefix}:color="auto" ${prefix}:fill="auto"/>`
        : `<${prefix}:shd ${prefix}:val="clear" ${prefix}:color="auto" ${prefix}:fill="${escapeAttribute(shd)}"/>`,
    );
  }
  const vertAlign = stated.get("vertAlign");
  if (vertAlign !== undefined) {
    out.push(
      `<${prefix}:vertAlign ${prefix}:val="${escapeAttribute(vertAlign)}"/>`,
    );
  }
  return out.length === 0
    ? ""
    : `<${prefix}:rPr>${out.join("")}</${prefix}:rPr>`;
};
