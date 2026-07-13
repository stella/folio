/**
 * Tolerant parser for `mutool draw -F stext -o <xml> <pdf>` output.
 *
 * The stext dialect is small and stable enough that a hand-rolled regex parser
 * is far cheaper to unit-test than wiring up a full XML DOM library for a
 * dev-only tool. We do not validate well-formedness; we just pull the
 * attributes and element bodies we need.
 *
 * Text reconstruction: a mutool `<line>` may or may not carry a `text`
 * attribute (older/newer builds differ), but its child `<char c="...">`
 * elements are always present and already in reading order, so they are the
 * one extraction path that works across mutool versions. We therefore always
 * reconstruct from `<char>` elements first; the `<line text="...">` attribute
 * (when present) is only a fallback for the degenerate case of a line with no
 * `<char>` children at all.
 */

import { normalizeLineText } from "./textNorm";
import type { LineBox, PageGeom } from "./types";

const PAGE_RE = /<page\b([^>]*)>([\s\S]*?)<\/page>/g;
const LINE_RE = /<line\b([^>]*)>([\s\S]*?)<\/line>/g;
const FONT_OPEN_RE = /<font\b([^>]*)>/;
const CHAR_RE = /<char\b([^>]*)\/>/g;

/** Extract a double-quoted attribute value by name, XML-entity-decoded.
 * `(?:^|\s)` anchors on a preceding boundary so e.g. `name="c"` never matches
 * a substring of a longer attribute name (`color`, `src`, ...). */
const attrRegexCache = new Map<string, RegExp>();

const extractAttr = (attrsSource: string, attrName: string): string | null => {
  let regex = attrRegexCache.get(attrName);
  if (regex === undefined) {
    regex = new RegExp(`(?:^|\\s)${attrName}="([^"]*)"`);
    attrRegexCache.set(attrName, regex);
  }
  const match = regex.exec(attrsSource);
  if (!match) return null;
  return decodeXmlEntities(match[1] ?? "");
};

/** Decode the handful of XML entities mutool emits. `&amp;` must be decoded
 * last so an escaped entity like `&amp;lt;` does not get double-unescaped
 * into `<`. */
const decodeXmlEntities = (raw: string): string =>
  raw
    .replace(/&#x([0-9a-fA-F]+);/g, (_, hex: string) =>
      String.fromCodePoint(Number.parseInt(hex, 16)),
    )
    .replace(/&#(\d+);/g, (_, dec: string) => String.fromCodePoint(Number.parseInt(dec, 10)))
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&");

type Bbox = { x0: number; y0: number; x1: number; y1: number };

/** Parse a `bbox="x0 y0 x1 y1"` attribute value; null when malformed. */
const parseBbox = (bboxAttr: string): Bbox | null => {
  const parts = bboxAttr.trim().split(/\s+/).map(Number);
  if (parts.length !== 4 || parts.some((n) => Number.isNaN(n))) return null;
  const [x0, y0, x1, y1] = parts;
  if (x0 === undefined || y0 === undefined || x1 === undefined || y1 === undefined) return null;
  return { x0, y0, x1, y1 };
};

type LineFont = { name?: string; sizePt?: number };

const normalizeFontEncodedText = (text: string, fontName: string | undefined): string => {
  // mutool exposes the character code behind Word's Wingdings square marker
  // as a section sign even though the rendered glyph is a black square.
  if (
    fontName !== undefined &&
    /^(?:[A-Z]{6}\+)?Wingdings(?:-|$)/iu.test(fontName) &&
    normalizeLineText(text) === "§"
  ) {
    return text.replace("§", "■");
  }
  return text;
};

/** The line's primary font: the first `<font>` element it contains. */
const extractFirstFont = (lineContent: string): LineFont => {
  const fontMatch = FONT_OPEN_RE.exec(lineContent);
  if (!fontMatch) return {};
  const fontAttrs = fontMatch[1] ?? "";
  const name = extractAttr(fontAttrs, "name");
  const sizeAttr = extractAttr(fontAttrs, "size");
  const sizePt = sizeAttr === null ? undefined : Number(sizeAttr);
  return {
    ...(name !== null ? { name } : {}),
    ...(sizePt !== undefined && !Number.isNaN(sizePt) ? { sizePt } : {}),
  };
};

type ParsedChars = { text: string; inkBbox: Bbox | null; baselinePt: number | null };

/** Reconstruct line text from `<char c="...">` elements (reading order) and
 * compute the ink bbox as the union of the NON-whitespace chars' quads. The
 * line's own `bbox` attribute includes trailing-space advances (Word emits a
 * trailing space on most lines), which the folio side never paints ink for;
 * excluding whitespace from the union keeps the two sides comparable. Falls
 * back to the line `text` attribute / line bbox when `<char>` data is absent. */
const parseChars = (lineContent: string): ParsedChars => {
  let chars = "";
  let ink: Bbox | null = null;
  const baselines: number[] = [];
  CHAR_RE.lastIndex = 0;
  let charMatch: RegExpExecArray | null = CHAR_RE.exec(lineContent);
  while (charMatch !== null) {
    const charAttrs = charMatch[1] ?? "";
    charMatch = CHAR_RE.exec(lineContent);
    const c = extractAttr(charAttrs, "c");
    if (c === null) continue;
    chars += c;
    if (c.trim() === "") continue;
    const baselineAttr = extractAttr(charAttrs, "y");
    if (baselineAttr !== null) {
      const baseline = Number(baselineAttr);
      if (Number.isFinite(baseline)) {
        baselines.push(baseline);
      }
    }
    const quad = parseQuad(extractAttr(charAttrs, "quad"));
    if (!quad) continue;
    ink =
      ink === null
        ? quad
        : {
            x0: Math.min(ink.x0, quad.x0),
            y0: Math.min(ink.y0, quad.y0),
            x1: Math.max(ink.x1, quad.x1),
            y1: Math.max(ink.y1, quad.y1),
          };
  }
  baselines.sort((a, b) => a - b);
  const middle = Math.floor(baselines.length / 2);
  let baselinePt: number | null = null;
  if (baselines.length % 2 === 0 && baselines.length > 0) {
    baselinePt = ((baselines[middle - 1] ?? 0) + (baselines[middle] ?? 0)) / 2;
  } else if (baselines.length > 0) {
    baselinePt = baselines[middle] ?? null;
  }
  return { text: chars, inkBbox: ink, baselinePt };
};

/** A `quad="x0 y0 x1 y1 x2 y2 x3 y3"` attribute (4 corners) reduced to its
 * axis-aligned bounding box; null when absent/malformed. */
const parseQuad = (quadAttr: string | null): Bbox | null => {
  if (quadAttr === null) return null;
  const parts = quadAttr.trim().split(/\s+/).map(Number);
  if (parts.length !== 8 || parts.some((n) => Number.isNaN(n))) return null;
  // Corners are (x, y) pairs: even indices are xs, odd indices are ys. The
  // length + NaN guard above already proved every entry is a real number, so
  // filtering by parity narrows to number[] without a cast.
  const xs = parts.filter((_, i) => i % 2 === 0);
  const ys = parts.filter((_, i) => i % 2 === 1);
  return {
    x0: Math.min(...xs),
    y0: Math.min(...ys),
    x1: Math.max(...xs),
    y1: Math.max(...ys),
  };
};

/** Parse all `<line>` elements within one `<page>` body into sorted LineBoxes,
 * dropping any line whose normalized text is empty. */
const parseLines = (pageContent: string): LineBox[] => {
  const lines: LineBox[] = [];
  LINE_RE.lastIndex = 0;
  let lineMatch: RegExpExecArray | null = LINE_RE.exec(pageContent);
  while (lineMatch !== null) {
    const lineAttrs = lineMatch[1] ?? "";
    const lineContent = lineMatch[2] ?? "";
    lineMatch = LINE_RE.exec(pageContent);

    const bboxAttr = extractAttr(lineAttrs, "bbox");
    const lineBbox = bboxAttr === null ? null : parseBbox(bboxAttr);

    const { text: charText, inkBbox, baselinePt } = parseChars(lineContent);
    const font = extractFirstFont(lineContent);
    const extractedText = charText !== "" ? charText : (extractAttr(lineAttrs, "text") ?? "");
    const text = normalizeFontEncodedText(extractedText, font.name);
    const bbox = inkBbox ?? lineBbox;
    if (bbox === null) continue;

    const normText = normalizeLineText(text);
    if (normText === "") continue;

    lines.push({
      text,
      normText,
      xPt: bbox.x0,
      yPt: bbox.y0,
      ...(baselinePt !== null ? { baselinePt } : {}),
      widthPt: bbox.x1 - bbox.x0,
      heightPt: bbox.y1 - bbox.y0,
      ...(font.name !== undefined ? { fontName: font.name } : {}),
      ...(font.sizePt !== undefined ? { fontSizePt: font.sizePt } : {}),
      region: "unknown",
    });
  }
  lines.sort((a, b) => a.yPt - b.yPt || a.xPt - b.xPt);
  return lines;
};

/** Parse a full `mutool draw -F stext` XML document into per-page geometry,
 * one PageGeom per `<page>` element in document order (1-based `number`). */
export const parseStextXml = (xml: string): PageGeom[] => {
  const pages: PageGeom[] = [];
  PAGE_RE.lastIndex = 0;
  let pageMatch: RegExpExecArray | null = PAGE_RE.exec(xml);
  let pageNumber = 0;
  while (pageMatch !== null) {
    pageNumber += 1;
    const pageAttrs = pageMatch[1] ?? "";
    const pageContent = pageMatch[2] ?? "";
    pageMatch = PAGE_RE.exec(xml);

    const widthAttr = extractAttr(pageAttrs, "width");
    const heightAttr = extractAttr(pageAttrs, "height");
    pages.push({
      number: pageNumber,
      widthPt: widthAttr === null ? 0 : Number(widthAttr),
      heightPt: heightAttr === null ? 0 : Number(heightAttr),
      lines: parseLines(pageContent),
    });
  }
  return pages;
};
