/**
 * Feature attribution + corpus clustering: the anti-whack-a-mole layer.
 *
 * (a) Tags every paragraph of a .docx with the OOXML layout features it
 *     exercises (tables, tabs, numbering, anchored drawings, spacing rules,
 *     justification, fields, CJK, ...).
 * (b) Attributes each `Divergence` (from `compare.ts`) to the features of its
 *     containing paragraph.
 * (c) Clusters divergences across the whole corpus by (kind x feature),
 *     ranked by lift, so one root cause surfaces as one ranked cluster
 *     instead of dozens of scattered per-line diffs.
 *
 * No XML library is used: a .docx is a zip, so parts are read via the
 * system `unzip` binary (macOS ships it) and `word/document.xml` is scanned
 * with a linear regex-based tokenizer. This keeps the tool dependency-free
 * and is resilient to attribute order, which a real OOXML producer never
 * guarantees.
 */

import { normalizeLineText, textSimilarity } from "./textNorm";
import type {
  AttributedDivergence,
  Cluster,
  Divergence,
  DivergenceKind,
  DocGeom,
  FeatureAttributedResult,
  LineBox,
  ParityResult,
} from "./types";

/** Thrown for expected extraction failures (missing/unreadable zip part). */
export class FeatureExtractError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "FeatureExtractError";
  }
}

export type ParagraphFeatures = {
  normText: string;
  features: string[];
};

export type DocFeatures = {
  paragraphs: ParagraphFeatures[];
  docFeatures: string[];
};

// ---------------------------------------------------------------------------
// document.xml scanning (pure, zip-free, exported for testing)
// ---------------------------------------------------------------------------

type Token =
  | { type: "tblOpen" }
  | { type: "tblClose" }
  | { type: "pOpen"; end: number; selfClosing: boolean }
  | { type: "pClose"; index: number };

/** Matches the literal `w:tbl` / `w:p` elements only, never prefixed
 * siblings like `w:tblPr`, `w:tblGrid`, `w:pPr`, `w:pStyle`, `w:pict`: the
 * lookahead requires the char right after the local name to be whitespace,
 * `/`, or `>`. */
const TOKEN_RE = /<w:tbl(?=[\s/>])[^>]*>|<\/w:tbl>|<w:p(?=[\s/>])[^>]*>|<\/w:p>/g;

const tokenize = (xml: string): Token[] => {
  const tokenRe = new RegExp(TOKEN_RE.source, "g");
  const tokens: Token[] = [];
  let match: RegExpExecArray | null = tokenRe.exec(xml);
  while (match !== null) {
    const text = match[0];
    const index = match.index;
    if (text.startsWith("</w:tbl")) {
      tokens.push({ type: "tblClose" });
    } else if (text.startsWith("<w:tbl")) {
      tokens.push({ type: "tblOpen" });
    } else if (text.startsWith("</w:p")) {
      tokens.push({ type: "pClose", index });
    } else {
      tokens.push({ type: "pOpen", end: index + text.length, selfClosing: text.endsWith("/>") });
    }
    match = tokenRe.exec(xml);
  }
  return tokens;
};

const decodeXmlEntities = (text: string): string =>
  text
    .replace(/&#x([0-9a-fA-F]+);/g, (_, hex: string) =>
      String.fromCodePoint(Number.parseInt(hex, 16)),
    )
    .replace(/&#(\d+);/g, (_, dec: string) => String.fromCodePoint(Number.parseInt(dec, 10)))
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&");

/** Self-closing `<w:t/>` is tried before the open/close alternative so the
 * lazy `[^>]*?>` in the second branch cannot swallow the `/>` and then hunt
 * for a `</w:t>` far downstream. */
const T_RE = /<w:t(?=[\s/>])[^>]*\/>|<w:t(?=[\s/>])[^>]*?>([\s\S]*?)<\/w:t>/g;

const extractParagraphText = (segment: string): string => {
  const tRe = new RegExp(T_RE.source, "g");
  let text = "";
  let match: RegExpExecArray | null = tRe.exec(segment);
  while (match !== null) {
    if (match[1] !== undefined) text += decodeXmlEntities(match[1]);
    match = tRe.exec(segment);
  }
  return text;
};

/** `w:tab` names both a literal tab-character run child and a tab-stop
 * definition inside `<w:tabs>`. Strip tab-stop blocks before testing for the
 * run-level tab so the two features (`tab` vs `tab-stops`) never conflate. */
const stripTabsBlocks = (segment: string): string =>
  segment.replace(/<w:tabs(?=[\s/>])[^>]*>[\s\S]*?<\/w:tabs>/g, "");

const hasSpacingLineRule = (segment: string, rule: "exact" | "atLeast"): boolean => {
  const spacingRe = /<w:spacing(?=[\s/>])[^>]*>/g;
  let match: RegExpExecArray | null = spacingRe.exec(segment);
  while (match !== null) {
    if (match[0].includes(`w:lineRule="${rule}"`)) return true;
    match = spacingRe.exec(segment);
  }
  return false;
};

const hasSpacingMultiple = (segment: string): boolean => {
  const spacingRe = /<w:spacing(?=[\s/>])[^>]*>/g;
  let match: RegExpExecArray | null = spacingRe.exec(segment);
  while (match !== null) {
    if (match[0].includes('w:lineRule="auto"') && /w:line="/.test(match[0])) return true;
    match = spacingRe.exec(segment);
  }
  return false;
};

/** Broad CJK detection: CJK radicals through unified ideographs, Hangul
 * syllables, and CJK compatibility ideographs. */
const CJK_RE = /[⺀-鿿가-힣豈-﫿]/u;
/** Hebrew + Arabic (+ supplement) blocks. */
const RTL_TEXT_RE = /[֐-ࣿ]/u;

/** Segment-local feature checks (paragraph content between the outer `<w:p
 * ...>` and its `</w:p>`), excluding the depth-derived table tags and the
 * text-derived cjk/rtl tags, which are handled separately in
 * `paragraphFeatures`. */
const SEGMENT_FEATURE_CHECKS: ReadonlyArray<readonly [string, (segment: string) => boolean]> = [
  ["tab-stops", (s) => /<w:tabs(?=[\s/>])/.test(s)],
  ["tab", (s) => /<w:tab(?=[\s/>])/.test(stripTabsBlocks(s))],
  ["numbering", (s) => /<w:numPr(?=[\s/>])/.test(s)],
  ["justify", (s) => /<w:jc(?=[\s/>])[^>]*w:val="(both|distribute)"/.test(s)],
  ["align-center", (s) => /<w:jc(?=[\s/>])[^>]*w:val="center"/.test(s)],
  ["align-right", (s) => /<w:jc(?=[\s/>])[^>]*w:val="(right|end)"/.test(s)],
  ["spacing-exact", (s) => hasSpacingLineRule(s, "exact")],
  ["spacing-atLeast", (s) => hasSpacingLineRule(s, "atLeast")],
  ["spacing-multiple", (s) => hasSpacingMultiple(s)],
  ["indent", (s) => /<w:ind(?=[\s/>])/.test(s)],
  ["float-anchor", (s) => /<wp:anchor(?=[\s/>])/.test(s)],
  ["inline-image", (s) => /<wp:inline(?=[\s/>])/.test(s)],
  ["textbox", (s) => s.includes("w:txbxContent")],
  ["field", (s) => /<w:fldChar(?=[\s/>])|<w:instrText(?=[\s/>])|<w:fldSimple(?=[\s/>])/.test(s)],
  ["hyperlink", (s) => /<w:hyperlink(?=[\s/>])/.test(s)],
  ["footnote-ref", (s) => /<w:footnoteReference(?=[\s/>])/.test(s)],
  ["endnote-ref", (s) => /<w:endnoteReference(?=[\s/>])/.test(s)],
  ["page-break", (s) => /<w:br(?=[\s/>])[^>]*w:type="page"/.test(s)],
  ["column-break", (s) => /<w:br(?=[\s/>])[^>]*w:type="column"/.test(s)],
  ["keep-next", (s) => /<w:keepNext(?=[\s/>])/.test(s)],
  ["keep-lines", (s) => /<w:keepLines(?=[\s/>])/.test(s)],
  ["widow-control-off", (s) => /<w:widowControl(?=[\s/>])[^>]*w:val="(0|false)"/.test(s)],
  ["sect-props", (s) => /<w:sectPr(?=[\s/>])/.test(s)],
  ["rtl", (s) => /<w:bidi(?=[\s/>])/.test(s)],
  ["vertical-merge", (s) => /<w:vMerge(?=[\s/>])/.test(s)],
  ["small-caps", (s) => /<w:smallCaps(?=[\s/>])/.test(s)],
  ["all-caps", (s) => /<w:caps(?=[\s/>])/.test(s)],
];

const paragraphFeatures = (segment: string, tblDepth: number): ParagraphFeatures => {
  const normText = normalizeLineText(extractParagraphText(segment));
  const features = new Set<string>();
  if (tblDepth >= 1) features.add("table");
  if (tblDepth >= 2) features.add("nested-table");
  for (const [tag, check] of SEGMENT_FEATURE_CHECKS) {
    if (check(segment)) features.add(tag);
  }
  if (CJK_RE.test(normText)) features.add("cjk");
  if (RTL_TEXT_RE.test(normText)) features.add("rtl");
  return { normText, features: Array.from(features) };
};

const computeBodyFeatures = (xml: string): string[] => {
  const features = new Set<string>();
  const colsRe = /<w:cols(?=[\s/>])[^>]*>/g;
  let match: RegExpExecArray | null = colsRe.exec(xml);
  while (match !== null) {
    const numMatch = /w:num="(\d+)"/.exec(match[0]);
    if (numMatch?.[1] && Number.parseInt(numMatch[1], 10) >= 2) features.add("multi-column");
    match = colsRe.exec(xml);
  }
  if (/w:orient="landscape"/.test(xml)) features.add("landscape");
  const sectPrCount = (xml.match(/<w:sectPr(?=[\s/>])/g) ?? []).length;
  if (sectPrCount >= 2) features.add("multi-section");
  return Array.from(features);
};

export type ScannedDocument = {
  paragraphs: ParagraphFeatures[];
  /** Doc-level tags derivable from `document.xml` alone; `extractDocFeatures`
   * adds the part-existence tags (headers, footers, footnotes, endnotes,
   * embedded-fonts) that require the rest of the zip. */
  bodyFeatures: string[];
};

/** Pure paragraph scanner over `document.xml` content, exported so tests can
 * exercise the OOXML parsing without building real .docx zips. Table nesting
 * depth is tracked as `<w:tbl>` / `</w:tbl>` are encountered in document
 * order; a paragraph's depth is the depth measured at the moment its opening
 * `<w:p>` is seen. Paragraphs never nest, but a run's drawing/textbox content
 * can contain further `<w:p>` elements (`w:txbxContent`): those are folded
 * into the depth counter so the outer paragraph's segment still closes at
 * its own `</w:p>`, at the cost of also inheriting the inner paragraphs'
 * text/features (acceptable: the outer paragraph already carries the
 * "textbox" tag). */
export const scanDocumentXml = (xml: string): ScannedDocument => {
  const paragraphs: ParagraphFeatures[] = [];
  let tblDepth = 0;
  let pDepth = 0;
  let currentStart = -1;
  let currentTblDepth = 0;

  for (const token of tokenize(xml)) {
    if (token.type === "tblOpen") {
      tblDepth += 1;
      continue;
    }
    if (token.type === "tblClose") {
      tblDepth = Math.max(0, tblDepth - 1);
      continue;
    }
    if (token.type === "pOpen") {
      if (token.selfClosing) {
        if (pDepth === 0) paragraphs.push(paragraphFeatures("", tblDepth));
        continue;
      }
      if (pDepth === 0) {
        currentStart = token.end;
        currentTblDepth = tblDepth;
      }
      pDepth += 1;
      continue;
    }
    // pClose
    if (pDepth === 0) continue; // stray/unbalanced close: ignore defensively
    pDepth -= 1;
    if (pDepth === 0 && currentStart >= 0) {
      paragraphs.push(paragraphFeatures(xml.slice(currentStart, token.index), currentTblDepth));
      currentStart = -1;
    }
  }

  return { paragraphs, bodyFeatures: computeBodyFeatures(xml) };
};

// ---------------------------------------------------------------------------
// zip access (no dependency: shells out to the system `unzip`)
// ---------------------------------------------------------------------------

const runUnzip = (args: string[]): { stdout: string; exitCode: number } => {
  const proc = Bun.spawnSync(["unzip", ...args]);
  return { stdout: proc.stdout.toString("utf8"), exitCode: proc.exitCode };
};

const readZipPart = (docxPath: string, partName: string): string | undefined => {
  const { stdout, exitCode } = runUnzip(["-p", docxPath, partName]);
  return exitCode === 0 ? stdout : undefined;
};

/** Parses `unzip -l` table output into part names. Header/footer rows and
 * the trailing totals line do not match the `<size> <date> <time> <name>`
 * row shape and are skipped naturally. */
const listZipParts = (docxPath: string): string[] => {
  const { stdout, exitCode } = runUnzip(["-l", docxPath]);
  if (exitCode !== 0) return [];
  const names: string[] = [];
  for (const line of stdout.split("\n")) {
    const match = /^\s*\d+\s+\S+\s+\S+\s+(.+?)\s*$/.exec(line);
    if (match?.[1]) names.push(match[1]);
  }
  return names;
};

const hasRealNote = (xml: string): boolean =>
  Array.from(xml.matchAll(/w:id="(-?\d+)"/g)).some((m) => Number.parseInt(m[1] ?? "", 10) >= 1);

// The extraction itself is fully synchronous (Bun.spawnSync); the exported
// signature is a Promise so future extractors (or a real XML parser) can go
// async without breaking callers. Not marked `async` since the body never
// awaits (a bare `async` with no `await` trips `require-await`); the final
// value is wrapped in `Promise.resolve` instead, and a synchronous throw
// here is still caught by an `await`ing caller's try/catch.
export const extractDocFeatures = (docxPath: string): Promise<DocFeatures> => {
  const documentXml = readZipPart(docxPath, "word/document.xml");
  if (documentXml === undefined) {
    throw new FeatureExtractError(
      `cannot read word/document.xml from ${docxPath} (missing or not a valid .docx zip)`,
    );
  }

  const { paragraphs, bodyFeatures } = scanDocumentXml(documentXml);
  const docFeatures = new Set(bodyFeatures);

  const partNames = listZipParts(docxPath);
  if (partNames.some((name) => /^word\/header\d*\.xml$/.test(name))) docFeatures.add("headers");
  if (partNames.some((name) => /^word\/footer\d*\.xml$/.test(name))) docFeatures.add("footers");

  if (partNames.includes("word/footnotes.xml")) {
    const footnotesXml = readZipPart(docxPath, "word/footnotes.xml");
    if (footnotesXml && hasRealNote(footnotesXml)) docFeatures.add("footnotes");
  }
  if (partNames.includes("word/endnotes.xml")) {
    const endnotesXml = readZipPart(docxPath, "word/endnotes.xml");
    if (endnotesXml && hasRealNote(endnotesXml)) docFeatures.add("endnotes");
  }
  if (partNames.includes("word/fontTable.xml")) {
    const fontTableXml = readZipPart(docxPath, "word/fontTable.xml");
    if (fontTableXml?.includes("w:embed")) docFeatures.add("embedded-fonts");
  }

  return Promise.resolve({ paragraphs, docFeatures: Array.from(docFeatures) });
};

// ---------------------------------------------------------------------------
// ground-truth font substitution detection
// ---------------------------------------------------------------------------

/** PostScript-name decorations a PDF may append to a faithfully-rendered
 * font, normalized to lowercase alphanumerics ("ArialMT" satisfies "Arial",
 * "Calibri-BoldItalic" satisfies "Calibri"). Anything else after the
 * requested family name means Word rendered a DIFFERENT family (e.g. a
 * request for "Inter" answered by "Interstate-Bold": remainder "state-bold"
 * is no style suffix, so it counts as a substitution). */
const PDF_FONT_STYLE_SUFFIXES = new Set([
  "",
  "mt",
  "ps",
  "psmt",
  "regular",
  "roman",
  "bold",
  "italic",
  "oblique",
  "bolditalic",
  "boldoblique",
  "light",
  "medium",
  "semibold",
  "black",
  "boldmt",
  "italicmt",
  "bolditalicmt",
  "psboldmt",
  "psitalicmt",
  "psbolditalicmt",
]);

const normalizeFontName = (name: string): string =>
  name
    .replace(/^[A-Z]{6}\+/, "") // PDF subset prefix, e.g. "ABCDEF+Calibri"
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "");

const observedSatisfiesRequested = (observed: string, requested: string): boolean => {
  if (!observed.startsWith(requested)) return false;
  return PDF_FONT_STYLE_SUFFIXES.has(observed.slice(requested.length));
};

/** Pure core of `detectFontSubstitution`: which requested font families have
 * no plausible match among the PDF's observed font names? */
export const computeFontSubstitutionTags = (
  requestedFonts: string[],
  observedPdfFonts: string[],
): string[] => {
  const observed = observedPdfFonts.map(normalizeFontName).filter((name) => name.length > 0);
  const tags: string[] = [];
  const seen = new Set<string>();
  for (const requestedRaw of requestedFonts) {
    const requested = normalizeFontName(requestedRaw);
    if (requested.length === 0 || seen.has(requested)) continue;
    seen.add(requested);
    if (observed.some((obs) => observedSatisfiesRequested(obs, requested))) continue;
    tags.push(`font-substituted:${requested}`);
  }
  return tags;
};

export type FontEnvironmentAssessment = {
  status: "native" | "shared-substitution" | "mismatch" | "unverified";
  tags: string[];
  comparedLines: number;
  matchingLines: number;
};

/** Whether two renderer-reported names identify the same font family. PDF
 * PostScript names commonly carry style suffixes (`ArialMT`,
 * `Calibri-Bold`), while CSS reports the undecorated family. */
export const fontFamiliesMatch = (leftRaw: string, rightRaw: string): boolean => {
  const left = normalizeFontName(leftRaw);
  const right = normalizeFontName(rightRaw);
  if (left.length === 0 || right.length === 0) return false;
  return observedSatisfiesRequested(left, right) || observedSatisfiesRequested(right, left);
};

type FontPair = {
  wordFont: string;
  folioFont: string;
  wordWidthPt: number;
  folioWidthPt: number;
};

const MIN_FONT_METRIC_SAMPLES = 8;
// Short glyph runs are useful probes; very narrow boxes amplify extractor
// rounding, while long/justified lines mostly measure their container width.
const MIN_FONT_METRIC_WIDTH_PT = 20;
const MAX_FONT_METRIC_WIDTH_PT = 200;
const FONT_METRIC_RELATIVE_TOLERANCE = 0.05;
// A repeated directional cluster can be hidden by justified lines whose
// measured width stays stable even when their glyph metrics do not.
const MIN_FONT_METRIC_OUTLIER_SHARE = 0.25;

const collectFontPairs = (wordGeom: DocGeom, folioGeom: DocGeom): FontPair[] => {
  const folioLinesByText = new Map<string, LineBox[]>();
  for (const page of folioGeom.pages) {
    for (const line of page.lines) {
      if (line.fontName === undefined) continue;
      const lines = folioLinesByText.get(line.normText);
      if (lines) {
        lines.push(line);
      } else {
        folioLinesByText.set(line.normText, [line]);
      }
    }
  }

  for (const lines of folioLinesByText.values()) lines.reverse();

  const pairs: FontPair[] = [];
  for (const page of wordGeom.pages) {
    for (const line of page.lines) {
      if (line.fontName === undefined) continue;
      const folioLine = folioLinesByText.get(line.normText)?.pop();
      if (folioLine?.fontName !== undefined) {
        pairs.push({
          wordFont: line.fontName,
          folioFont: folioLine.fontName,
          wordWidthPt: line.widthPt,
          folioWidthPt: folioLine.widthPt,
        });
      }
    }
  }
  return pairs;
};

const hasFontMetricMismatch = (pairs: FontPair[]): boolean => {
  const ratios = pairs
    .filter(
      ({ wordFont, folioFont, wordWidthPt, folioWidthPt }) =>
        fontFamiliesMatch(wordFont, folioFont) &&
        wordWidthPt >= MIN_FONT_METRIC_WIDTH_PT &&
        folioWidthPt >= MIN_FONT_METRIC_WIDTH_PT &&
        wordWidthPt <= MAX_FONT_METRIC_WIDTH_PT &&
        folioWidthPt <= MAX_FONT_METRIC_WIDTH_PT,
    )
    .map(({ wordWidthPt, folioWidthPt }) => folioWidthPt / wordWidthPt)
    .toSorted((a, b) => a - b);
  if (ratios.length < MIN_FONT_METRIC_SAMPLES) return false;

  const middle = Math.floor(ratios.length / 2);
  const median =
    ratios.length % 2 === 0
      ? ((ratios[middle - 1] ?? 1) + (ratios[middle] ?? 1)) / 2
      : (ratios[middle] ?? 1);
  if (Math.abs(median - 1) > FONT_METRIC_RELATIVE_TOLERANCE) return true;

  const wider = ratios.filter((ratio) => ratio > 1 + FONT_METRIC_RELATIVE_TOLERANCE).length;
  const narrower = ratios.filter((ratio) => ratio < 1 - FONT_METRIC_RELATIVE_TOLERANCE).length;
  return Math.max(wider, narrower) / ratios.length >= MIN_FONT_METRIC_OUTLIER_SHARE;
};

/** Classify the actual fonts resolved by Word and Chromium. Requested-font
 * substitution is harmless for geometric parity when both renderers resolve
 * every comparable line to the same family. */
export const assessFontEnvironment = (
  requestedFonts: string[],
  wordGeom: DocGeom,
  folioGeom: DocGeom,
): FontEnvironmentAssessment => {
  const observedWordFonts = wordGeom.pages.flatMap((page) =>
    page.lines.flatMap((line) => (line.fontName === undefined ? [] : [line.fontName])),
  );
  const substitutionTags = computeFontSubstitutionTags(requestedFonts, observedWordFonts);
  const pairs = collectFontPairs(wordGeom, folioGeom);
  const matchingLines = pairs.filter(({ wordFont, folioFont }) =>
    fontFamiliesMatch(wordFont, folioFont),
  ).length;
  const metricMismatch = hasFontMetricMismatch(pairs);

  if (pairs.length === 0) {
    return {
      status: "unverified",
      tags: ["font-parity-unverified"],
      comparedLines: 0,
      matchingLines: 0,
    };
  }
  if (matchingLines !== pairs.length || metricMismatch) {
    return {
      status: "mismatch",
      tags: [
        ...(matchingLines !== pairs.length ? ["font-renderer-mismatch"] : []),
        ...(metricMismatch ? ["font-renderer-metric-mismatch"] : []),
        ...substitutionTags,
      ],
      comparedLines: pairs.length,
      matchingLines,
    };
  }
  if (substitutionTags.length > 0) {
    return {
      status: "shared-substitution",
      tags: substitutionTags.map((tag) => tag.replace("font-substituted:", "font-shared:")),
      comparedLines: pairs.length,
      matchingLines,
    };
  }
  return { status: "native", tags: [], comparedLines: pairs.length, matchingLines };
};

const RFONTS_ASCII_RE = /<w:rFonts\b[^>]*\bw:(?:ascii|hAnsi)="([^"]+)"/g;

/** Font families the document actually asks for: every `w:rFonts` ascii/hAnsi
 * reference in `document.xml`, plus the `docDefaults` fonts from `styles.xml`
 * (they cover all text that names no explicit font). Deliberately NOT the
 * whole of `styles.xml`: it routinely declares dozens of unused styles whose
 * fonts never render, which would drown the signal in false positives. */
const collectRequestedFonts = (docxPath: string): string[] => {
  const fonts: string[] = [];
  const documentXml = readZipPart(docxPath, "word/document.xml") ?? "";
  for (const match of documentXml.matchAll(RFONTS_ASCII_RE)) {
    if (match[1] !== undefined) fonts.push(match[1]);
  }
  const stylesXml = readZipPart(docxPath, "word/styles.xml") ?? "";
  const docDefaults = /<w:docDefaults>[\s\S]*?<\/w:docDefaults>/.exec(stylesXml)?.[0] ?? "";
  for (const match of docDefaults.matchAll(RFONTS_ASCII_RE)) {
    if (match[1] !== undefined) fonts.push(match[1]);
  }
  return fonts;
};

/**
 * Detects when the Word ground truth itself was rendered with substituted
 * fonts (the machine driving Word lacks a font the document requests, e.g. a
 * web font like "Inter"). Divergences in such documents may reflect font
 * availability rather than folio layout bugs, so the report must carry the
 * warning. Returned tags (e.g. "font-substituted:inter") are meant to be
 * appended to `DocFeatures.docFeatures` before attribution.
 */
export const detectFontEnvironment = (
  docxPath: string,
  wordGeom: DocGeom,
  folioGeom: DocGeom,
): FontEnvironmentAssessment =>
  assessFontEnvironment(collectRequestedFonts(docxPath), wordGeom, folioGeom);

// ---------------------------------------------------------------------------
// divergence -> paragraph attribution
// ---------------------------------------------------------------------------

/** Text used to locate a divergence's containing paragraph. `page-count` has
 * no associated text (it is a whole-document divergence). */
const divergenceSearchText = (divergence: Divergence): string | undefined => {
  if (divergence.kind === "page-count") return undefined;
  if (divergence.kind === "line-break") {
    return divergence.referenceTexts[0] ?? divergence.folioTexts[0];
  }
  if (divergence.kind === "text-mismatch") return divergence.referenceText;
  return divergence.text;
};

const MIN_MATCH_LEN = 4;
const SIMILARITY_THRESHOLD = 0.6;

/** First tries a (bidirectional, case-sensitive) substring match, requiring
 * the shorter string to carry at least `MIN_MATCH_LEN` chars so short/empty
 * strings cannot match spuriously; falls back to best-`textSimilarity`. */
const findMatchingParagraph = (
  searchText: string,
  paragraphs: ParagraphFeatures[],
): ParagraphFeatures | undefined => {
  for (const paragraph of paragraphs) {
    if (paragraph.normText.length === 0) continue;
    const shorterLen = Math.min(paragraph.normText.length, searchText.length);
    if (shorterLen < MIN_MATCH_LEN) continue;
    if (paragraph.normText.includes(searchText) || searchText.includes(paragraph.normText)) {
      return paragraph;
    }
  }

  let best: { paragraph: ParagraphFeatures; score: number } | undefined;
  for (const paragraph of paragraphs) {
    if (paragraph.normText.length === 0) continue;
    const score = textSimilarity(paragraph.normText, searchText);
    if (score >= SIMILARITY_THRESHOLD && (!best || score > best.score)) {
      best = { paragraph, score };
    }
  }
  return best?.paragraph;
};

export const attributeDivergences = (
  result: ParityResult,
  doc: DocFeatures,
): FeatureAttributedResult => {
  const docPrefixed =
    doc.docFeatures.length > 0 ? doc.docFeatures.map((f) => `doc:${f}`) : ["doc:unattributed"];

  const attributed: AttributedDivergence[] = result.divergences.map((divergence) => {
    const searchText = divergenceSearchText(divergence);
    if (searchText === undefined) return { divergence, features: docPrefixed };
    const matched = findMatchingParagraph(searchText, doc.paragraphs);
    return { divergence, features: matched ? matched.features : docPrefixed };
  });

  return { ...result, attributed, docFeatures: doc.docFeatures };
};

// ---------------------------------------------------------------------------
// corpus clustering
// ---------------------------------------------------------------------------

/** Floor on baseline prevalence so a feature that is vanishingly rare (or
 * absent) in the baseline pool cannot produce an unbounded/undefined lift. */
const MIN_BASELINE_PREVALENCE = 0.005;

const divergenceMagnitude = (divergence: Divergence): number | undefined => {
  if (divergence.kind === "x-drift" || divergence.kind === "width-drift")
    return Math.abs(divergence.deltaPt);
  if (divergence.kind === "y-drift") return Math.abs(divergence.residualPt);
  return undefined;
};

type Bucket = {
  count: number;
  docs: Set<string>;
  magnitudes: number[];
  examples: AttributedDivergence[];
};

const KEY_SEP = " ";
const bucketKey = (kind: DivergenceKind, feature: string): string => `${kind}${KEY_SEP}${feature}`;
const splitBucketKey = (key: string): { kind: DivergenceKind; feature: string } => {
  const sepIndex = key.indexOf(KEY_SEP);
  return { kind: key.slice(0, sepIndex) as DivergenceKind, feature: key.slice(sepIndex + 1) };
};

export const clusterCorpus = (
  results: FeatureAttributedResult[],
  corpusParagraphs: ParagraphFeatures[][],
): Cluster[] => {
  // Paragraph-level features (e.g. "table", "tab") are baselined against
  // every paragraph in the corpus: P(feature) = paragraphs carrying it /
  // total paragraphs. Doc-level features (the "doc:"-prefixed fallback tags)
  // have no per-paragraph meaning, so they are baselined asymmetrically
  // against `results` instead: P(feature) = docs whose `docFeatures` carries
  // it / total docs. Both pools are floored at MIN_BASELINE_PREVALENCE.
  const allParagraphs = corpusParagraphs.flat();
  const totalParagraphs = allParagraphs.length;
  const paragraphFeatureCounts = new Map<string, number>();
  for (const paragraph of allParagraphs) {
    for (const feature of paragraph.features) {
      paragraphFeatureCounts.set(feature, (paragraphFeatureCounts.get(feature) ?? 0) + 1);
    }
  }

  const totalDocs = results.length;
  const docFeatureCounts = new Map<string, number>();
  for (const result of results) {
    for (const feature of result.docFeatures) {
      docFeatureCounts.set(feature, (docFeatureCounts.get(feature) ?? 0) + 1);
    }
  }

  const baselinePrevalence = (feature: string): number => {
    if (feature.startsWith("doc:")) {
      const bare = feature.slice("doc:".length);
      const rate = totalDocs > 0 ? (docFeatureCounts.get(bare) ?? 0) / totalDocs : 0;
      return Math.max(rate, MIN_BASELINE_PREVALENCE);
    }
    const rate =
      totalParagraphs > 0 ? (paragraphFeatureCounts.get(feature) ?? 0) / totalParagraphs : 0;
    return Math.max(rate, MIN_BASELINE_PREVALENCE);
  };

  const buckets = new Map<string, Bucket>();
  const kindTotals = new Map<DivergenceKind, number>();

  for (const result of results) {
    for (const attributed of result.attributed) {
      const { kind } = attributed.divergence;
      kindTotals.set(kind, (kindTotals.get(kind) ?? 0) + 1);
      const magnitude = divergenceMagnitude(attributed.divergence);
      for (const feature of attributed.features) {
        const key = bucketKey(kind, feature);
        const bucket = buckets.get(key) ?? {
          count: 0,
          docs: new Set<string>(),
          magnitudes: [],
          examples: [],
        };
        bucket.count += 1;
        bucket.docs.add(result.file);
        if (magnitude !== undefined) bucket.magnitudes.push(magnitude);
        if (bucket.examples.length < 5) bucket.examples.push(attributed);
        buckets.set(key, bucket);
      }
    }
  }

  const minCount = results.length < 5 ? 1 : 2;
  const clusters: Cluster[] = [];
  for (const [key, bucket] of buckets) {
    if (bucket.count < minCount) continue;
    const { kind, feature } = splitBucketKey(key);
    const kindTotal = kindTotals.get(kind) ?? 0;
    const observedRate = kindTotal > 0 ? bucket.count / kindTotal : 0;
    const lift = observedRate / baselinePrevalence(feature);
    const meanMagnitudePt =
      bucket.magnitudes.length > 0
        ? bucket.magnitudes.reduce((sum, value) => sum + value, 0) / bucket.magnitudes.length
        : undefined;
    clusters.push({
      kind,
      feature,
      count: bucket.count,
      docs: Array.from(bucket.docs).sort(),
      lift,
      ...(meanMagnitudePt !== undefined ? { meanMagnitudePt } : {}),
      examples: bucket.examples,
    });
  }

  clusters.sort((a, b) => b.lift * Math.sqrt(b.count) - a.lift * Math.sqrt(a.count));
  return clusters;
};
