/**
 * Measure-backend parity: does the advance the font binary gives match the
 * advance the browser's canvas measures?
 *
 * Folio measures text through a swappable seam (`measureProvider.ts`). Until
 * recently one backend sat behind it, a canvas. There are now two: the canvas,
 * and `createHeadlessMeasureProvider`, which reads advances out of a parsed
 * `hmtx`. Two backends is a new place for folio to disagree with itself, and
 * the disagreement is not cosmetic: pagination can be decided on one backend
 * and the export painted with the other, so wherever the two differ a page is
 * laid out against a fiction. That is the defect class `measure-parity.spec.ts`
 * exists to catch, one layer further down.
 *
 * The arithmetic ON TOP of the advances is already shared between the backends
 * (`advanceComposition.ts`) and already proven identical by the unit test
 * beside it. What nothing proves is the raw advance of a glyph: a canvas kerns
 * and forms ligatures, a flat `hmtx` table does neither. This spec measures
 * that residue and pins it to a number.
 *
 * Both numbers come from the SAME browser in one pass, and from the same bytes:
 * each bundled `.woff` is fetched once, handed to the canvas as a `FontFace`
 * and parsed by the reader below. A machine with different system fonts moves
 * neither side, because neither side consults one.
 *
 * That is where the resemblance to `measure-parity.spec.ts` stops, and the
 * difference matters enough to state plainly. There, both numbers are the
 * browser's, so the platform's rasteriser cancels out. Here one number is the
 * browser's and the other is raw font units, so the rasteriser does NOT cancel:
 * Chromium on Linux hints advances to whole pixels, Chromium on macOS returns
 * fractional ones. Measured over this corpus, the same faces at the same sizes
 * disagree with `hmtx` by 0.046% on macOS and by up to 1.8 px PER GLYPH on
 * Linux.
 *
 * That is a finding about folio, not an artefact of this spec: folio's text
 * measurement is platform-dependent, so a document laid out by the headless
 * backend on a server and painted by a browser on Linux can break lines and
 * paginate differently from the same document on macOS. The spec therefore
 * detects which mode the platform is in, reports it, and asserts the bound that
 * mode actually admits, rather than averaging the two into one tolerance that
 * would describe neither.
 *
 * Registering the faces here rather than relying on the page's own loading is
 * what buys that property: `bundledFontSource.ts` already states it, that a
 * face the browser resolves through its own font list is a face nobody
 * measured. It also sidesteps a live defect: the playground's `@font-face`
 * rules point at the realpath of bun's global link cache, which is outside the
 * dev server's `fs.allow` root, so every bundled face 403s and the page paints
 * system fallbacks. The face descriptors still come from folio's own rules, so
 * the family names, weights and subset ranges stay folio's; only the host part
 * of the URL is repaired.
 *
 * The sfnt/WOFF reader below is inlined on purpose rather than imported from
 * `packages/core/src/fonts/sfnt`. A spec that read the font through the very
 * parser whose advances it is checking would prove nothing about that parser.
 */

import path from "node:path";

import { test, expect, type Page, type TestInfo } from "@playwright/test";

/**
 * On a platform whose canvas returns fractional advances, how far the two
 * backends may disagree on one string, as a fraction of the canvas width.
 *
 * This is not rounding slack: it bounds a real difference in what the two
 * backends can see. The canvas forms the font's `liga` substitutions, so a
 * ligated cluster is one narrower glyph on that side and its separate parts on
 * the other; `hmtx` is a flat per-glyph table with no such notion. Kerning is
 * switched off at the canvas below, so ligatures are the whole of it, plus the
 * floor the unshaped-run test pins.
 *
 * Sized from the observed worst case over this corpus: 3.0% on a run of `fi fl
 * ff ffi ffl` in Carlito 700, the one bundled family that ligates in the latin
 * subset. Set at 5% so a font revision cannot turn a green run red on its own,
 * and no wider, because that worst case is exactly the disagreement this is
 * meant to bound rather than hide. Everything the corpus measures in the other
 * four families sits at 0.05%. It shrinks to nothing if the headless backend
 * ever learns to shape.
 */
const LINEAR_TOLERANCE_RATIO = 0.05;

/**
 * Absolute floor under the relative budget. Below roughly this width a string
 * is one or two glyphs, where the ratio above does no useful work: the
 * denominator is small enough that ordinary sub-pixel rounding reads as a large
 * fraction.
 */
const LINEAR_FLOOR_PX = 0.5;

/**
 * On a platform that hints advances to whole pixels, how far the two backends
 * may disagree, per glyph.
 *
 * Per glyph rather than as a fraction because the error is a pixel phenomenon,
 * not a proportional one: it accumulates once per glyph and does not grow with
 * the size. And it is deliberately much larger than half a pixel, because full
 * hinting is not rounding. The hinter snaps stems to the pixel grid and takes
 * the advance from the hinted outline, so a narrow glyph can lose a whole pixel
 * and a bold one gain more than one: measured worst case 1.82 px per glyph
 * (`AV` in Tinos 700 at 22 pt), with `To`, `LT` and runs of `i` close behind
 * across every family.
 *
 * Set at 2.5, which is the size of the real disagreement plus a margin for a
 * rasteriser tuned differently, not a tolerance sized to make CI pass. A budget
 * this wide admits very little about shaping, which is the honest position: on
 * a quantizing platform the glyph advance the browser reports is simply not a
 * tight function of `hmtx`, and the ligature assertion below is skipped there
 * rather than weakened to fit.
 */
const QUANTIZED_BUDGET_PX_PER_GLYPH = 2.5;

/** Word's families and the bundled faces folio substitutes for them. */
const SUBSTITUTED_FAMILIES = [
  { word: "Arial", bundled: "Arimo" },
  { word: "Calibri", bundled: "Carlito" },
  { word: "Cambria", bundled: "Caladea" },
  { word: "Courier New", bundled: "Cousine" },
  { word: "Times New Roman", bundled: "Tinos" },
] as const;

/** Folio paints 400 and 700 only, as `DisplayFontFace` states. */
const WEIGHTS = [400, 700] as const;

/** Footnote, body and heading sizes, in points as a document authors them. */
const SIZES_PT = [9, 11, 22] as const;

const PT_TO_PX = 96 / 72;

/**
 * Where the bundled faces live, relative to the config root. The same files
 * `createBundledFontSource` reads off disk; the workspace symlink is used
 * rather than the realpath because the dev server serves what is under its
 * root, and the realpath is in bun's global cache.
 */
const FONTSOURCE_WORKSPACE_DIR = path.join("packages", "react", "node_modules", "@fontsource");

/**
 * Strings carrying the `liga` clusters. Kept separate because one test measures
 * them against the rest of the corpus rather than against `hmtx`.
 */
const LIGATURE_STRINGS = ["fi fl ff ffi ffl", "office affair fluffy"] as const;

/** A string with no ligature and no shaping at all, for the unshaped floor. */
const UNSHAPED_STRING = "          ";

/**
 * The corpus. Short kerning pairs are the point of it, not decoration: they are
 * where the two backends can differ most, because one adjustment is a large
 * fraction of a two-glyph string. The rest bounds the residue on the text a
 * document is actually made of.
 */
const CORPUS = [
  // Plain running text.
  "the quick brown fox jumps over the lazy dog",
  "abcdefghijklmnopqrstuvwxyz",
  "ABCDEFGHIJKLMNOPQRSTUVWXYZ",
  "Hamburgefonstiv",
  "MMMMMMMMMM",
  "iiiiiiiiii",
  // Digits and punctuation.
  "0123456789",
  "1,234.56 EUR / 2026-09-06",
  ".,;:!?()[]{}-/@#%&*+=",
  "quotation “curly” and ‘single’",
  "Article IV — Termination (30 days)",
  "§ 1811 odst. 2 písm. a)",
  // Kerning-prone pairs, each alone so a single adjustment is not diluted.
  "AV",
  "To",
  "Yo",
  "We",
  "P.",
  "F,",
  "LT",
  "av",
  "r.",
  "AVATAR To Yo We P. F, LT av r.",
  "Wa Wo Ya Yc Tw Vo",
  "Ty py ry gy jy",
  "WAVE TAV LTAV",
  // Ligature-prone: the other thing a canvas does that `hmtx` cannot.
  ...LIGATURE_STRINGS,
  // Czech and Polish diacritics, which live in a different subset face than the
  // ASCII beside them, so these also check that both backends pick the same
  // face per code point.
  "příliš žluťoučký kůň",
  "úpěl ďábelské ódy",
  "Příliš Žluťoučký Kůň",
  "zażółć gęślą jaźń",
  "Zażółć Gęślą Jaźń",
  // A long realistic sentence: the case that decides real pagination.
  "The Parties agree that this Agreement shall be governed by and construed in accordance with the laws of the Czech Republic.",
  // Whitespace only: an advance nobody shapes, which is the floor the two
  // backends cannot get below.
  UNSHAPED_STRING,
  " leading and trailing ",
] as const;

/** One string measured both ways, at one face and size. */
type AdvanceSample = {
  family: string;
  weight: number;
  sizePx: number;
  text: string;
  canvasPx: number;
  headlessPx: number;
};

/** The worst disagreement seen for one face, in each of the three units. */
type FaceResidue = {
  family: string;
  weight: number;
  samples: number;
  maxAbsPx: number;
  maxAbsSample: AdvanceSample;
  maxRatio: number;
  maxRatioSample: AdvanceSample;
  maxPerGlyphPx: number;
  maxPerGlyphSample: AdvanceSample;
};

/**
 * How the platform's canvas reports an advance: `linear` scales it from font
 * units, `quantized` hints it to a whole pixel. Which one holds decides what
 * can be asserted, so it is detected from the measurements rather than from the
 * operating system.
 */
type AdvanceMode = {
  mode: "linear" | "quantized";
  /** The sample that settled it: a provably fractional `hmtx` width. */
  witness: AdvanceSample;
};

const deltaOf = ({ canvasPx, headlessPx }: AdvanceSample): number => canvasPx - headlessPx;

const ratioOf = (sample: AdvanceSample): number =>
  sample.canvasPx <= 0 ? 0 : Math.abs(deltaOf(sample)) / sample.canvasPx;

/** Code points, which is what both backends step over one advance at a time. */
const glyphCountOf = ({ text }: AdvanceSample): number => [...text].length;

/** The unit a hinting error is actually measured in: it lands once per glyph. */
const perGlyphOf = (sample: AdvanceSample): number =>
  Math.abs(deltaOf(sample)) / glyphCountOf(sample);

/**
 * Far enough from a whole number that a canvas returning one is quantizing
 * rather than agreeing. Half of the widest linear residue this corpus has ever
 * shown would still be orders of magnitude below this.
 */
const FRACTIONAL_MARGIN = 0.2;

const isProvablyFractional = (value: number): boolean => {
  const fraction = value - Math.floor(value);
  return fraction > FRACTIONAL_MARGIN && fraction < 1 - FRACTIONAL_MARGIN;
};

/**
 * Decide the platform's advance mode from the samples themselves: take every
 * string whose `hmtx` width lands mid-pixel, and ask whether the canvas still
 * returned a whole number for it. All of them means the browser hinted every
 * advance; none of them means it scaled them. A split is a platform nobody has
 * characterised, so it fails rather than picking a branch.
 */
const detectAdvanceMode = (samples: readonly AdvanceSample[]): AdvanceMode => {
  const fractional = samples.filter((sample) => isProvablyFractional(sample.headlessPx));
  const witness = fractional.at(0);
  if (witness === undefined) {
    throw new Error("no corpus string has a mid-pixel hmtx width, so the mode cannot be read");
  }
  const whole = fractional.filter((sample) => Number.isInteger(sample.canvasPx));
  if (whole.length === fractional.length) return { mode: "quantized", witness };
  if (whole.length === 0) return { mode: "linear", witness };
  throw new Error(
    `canvas returned a whole number for ${String(whole.length)} of ${String(fractional.length)} mid-pixel widths, which is neither mode`,
  );
};

const budgetFor = (sample: AdvanceSample, { mode }: AdvanceMode): number =>
  mode === "quantized"
    ? QUANTIZED_BUDGET_PX_PER_GLYPH * glyphCountOf(sample)
    : Math.max(LINEAR_FLOOR_PX, sample.canvasPx * LINEAR_TOLERANCE_RATIO);

const describeSample = (sample: AdvanceSample): string => {
  const delta = deltaOf(sample);
  return (
    `"${sample.text}" at ${sample.sizePx.toFixed(2)}px: ` +
    `canvas ${sample.canvasPx.toFixed(3)}px, hmtx ${sample.headlessPx.toFixed(3)}px ` +
    `(${delta > 0 ? "+" : ""}${delta.toFixed(3)}px, ${(ratioOf(sample) * 100).toFixed(3)}%)`
  );
};

type SampleRequest = {
  families: readonly string[];
  weights: readonly number[];
  sizesPx: readonly number[];
  corpus: readonly string[];
  fontsourceBaseUrl: string;
};

/**
 * Measure every corpus string twice, in the page, in one pass: once with the
 * browser's canvas and once from the `hmtx` table of the same bytes the canvas
 * was handed. Everything happens inside one `page.evaluate` so no reflow, font
 * swap or resize can slip between the two readings.
 */
const collectAdvanceSamples = async (
  page: Page,
  request: SampleRequest,
): Promise<AdvanceSample[]> =>
  page.evaluate(async ({ families, weights, sizesPx, corpus, fontsourceBaseUrl }) => {
    /**
     * The faces are registered under their own family name rather than the
     * bundled one, so the page's existing `@font-face` rules for that family
     * cannot enter the matching set and decide which face the canvas shapes.
     */
    const PARITY_FAMILY_SUFFIX = "Parity";

    // WOFF 1.0 is a repackaging, not a transform: every table is stored raw or
    // zlib-deflated and is recovered by inflating it in place. `deflate` in the
    // Compression Streams API is the zlib-wrapped form, which is what WOFF uses.
    const inflate = async (bytes: Uint8Array<ArrayBuffer>): Promise<Uint8Array<ArrayBuffer>> => {
      const stream = new Blob([bytes]).stream().pipeThrough(new DecompressionStream("deflate"));
      return new Uint8Array(await new Response(stream).arrayBuffer());
    };

    const WOFF_SIGNATURE = "wOFF";
    const WOFF_NUM_TABLES = 12;
    const WOFF_DIRECTORY = 44;
    const WOFF_ENTRY_SIZE = 20;

    const tagAt = (view: DataView, offset: number): string =>
      String.fromCharCode(
        view.getUint8(offset),
        view.getUint8(offset + 1),
        view.getUint8(offset + 2),
        view.getUint8(offset + 3),
      );

    const readWoffTables = async (buffer: ArrayBuffer): Promise<Map<string, DataView>> => {
      const view = new DataView(buffer);
      if (tagAt(view, 0) !== WOFF_SIGNATURE) {
        throw new Error(`not a WOFF container: ${tagAt(view, 0)}`);
      }
      const tables = new Map<string, DataView>();
      const numTables = view.getUint16(WOFF_NUM_TABLES);
      for (let index = 0; index < numTables; index++) {
        const entry = WOFF_DIRECTORY + index * WOFF_ENTRY_SIZE;
        const offset = view.getUint32(entry + 4);
        const compLength = view.getUint32(entry + 8);
        const origLength = view.getUint32(entry + 12);
        const stored = new Uint8Array(buffer, offset, compLength);
        const data = compLength === origLength ? stored : await inflate(stored);
        tables.set(tagAt(view, entry), new DataView(data.buffer, data.byteOffset, data.byteLength));
      }
      return tables;
    };

    const tableOrThrow = (tables: Map<string, DataView>, tag: string, url: string): DataView => {
      const table = tables.get(tag);
      if (table === undefined) throw new Error(`${url} has no ${tag} table`);
      return table;
    };

    // `cmap` format 4, the segmented BMP mapping. The corpus is all BMP, so a
    // format 12 path would be code nothing here exercises.
    const buildCodePointToGlyph = (
      cmap: DataView,
      url: string,
    ): ((codePoint: number) => number) => {
      const numTables = cmap.getUint16(2);
      let subtable = -1;
      for (let index = 0; index < numTables && subtable < 0; index++) {
        const record = 4 + index * 8;
        const platform = cmap.getUint16(record);
        const encoding = cmap.getUint16(record + 2);
        const offset = cmap.getUint32(record + 4);
        const isUnicodeBmp = (platform === 3 && encoding === 1) || platform === 0;
        if (isUnicodeBmp && cmap.getUint16(offset) === 4) subtable = offset;
      }
      if (subtable < 0) throw new Error(`${url} has no format 4 cmap subtable`);

      const segCount = cmap.getUint16(subtable + 6) / 2;
      const endBase = subtable + 14;
      const startBase = endBase + segCount * 2 + 2;
      const deltaBase = startBase + segCount * 2;
      const rangeOffsetBase = deltaBase + segCount * 2;

      return (codePoint) => {
        if (codePoint > 0xffff) return 0;
        for (let segment = 0; segment < segCount; segment++) {
          if (cmap.getUint16(endBase + segment * 2) < codePoint) continue;
          const start = cmap.getUint16(startBase + segment * 2);
          if (start > codePoint) return 0;
          const delta = cmap.getInt16(deltaBase + segment * 2);
          const rangeOffset = cmap.getUint16(rangeOffsetBase + segment * 2);
          if (rangeOffset === 0) return (codePoint + delta) & 0xffff;
          const glyphAddress =
            rangeOffsetBase + segment * 2 + rangeOffset + (codePoint - start) * 2;
          const glyph = cmap.getUint16(glyphAddress);
          return glyph === 0 ? 0 : (glyph + delta) & 0xffff;
        }
        return 0;
      };
    };

    type ParsedFace = {
      url: string;
      ranges: readonly (readonly [number, number])[];
      /** Advance in em fractions, or null when the face has no such glyph. */
      advanceEm: (codePoint: number) => number | null;
    };

    const parseFace = (
      url: string,
      ranges: readonly (readonly [number, number])[],
      tables: Map<string, DataView>,
    ): ParsedFace => {
      const unitsPerEm = tableOrThrow(tables, "head", url).getUint16(18);
      const numberOfHMetrics = tableOrThrow(tables, "hhea", url).getUint16(34);
      const hmtx = tableOrThrow(tables, "hmtx", url);
      const toGlyph = buildCodePointToGlyph(tableOrThrow(tables, "cmap", url), url);

      // Glyphs past `numberOfHMetrics` share the last recorded advance; that is
      // how a monospaced tail is stored, not a missing metric.
      const advanceEm = (codePoint: number): number | null => {
        const glyph = toGlyph(codePoint);
        if (glyph === 0) return null;
        const record = Math.min(glyph, numberOfHMetrics - 1);
        return hmtx.getUint16(record * 4) / unitsPerEm;
      };

      return { url, ranges, advanceEm };
    };

    const parseUnicodeRange = (spec: string): readonly (readonly [number, number])[] =>
      spec
        .split(",")
        .flatMap<readonly [number, number]>((token) => {
          const text = token.trim().replace(/^u\+/iu, "");
          if (text.length === 0) return [];
          if (text.includes("-")) {
            const bounds = text.split("-");
            return [[parseInt(bounds.at(0) ?? "", 16), parseInt(bounds.at(1) ?? "", 16)]];
          }
          if (text.includes("?")) {
            return [
              [parseInt(text.replaceAll("?", "0"), 16), parseInt(text.replaceAll("?", "F"), 16)],
            ];
          }
          const single = parseInt(text, 16);
          return Number.isNaN(single) ? [] : [[single, single]];
        })
        .filter(([low, high]) => Number.isFinite(low) && Number.isFinite(high));

    const unquote = (value: string): string => value.trim().replace(/^["']|["']$/gu, "");

    // Folio's own `@font-face` rules are the resolution table: family, weight,
    // subset range and file name all come from them. A cross-origin sheet
    // exposes no rules, and folio's are same-document anyway.
    const fontFaceRules = (): CSSFontFaceRule[] => {
      const rules: CSSFontFaceRule[] = [];
      for (let sheetIndex = 0; sheetIndex < document.styleSheets.length; sheetIndex++) {
        const sheet = document.styleSheets[sheetIndex];
        if (sheet === undefined) continue;
        if (sheet.href !== null && !sheet.href.startsWith(location.origin)) continue;
        const list = sheet.cssRules;
        for (let ruleIndex = 0; ruleIndex < list.length; ruleIndex++) {
          const rule = list.item(ruleIndex);
          if (rule instanceof CSSFontFaceRule) rules.push(rule);
        }
      }
      return rules;
    };

    // Which code points the corpus needs, so only the subset faces that carry
    // one are fetched.
    const wanted = new Set<number>();
    for (const text of corpus) {
      for (const character of text) {
        const codePoint = character.codePointAt(0);
        if (codePoint !== undefined) wanted.add(codePoint);
      }
    }
    const wantedCodePoints = [...wanted];

    const covers = (ranges: readonly (readonly [number, number])[], codePoint: number): boolean =>
      ranges.some(([low, high]) => codePoint >= low && codePoint <= high);

    const rules = fontFaceRules();
    const samples: {
      family: string;
      weight: number;
      sizePx: number;
      text: string;
      canvasPx: number;
      headlessPx: number;
    }[] = [];

    const canvas = document.createElement("canvas");
    const context = canvas.getContext("2d");
    if (context === null) throw new Error("no 2d canvas context");

    for (const family of families) {
      for (const weight of weights) {
        const faces: ParsedFace[] = [];
        for (const rule of rules) {
          if (
            unquote(rule.style.getPropertyValue("font-family")).toLowerCase() !==
            family.toLowerCase()
          ) {
            continue;
          }
          if (rule.style.getPropertyValue("font-style").trim() !== "normal") continue;
          if (rule.style.getPropertyValue("font-weight").trim() !== String(weight)) continue;
          const source = /url\((["']?)([^"')]+\.woff)\1\)/u.exec(
            rule.style.getPropertyValue("src"),
          );
          if (source === null) continue;
          const rangeSpec = rule.style.getPropertyValue("unicode-range");
          const ranges = parseUnicodeRange(rangeSpec);
          // A face with no `unicode-range` covers everything; a subset face is
          // fetched only when the corpus reaches into its range.
          const needed =
            ranges.length === 0 || wantedCodePoints.some((codePoint) => covers(ranges, codePoint));
          if (!needed) continue;

          // SAFETY: group 2 is mandatory in a regex that matched.
          const withinPackage = /@fontsource\/(.+)$/u.exec(source[2]!);
          if (withinPackage === null) {
            throw new Error(`@font-face src is not an @fontsource file: ${source[2]!}`);
          }
          // SAFETY: group 1 is mandatory in a regex that matched.
          const url = `${fontsourceBaseUrl}/${withinPackage[1]!}`;
          const response = await fetch(url);
          if (!response.ok) throw new Error(`${url} responded ${String(response.status)}`);
          const buffer = await response.arrayBuffer();

          // Parse before registering: the dev server answers an unknown `/@fs`
          // path with the SPA fallback, so a 200 is not proof of a font, and
          // the reader's signature check is what turns that into a clear
          // failure rather than "invalid font data".
          const tables = await readWoffTables(buffer);

          // The same bytes on both sides: the canvas shapes this face, and the
          // reader below reads its `hmtx`. Anything else and the harness would
          // be measuring its own font plumbing.
          const face = new FontFace(`${family} ${PARITY_FAMILY_SUFFIX}`, buffer, {
            weight: String(weight),
            style: "normal",
            ...(rangeSpec.trim().length > 0 ? { unicodeRange: rangeSpec } : {}),
          });
          document.fonts.add(await face.load());

          faces.push(parseFace(url, ranges, tables));
        }
        if (faces.length === 0) {
          throw new Error(`no bundled .woff face for ${family} ${String(weight)}`);
        }

        // Nothing probes for a silent fallback here, because in these families
        // no probe can: they are metric-compatible substitutes, so a fallback
        // to the face they substitute (Tinos to Times, Arimo to Arial) measures
        // identical widths by design. `document.fonts.check` is no help either
        // — it reports true for a family nobody registered, since a fallback
        // can always render the text. What rules a fallback out instead is that
        // `face.load()` above rejects on bytes the browser will not accept, and
        // that this family name is unique to the face just registered, so the
        // matching set has exactly one member.
        const parityFamily = `${family} ${PARITY_FAMILY_SUFFIX}`;

        const emWidth = (text: string): number => {
          let total = 0;
          for (const character of text) {
            const codePoint = character.codePointAt(0);
            if (codePoint === undefined) continue;
            // Pick the face the browser picks: the declared `unicode-range`
            // owner first, then any face carrying the glyph. A code point no
            // face carries is a hole in the harness, so it throws rather than
            // measuring `.notdef` as if it were a width.
            const owner =
              faces.find(
                (face) => covers(face.ranges, codePoint) && face.advanceEm(codePoint) !== null,
              ) ?? faces.find((face) => face.advanceEm(codePoint) !== null);
            if (owner === undefined) {
              const hex = codePoint.toString(16).toUpperCase().padStart(4, "0");
              throw new Error(`${family} ${String(weight)} has no glyph for U+${hex}`);
            }
            total += owner.advanceEm(codePoint) ?? 0;
          }
          return total;
        };

        for (const sizePx of sizesPx) {
          // Kerning off on purpose: the comparison is about which metric SOURCE
          // is read, and leaving kerning on would mostly measure the browser's
          // pair table. Ligatures stay on, because canvas offers no switch for
          // them and folio's painted DOM text has them on too.
          context.font = `${String(weight)} ${String(sizePx)}px "${parityFamily}"`;
          context.fontKerning = "none";
          for (const text of corpus) {
            samples.push({
              family,
              weight,
              sizePx,
              text,
              canvasPx: context.measureText(text).width,
              headlessPx: emWidth(text) * sizePx,
            });
          }
        }
      }
    }
    return samples;
  }, request);

/** The sample with the highest score. Callers pass a non-empty group. */
const worstBy = (
  group: readonly AdvanceSample[],
  score: (sample: AdvanceSample) => number,
): AdvanceSample => {
  // SAFETY: a group exists only because a sample was pushed into it.
  let worst = group[0]!;
  for (const sample of group) {
    if (score(sample) > score(worst)) worst = sample;
  }
  return worst;
};

const summarise = (samples: readonly AdvanceSample[]): FaceResidue[] => {
  const byFace = new Map<string, AdvanceSample[]>();
  for (const sample of samples) {
    const key = `${sample.family} ${String(sample.weight)}`;
    const bucket = byFace.get(key);
    if (bucket === undefined) byFace.set(key, [sample]);
    else bucket.push(sample);
  }
  return [...byFace.values()].map((bucket) => {
    const worstAbs = worstBy(bucket, (sample) => Math.abs(deltaOf(sample)));
    const worstRatio = worstBy(bucket, ratioOf);
    const worstPerGlyph = worstBy(bucket, perGlyphOf);
    return {
      family: worstAbs.family,
      weight: worstAbs.weight,
      samples: bucket.length,
      maxAbsPx: Math.abs(deltaOf(worstAbs)),
      maxAbsSample: worstAbs,
      maxRatio: ratioOf(worstRatio),
      maxRatioSample: worstRatio,
      maxPerGlyphPx: perGlyphOf(worstPerGlyph),
      maxPerGlyphSample: worstPerGlyph,
    };
  });
};

const wordFamilyFor = (bundled: string): string =>
  SUBSTITUTED_FAMILIES.find((entry) => entry.bundled === bundled)?.word ?? "?";

const describeMode = ({ mode, witness }: AdvanceMode): string =>
  `advance mode: ${mode} — the canvas returned ${mode === "quantized" ? "a whole number" : "a fractional width"} ` +
  `for every mid-pixel hmtx width, e.g. ${witness.family} ${String(witness.weight)} ${describeSample(witness)}`;

const reportLines = (residues: readonly FaceResidue[], advanceMode: AdvanceMode): string[] => [
  describeMode(advanceMode),
  ...residues.map(
    (residue) =>
      `${residue.family} (for ${wordFamilyFor(residue.family)}) ${String(residue.weight)}, ` +
      `${String(residue.samples)} strings: ` +
      `max |delta| ${residue.maxAbsPx.toFixed(3)}px on ${describeSample(residue.maxAbsSample)} | ` +
      `max relative ${(residue.maxRatio * 100).toFixed(3)}% on ${describeSample(residue.maxRatioSample)} | ` +
      `max per glyph ${residue.maxPerGlyphPx.toFixed(3)}px on "${residue.maxPerGlyphSample.text}"`,
  ),
];

/** Load a document into the playground and wait for the first page to paint. */
const openPlayground = async (page: Page): Promise<void> => {
  await page.goto("/?file=docx-editor-demo.docx");
  await page.waitForSelector(".layout-page .layout-line", { timeout: 30_000 });
  // Web fonts change advances, so a comparison taken mid-swap is meaningless.
  await page.evaluate(async () => {
    await document.fonts.ready;
  });
};

const attachReport = async (testInfo: TestInfo, lines: readonly string[]): Promise<void> => {
  for (const line of lines) console.log(line);
  await testInfo.attach("measure-backend-parity.txt", {
    body: lines.join("\n"),
    contentType: "text/plain",
  });
};

/**
 * The dev server serves any file under the repository through `/@fs`. The
 * repository is located from the config file rather than from `rootDir`, which
 * Playwright sets to the test directory, or from the cwd, which belongs to
 * whoever invoked the run.
 */
const fontsourceBaseUrlFor = (testInfo: TestInfo): string => {
  const { configFile } = testInfo.config;
  if (configFile === undefined) throw new Error("no playwright config file to locate the repo");
  return `/@fs${path.join(path.dirname(configFile), FONTSOURCE_WORKSPACE_DIR)}`;
};

const BUNDLED_FAMILIES = SUBSTITUTED_FAMILIES.map((entry) => entry.bundled);
const SIZES_PX = SIZES_PT.map((pt) => pt * PT_TO_PX);

test.describe("measure-backend parity", () => {
  test("canvas advances and hmtx advances agree within the platform's budget", async ({
    page,
  }, testInfo) => {
    await openPlayground(page);

    const samples = await collectAdvanceSamples(page, {
      families: BUNDLED_FAMILIES,
      weights: [...WEIGHTS],
      sizesPx: SIZES_PX,
      corpus: [...CORPUS],
      fontsourceBaseUrl: fontsourceBaseUrlFor(testInfo),
    });

    // Exact, not `> 0`: if face resolution ever silently drops a family or a
    // size, this spec must fail rather than compare whatever is left.
    expect(samples.length).toBe(
      SUBSTITUTED_FAMILIES.length * WEIGHTS.length * SIZES_PT.length * CORPUS.length,
    );

    const advanceMode = detectAdvanceMode(samples);
    await attachReport(testInfo, reportLines(summarise(samples), advanceMode));

    const offenders = samples.filter(
      (sample) => Math.abs(deltaOf(sample)) > budgetFor(sample, advanceMode),
    );
    expect(
      offenders.map(
        (sample) => `${sample.family} ${String(sample.weight)} ${describeSample(sample)}`,
      ),
    ).toEqual([]);
  });

  test("an unshaped run isolates the residue that is not shaping", async ({ page }, testInfo) => {
    // A run of spaces forms no ligature and takes no pair adjustment, so the
    // whole shaping argument drops out and what is left is the floor neither
    // backend can get below. Separating that from the ligature residue is what
    // makes the budget above readable: otherwise a growing floor would hide
    // under a tolerance sized for shaping.
    //
    // A space is also the one glyph with no outline, which is why this floor is
    // tight even where the platform hints: there is no stem for the hinter to
    // snap, so a quantizing browser only ROUNDS the advance, and the error is
    // bounded by half a pixel per glyph in a way a lettered string is not.
    await openPlayground(page);

    const samples = await collectAdvanceSamples(page, {
      families: BUNDLED_FAMILIES,
      weights: [...WEIGHTS],
      sizesPx: SIZES_PX,
      corpus: [UNSHAPED_STRING],
      fontsourceBaseUrl: fontsourceBaseUrlFor(testInfo),
    });

    /**
     * The unshaped floor where the canvas scales advances: a fraction of the
     * run's width. Observed at 0.05%, flat across families, weights and sizes.
     * Set four times wider because the residue is the platform's, not folio's,
     * and still two orders of magnitude below the ligature residue, so this
     * stays an assertion about arithmetic rather than about shaping.
     */
    const UNSHAPED_LINEAR_TOLERANCE_RATIO = 0.002;

    /**
     * The same floor where the canvas hints advances. Half a pixel per glyph is
     * the structural bound on rounding an advance to the pixel grid, and a
     * space has no outline to be hinted beyond that; the measured worst case is
     * 0.453 px per glyph, just inside it. Set at 0.75 only to leave room for a
     * rasteriser that rounds away from zero exactly at the boundary, which is
     * why it stays far below the 2.5 px a lettered glyph is allowed.
     */
    const UNSHAPED_QUANTIZED_BUDGET_PX_PER_GLYPH = 0.75;

    const advanceMode = detectAdvanceMode(samples);
    const lines = [
      describeMode(advanceMode),
      ...samples.map(
        (sample) =>
          `${sample.family} ${String(sample.weight)} ${describeSample(sample)}, ` +
          `${perGlyphOf(sample).toFixed(3)}px per glyph`,
      ),
    ];
    await attachReport(testInfo, lines);

    const overFloor = (sample: AdvanceSample): boolean =>
      advanceMode.mode === "quantized"
        ? perGlyphOf(sample) > UNSHAPED_QUANTIZED_BUDGET_PX_PER_GLYPH
        : ratioOf(sample) > UNSHAPED_LINEAR_TOLERANCE_RATIO;

    expect(
      samples
        .filter(overFloor)
        .map((sample) => `${sample.family} ${String(sample.weight)} ${describeSample(sample)}`),
    ).toEqual([]);
  });

  test("ligature substitution is the whole of the shaping residue", async ({ page }, testInfo) => {
    // Which bundled families ligate is a fact about the font files folio ships,
    // and it decides where the headless backend mismeasures: `hmtx` cannot see
    // a `liga` substitution, so a family that ligates is a family whose width
    // the headless backend overstates. Pinning the exact set makes a font
    // revision that starts ligating somewhere new show up as a failure rather
    // than as silently worse pagination.
    await openPlayground(page);

    const samples = await collectAdvanceSamples(page, {
      families: BUNDLED_FAMILIES,
      weights: [...WEIGHTS],
      sizesPx: SIZES_PX,
      corpus: [...LIGATURE_STRINGS],
      fontsourceBaseUrl: fontsourceBaseUrlFor(testInfo),
    });

    /**
     * Where a shaping difference starts, as a fraction of the width. The
     * families that ligate sit at 2.5% and above; the ones that do not sit at
     * 0.05%, the linear floor. Anywhere in the two orders of magnitude between
     * separates them, so this is placed an order of magnitude above the floor
     * rather than just below the signal.
     */
    const SHAPING_SIGNAL_RATIO = 0.005;

    const advanceMode = detectAdvanceMode(samples);
    await attachReport(testInfo, reportLines(summarise(samples), advanceMode));

    // A hinted advance moves a lettered glyph by up to 2.5 px, and a ligature
    // saves about 0.6 px per cluster, so on a quantizing platform the signal is
    // under the noise and no honest threshold exists. Reported above either
    // way; asserted only where the canvas scales advances.
    test.skip(
      advanceMode.mode === "quantized",
      "the platform hints advances to whole pixels, which is larger than the ligature residue it would have to resolve",
    );

    const ligating = summarise(samples)
      .filter((residue) => residue.maxRatio > SHAPING_SIGNAL_RATIO)
      .map((residue) => `${residue.family} ${String(residue.weight)}`);
    expect(ligating).toEqual(["Carlito 400", "Carlito 700"]);
  });
});
