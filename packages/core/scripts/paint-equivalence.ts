/**
 * Paint one display list with both backends and measure the difference.
 *
 *   bun packages/core/scripts/paint-equivalence.ts [fixture.docx | directory]
 *       [--json] [--out <dir>] [--threshold <0..1>] [--update-baseline]
 *
 * The point of the harness is the *single* display list: layout runs once, the
 * builder runs once, and the DOM backend and the PDF backend are handed the
 * same structure. A raster difference is therefore a backend difference, never
 * a layout difference, which is what makes a number here attributable.
 *
 * ## What the score does not say
 *
 * A page can score 1.0 because neither backend painted something. The display
 * list's `unsupported` entries and every font substitution, from the measurer
 * and from the PDF writer alike, are reported beside the score for exactly
 * that reason: they bound what the score is evidence of.
 *
 * ## Why the absolute score is ~0.97 and not ~0.999
 *
 * The absolute number is bounded by three known residuals between the two
 * backends, not by how wrong either one is. Read a drop, not a level.
 *
 * 1. **Advance drift, and it dominates.** The harness builds the display list
 *    with the *headless* measure provider, whose advances come from `hmtx`
 *    with no kerning and no ligatures, and then renders that list in a
 *    *browser*, which advances glyphs on its own shaped metrics. Measured over
 *    `podily-bps.docx`: of 1995 runs, the width Chrome paints differs from the
 *    run's declared `advancesPx` by more than 2 px in 53.7% of them (median
 *    3.0 px, p90 17.8 px). Each line starts flush and separates left to right,
 *    which is what the diff PNGs show. **This is an artifact of how the
 *    harness is wired, not of the product path**: in the editor the same list
 *    is built by the canvas provider inside the same browser that paints it,
 *    so the two agree there. Nobody should read podily's 0.94 as the editor
 *    being 6% wrong.
 * 2. **Baseline placement: ~0.22 px, and no longer the backend's doing.**
 *    Measured on `sample.docx` page 1 (Carlito bold at 18.667 px): the DOM
 *    backend asks for `top: 107.399px` and Chrome lays the span out at
 *    `107.391px`, so CSS does exactly what it was told. The residual is that
 *    Chrome rounds the font box to whole pixels (`fontBoundingBoxAscent` 18 px
 *    against the provider's 17.773 px), putting the painted baseline 0.22 px
 *    below the display list's. Same headless-versus-browser seam as (1), a
 *    quarter of a pixel wide.
 * 3. **Stroke alignment, ~1 px on hairlines.** A 1 px stroke centred on an
 *    integer coordinate covers two half-rows in the PDF rasterizer and one
 *    whole row in CSS: a table border sampled at its top edge reads `25,43,59`
 *    from `mutool` and `0,0,0` from Chrome.
 *
 * Two further differences are *shared* rather than divergent, so they do not
 * move the score: glyphs outside the loaded font subset paint the same
 * `.notdef` box in both arms, and both arms rasterize with different engines
 * (MuPDF and Skia), whose antialiasing spreads ink over slightly different
 * rows.
 *
 * ## What makes the run fail
 *
 * The gate is the committed baseline, not an absolute score. With no
 * `--threshold`, the run fails only when a fixture scores more than
 * {@link BASELINE_REGRESSION_TOLERANCE} below its baseline entry, or when a
 * default-corpus fixture the baseline names did not run at all. A fixture with
 * no baseline entry is reported as new and does not fail. `--threshold <n>`
 * adds an absolute floor on top of that check, for a caller who wants one. A
 * script whose plain invocation always exits 1 teaches everyone to ignore its
 * exit code, at which point it has stopped being a signal.
 *
 * ## Optional arms
 *
 * `mutool` rasterizes the PDF and Playwright's chromium rasterizes the DOM.
 * Either being absent is reported and exits 0, the way `parity/` treats a
 * reference renderer it cannot find: a missing tool is not a regression.
 */

import { existsSync } from "node:fs";
import { mkdir, readdir, readFile, rm, stat, writeFile } from "node:fs/promises";
import path from "node:path";
import { pathToFileURL } from "node:url";

import { chromium } from "@playwright/test";
import { Result } from "better-result";

import { comparePageRasters } from "../../../parity/rasterCompare";
import type { RasterPageComparison } from "../../../parity/types";
import { buildDisplayList } from "../src/display-list/build/buildDisplayList";
import { renderDisplayListToDom } from "../src/display-list/dom/renderDisplayListToDom";
import type { DisplayFontFace, DisplayPage, DisplayUnsupported } from "../src/display-list/types";
import type { HeadlessFontSubstitution } from "../src/fonts/headlessMeasure";
import { installHeadlessMeasureProvider } from "../src/fonts/headlessMeasure";
import type { HeadlessLayoutError, HeadlessLayoutGap } from "../src/headless-layout";
import { layoutDocxHeadless } from "../src/headless-layout";
import type { PdfSubstitution, WritePdfError } from "../src/pdf/writePdf";
import { writePdf } from "../src/pdf/writePdf";
import { bundledFontFaceCss, createBundledFontSource } from "./bundledFontSource";

const USAGE = [
  "usage: bun packages/core/scripts/paint-equivalence.ts [fixture.docx | directory] " +
    "[--json] [--out <dir>] [--threshold <0..1>] [--update-baseline]",
  "",
  "Gates on the committed baseline, not on an absolute score: exit 1 when a fixture",
  "scores more than 0.005 below its baseline entry, or when a default-corpus fixture",
  "the baseline names did not run. A fixture with no baseline entry is reported as new",
  "and does not fail. The absolute score is bounded by the two backends' known",
  "residuals (see the module header) rather than by correctness, so read a drop, not a",
  "level.",
  "",
  "  --threshold <0..1>  add an absolute floor on top of the regression check",
  "  --update-baseline   rewrite paint-equivalence.baseline.json from this run",
].join("\n");

/** Fixed so two runs over one fixture produce byte-identical PDFs. */
const FIXED_TIMESTAMP = "1970-01-01T00:00:00.000Z";
const PRODUCER = "folio paint-equivalence";

/** How far a fixture may drop below its baseline before the run fails. */
const BASELINE_REGRESSION_TOLERANCE = 0.005;
const BASELINE_DECIMALS = 4;
const BOLD_WEIGHT = 700;

/** Both arms rasterize at the display list's own unit: CSS pixels at 96 dpi. */
const RASTER_DPI = 96;

const REPO_ROOT = path.resolve(import.meta.dir, "..", "..", "..");
const DEFAULT_OUTPUT_DIR = path.join(REPO_ROOT, ".cache", "paint-equivalence");
const BASELINE_PATH = path.join(import.meta.dir, "paint-equivalence.baseline.json");

/**
 * The default corpus: one small, one medium and one large document, so a bare
 * run is fast enough to be habitual. A directory argument runs everything in
 * it.
 */
const DEFAULT_FIXTURES = [
  "tests/visual/fixtures/docx-editor-demo.docx",
  "tests/visual/fixtures/sample.docx",
  "tests/visual/fixtures/podily-bps.docx",
] as const;

const VIEWPORT = { width: 1400, height: 1200 };
const CHROMIUM_MISSING_MARKER = "Executable doesn't exist";
const CHROMIUM_MISSING_MESSAGE =
  "Playwright chromium missing; run: bunx playwright install chromium";
const PAGE_SELECTOR = ".layout-page";
/**
 * Playwright's 30 s default is not enough for a long document on a loaded
 * machine: the capture waits for the element to be stable, and a 36-page
 * fixture sharing a host with other builds passes that mark late.
 */
const SCREENSHOT_TIMEOUT_MS = 120_000;
const PAGE_PNG_RE = /^p(\d+)\.png$/;
const DOCX_EXTENSION = ".docx";

// ---------------------------------------------------------------------------
// Arguments
// ---------------------------------------------------------------------------

type ParsedArgs = {
  target: string | null;
  json: boolean;
  outputDir: string;
  /** Null when the caller named none: the baseline is then the only gate. */
  threshold: number | null;
  updateBaseline: boolean;
};

const parseArgs = (argv: readonly string[]): ParsedArgs | null => {
  const positional: string[] = [];
  let json = false;
  let outputDir = DEFAULT_OUTPUT_DIR;
  let threshold: number | null = null;
  let updateBaseline = false;

  for (let index = 0; index < argv.length; index++) {
    const arg = argv[index];
    if (arg === "--json") {
      json = true;
      continue;
    }
    if (arg === "--update-baseline") {
      updateBaseline = true;
      continue;
    }
    if (arg === "--out" || arg === "--threshold") {
      const value = argv[++index];
      if (value === undefined) {
        return null;
      }
      if (arg === "--out") {
        outputDir = path.resolve(value);
        continue;
      }
      threshold = Number(value);
      if (!Number.isFinite(threshold) || threshold < 0 || threshold > 1) {
        return null;
      }
      continue;
    }
    if (arg === undefined || arg.startsWith("--")) {
      return null;
    }
    positional.push(arg);
  }

  if (positional.length > 1) {
    return null;
  }
  return { target: positional.at(0) ?? null, json, outputDir, threshold, updateBaseline };
};

// ---------------------------------------------------------------------------
// Error printing, as `compare.ts` prints it
// ---------------------------------------------------------------------------

/** Guard against a chain that loops or is pathologically deep. */
const MAX_CAUSE_DEPTH = 8;

const describeError = (value: unknown): string => {
  if (value instanceof Error) {
    return `${value.name}: ${value.message}`;
  }
  return typeof value === "string" ? value : JSON.stringify(value);
};

const readCause = (value: unknown): unknown =>
  typeof value === "object" && value !== null && "cause" in value ? value.cause : undefined;

const causeChain = (error: unknown): string[] => {
  const lines: string[] = [];
  const seen = new Set<unknown>();
  let current: unknown = error;
  while (current !== undefined && current !== null && !seen.has(current)) {
    if (lines.length >= MAX_CAUSE_DEPTH) {
      lines.push("... further causes omitted");
      break;
    }
    seen.add(current);
    lines.push(describeError(current));
    current = readCause(current);
  }
  return lines;
};

const printFailure = (fixture: string, error: unknown): void => {
  console.error(`${fixture} failed`);
  const [head, ...causes] = causeChain(error);
  console.error(`  ${head ?? describeError(error)}`);
  for (const [index, line] of causes.entries()) {
    console.error(`${"  ".repeat(index + 2)}caused by: ${line}`);
  }
};

// ---------------------------------------------------------------------------
// A server-side document, just wide enough for the DOM backend
// ---------------------------------------------------------------------------

type StubElement = {
  readonly tagName: string;
  readonly style: Record<string, string>;
  readonly dataset: Record<string, string>;
  readonly attributes: Record<string, string>;
  readonly children: StubElement[];
  className: string;
  id: string;
  textContent: string;
  alt: string;
  src: string;
  href: string;
  title: string;
  readonly append: (...nodes: StubElement[]) => void;
  readonly appendChild: (node: StubElement) => StubElement;
  readonly setAttribute: (name: string, value: string) => void;
};

const VOID_TAGS = new Set(["img", "br", "hr"]);

const createStubElement = (tagName: string): StubElement => {
  const children: StubElement[] = [];
  const attributes: Record<string, string> = {};
  return {
    tagName,
    style: {},
    dataset: {},
    attributes,
    children,
    className: "",
    id: "",
    textContent: "",
    alt: "",
    src: "",
    href: "",
    title: "",
    append: (...nodes) => {
      children.push(...nodes);
    },
    appendChild: (node) => {
      children.push(node);
      return node;
    },
    setAttribute: (name, value) => {
      attributes[name] = value;
    },
  };
};

/**
 * A `Document` with only what the DOM backend reaches for, so the backend can
 * run with no bundler, no playground and no browser.
 */
const stubDocument = (): Document => {
  const doc = { createElement: createStubElement };
  // SAFETY: the backend calls `createElement` and then touches `style`,
  // `dataset`, `className`, `id`, `textContent`, `alt`, `src`, `href`, `title`
  // and `append`, every one of which `createStubElement` provides. A backend
  // that grows a new requirement throws on the missing property here rather
  // than silently serializing less than it painted.
  return doc as unknown as Document;
};

/** The stub document created it, so it is a {@link StubElement} wearing the
 * backend's declared return type. */
// SAFETY: only elements from `stubDocument()` reach this.
const asStub = (element: HTMLElement) => element as unknown as StubElement;

/** A vendor property is camelCase with no leading capital, so the general
 * hyphenation rule cannot infer its leading dash. */
const VENDOR_PREFIXES = ["webkit-", "moz-", "ms-", "o-"] as const;

const toKebabCase = (name: string): string => {
  const hyphenated = name.replaceAll(/[A-Z]/gu, (letter) => `-${letter.toLowerCase()}`);
  return VENDOR_PREFIXES.some((prefix) => hyphenated.startsWith(prefix))
    ? `-${hyphenated}`
    : hyphenated;
};

const escapeText = (value: string): string =>
  value.replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;");

const escapeAttribute = (value: string): string => escapeText(value).replaceAll('"', "&quot;");

const serializeStyle = (style: Record<string, string>): string =>
  Object.entries(style)
    .map(([property, value]) => `${toKebabCase(property)}: ${value}`)
    .join("; ");

/**
 * Serialized without one character of added whitespace: the backend sets
 * `white-space: pre` on every run, so an indentation newline would be painted.
 */
const RAW_TEXT_TAGS = new Set(["style", "script"]);

const serializeElement = (element: StubElement): string => {
  const attributes: string[] = [];
  if (element.className !== "") attributes.push(`class="${escapeAttribute(element.className)}"`);
  if (element.id !== "") attributes.push(`id="${escapeAttribute(element.id)}"`);
  if (element.href !== "") attributes.push(`href="${escapeAttribute(element.href)}"`);
  if (element.src !== "") attributes.push(`src="${escapeAttribute(element.src)}"`);
  if (element.title !== "") attributes.push(`title="${escapeAttribute(element.title)}"`);
  if (element.tagName === "img") attributes.push(`alt="${escapeAttribute(element.alt)}"`);
  for (const [name, value] of Object.entries(element.attributes)) {
    attributes.push(`${name}="${escapeAttribute(value)}"`);
  }
  for (const [key, value] of Object.entries(element.dataset)) {
    attributes.push(`data-${toKebabCase(key)}="${escapeAttribute(value)}"`);
  }
  const style = serializeStyle(element.style);
  if (style !== "") attributes.push(`style="${escapeAttribute(style)}"`);

  const open = [element.tagName, ...attributes].join(" ");
  if (VOID_TAGS.has(element.tagName)) {
    return `<${open} />`;
  }
  if (element.children.length > 0) {
    return `<${open}>${element.children.map(serializeElement).join("")}</${element.tagName}>`;
  }
  // A `<style>` holds CSS, not markup: escaping it would corrupt the
  // `@font-face` rules the backend emits for embedded faces.
  const inner = RAW_TEXT_TAGS.has(element.tagName)
    ? element.textContent
    : escapeText(element.textContent);
  return `<${open}>${inner}</${element.tagName}>`;
};

const BLOB_URL_RE = /blob:[^\s)"']+/gu;

const toDataUrl = async (blobUrl: string): Promise<string> => {
  const response = await fetch(blobUrl);
  const bytes = new Uint8Array(await response.arrayBuffer());
  const mimeType = response.headers.get("content-type") ?? "application/octet-stream";
  URL.revokeObjectURL(blobUrl);
  return `data:${mimeType};base64,${Buffer.from(bytes).toString("base64")}`;
};

/**
 * The backend mints a `blob:` URL for an image or an embedded face whenever
 * one is available, and Bun has `URL.createObjectURL`. A blob URL is scoped to
 * this process, so the page would load nothing; the bytes are read back here
 * and inlined so the file needs no server.
 */
const inlineBlobUrls = async (element: StubElement): Promise<void> => {
  if (element.src.startsWith("blob:")) {
    element.src = await toDataUrl(element.src);
  }
  for (const blobUrl of new Set(element.textContent.match(BLOB_URL_RE))) {
    // oxlint-disable-next-line no-await-in-loop -- one face at a time; the set is small and bounded
    element.textContent = element.textContent.replaceAll(blobUrl, await toDataUrl(blobUrl));
  }
  for (const child of element.children) {
    // oxlint-disable-next-line no-await-in-loop -- bytes are read back in paint order
    await inlineBlobUrls(child);
  }
};

/**
 * Each page sits in a slot of whole pixels.
 *
 * A page height is fractional (A4 is 1122.52 px), so pages stacked in normal
 * flow start at fractional offsets and a screenshot of page 2 covers one more
 * device row than a screenshot of page 1. That is the harness measuring its
 * own page container, not a backend: the slot rounds the *offset* while the
 * page keeps its exact size, so every page is captured from an integer origin
 * exactly as `mutool` rasterizes from the page corner.
 */
const pageSlot = (element: StubElement, page: DisplayPage): string =>
  `<div style="position: relative; overflow: hidden; width: ${String(Math.ceil(page.widthPx))}px; height: ${String(Math.ceil(page.heightPx))}px">${serializeElement(element)}</div>`;

type BuildPageHtmlOptions = {
  readonly pages: readonly DisplayPage[];
  readonly elements: readonly StubElement[];
  readonly fontFaceCss: string;
};

const buildPageHtml = ({ pages, elements, fontFaceCss }: BuildPageHtmlOptions): string =>
  [
    "<!doctype html>",
    '<html><head><meta charset="utf-8"><title>paint equivalence</title><style>',
    fontFaceCss,
    "html, body { margin: 0; padding: 0; background: #fff; }",
    "</style></head><body>",
    elements
      .map((element, index) => {
        const page = pages.at(index);
        return page === undefined ? serializeElement(element) : pageSlot(element, page);
      })
      .join(""),
    "</body></html>",
  ].join("\n");

// ---------------------------------------------------------------------------
// The two raster arms
// ---------------------------------------------------------------------------

type RasterArm =
  | { readonly status: "rendered"; readonly pagePngs: readonly string[] }
  | { readonly status: "skipped"; readonly reason: string };

const pageNumberOf = (filename: string): number => {
  const match = PAGE_PNG_RE.exec(filename);
  return match?.[1] === undefined ? 0 : Number(match[1]);
};

/**
 * A page directory is emptied before it is filled. The PDF arm names its
 * output by glob, so a run over a document that lost a page would otherwise
 * compare this run's pages against a stale one left behind by the last.
 */
const freshDir = async (directory: string): Promise<void> => {
  await rm(directory, { recursive: true, force: true });
  await mkdir(directory, { recursive: true });
};

const listPagePngs = async (directory: string): Promise<string[]> => {
  const entries = await readdir(directory);
  return entries
    .filter((name) => PAGE_PNG_RE.test(name))
    .sort((a, b) => pageNumberOf(a) - pageNumberOf(b))
    .map((name) => path.join(directory, name));
};

type RasterizePdfOptions = {
  readonly pdfPath: string;
  readonly pagesDir: string;
};

const rasterizePdf = async ({ pdfPath, pagesDir }: RasterizePdfOptions): Promise<RasterArm> => {
  if (Bun.which("mutool") === null) {
    return { status: "skipped", reason: "mutool is not on PATH; the PDF arm did not rasterize." };
  }
  await freshDir(pagesDir);
  const proc = Bun.spawn(
    [
      "mutool",
      "draw",
      "-F",
      "png",
      "-r",
      String(RASTER_DPI),
      "-o",
      path.join(pagesDir, "p%d.png"),
      pdfPath,
    ],
    { stdout: "ignore", stderr: "pipe" },
  );
  const [exitCode, stderr] = await Promise.all([proc.exited, new Response(proc.stderr).text()]);
  if (exitCode !== 0) {
    return { status: "skipped", reason: `mutool exited ${String(exitCode)}: ${stderr.trim()}` };
  }
  return { status: "rendered", pagePngs: await listPagePngs(pagesDir) };
};

type ScreenshotDomOptions = {
  readonly htmlPath: string;
  readonly pagesDir: string;
  readonly pageCount: number;
};

const screenshotDom = async ({
  htmlPath,
  pagesDir,
  pageCount,
}: ScreenshotDomOptions): Promise<RasterArm> => {
  await freshDir(pagesDir);
  const launched = await Result.tryPromise({
    try: () => chromium.launch({ headless: true }),
    catch: (cause) => (cause instanceof Error ? cause.message : String(cause)),
  });
  if (launched.isErr()) {
    return {
      status: "skipped",
      reason: launched.error.includes(CHROMIUM_MISSING_MARKER)
        ? CHROMIUM_MISSING_MESSAGE
        : `chromium could not launch: ${launched.error}`,
    };
  }

  const browser = launched.value;
  const captured = await Result.tryPromise({
    try: async () => {
      const context = await browser.newContext({
        viewport: VIEWPORT,
        deviceScaleFactor: 1,
        colorScheme: "light",
      });
      const page = await context.newPage();
      await page.goto(pathToFileURL(htmlPath).href, { waitUntil: "load" });
      await page.evaluate(async () => {
        await document.fonts.ready;
      });

      const pagePngs: string[] = [];
      for (let index = 0; index < pageCount; index++) {
        const locator = page.locator(PAGE_SELECTOR).nth(index);
        const pngPath = path.join(pagesDir, `p${String(index + 1)}.png`);
        // oxlint-disable-next-line no-await-in-loop -- pages are captured in order so filenames stay page-aligned
        await locator.scrollIntoViewIfNeeded();
        // oxlint-disable-next-line no-await-in-loop -- one capture at a time keeps the rendered layout stable
        await locator.screenshot({ path: pngPath, timeout: SCREENSHOT_TIMEOUT_MS });
        pagePngs.push(pngPath);
      }
      return pagePngs;
    },
    catch: (cause) => (cause instanceof Error ? cause.message : String(cause)),
  });
  await browser.close();

  // A browser that launches and then fails mid-capture is the same kind of
  // event as one that never launched: a tool problem, not a paint regression.
  // Crashing the run here would fail a fixture for the machine being busy.
  return captured.isErr()
    ? { status: "skipped", reason: `chromium could not capture the pages: ${captured.error}` }
    : { status: "rendered", pagePngs: captured.value };
};

// ---------------------------------------------------------------------------
// One fixture
// ---------------------------------------------------------------------------

type PageScore = {
  readonly page: number;
  readonly status: RasterPageComparison["status"];
  readonly similarity: number;
  readonly diffPixels: number;
  readonly totalPixels: number;
};

type FixtureReport = {
  readonly fixture: string;
  readonly pageCount: number;
  readonly pdfByteSize: number;
  readonly exportMs: number;
  readonly meanSimilarity: number | null;
  /** `comparePageRasters`' pixel-weighted score, for contrast with the mean. */
  readonly pixelWeightedScore: number | null;
  readonly worstPage: PageScore | null;
  readonly pages: readonly PageScore[];
  readonly skipped: readonly string[];
  readonly unsupported: readonly DisplayUnsupported[];
  readonly layoutGaps: readonly HeadlessLayoutGap[];
  readonly measurementSubstitutions: readonly HeadlessFontSubstitution[];
  readonly embeddingSubstitutions: readonly PdfSubstitution[];
};

const slugOf = (fixturePath: string): string =>
  path.basename(fixturePath, path.extname(fixturePath)).replaceAll(/[^\w.-]+/gu, "-");

const meanOf = (values: readonly number[]): number | null =>
  values.length === 0 ? null : values.reduce((sum, value) => sum + value, 0) / values.length;

const toFontRequest = ({ family, weight, italic }: DisplayFontFace) => ({
  family,
  bold: weight >= BOLD_WEIGHT,
  italic,
});

type RunFixtureOptions = {
  readonly fixturePath: string;
  readonly outputDir: string;
};

type RunFixtureResult = Result<FixtureReport, HeadlessLayoutError | WritePdfError>;

const runFixture = async ({
  fixturePath,
  outputDir,
}: RunFixtureOptions): Promise<RunFixtureResult> => {
  const fixtureDir = path.join(outputDir, slugOf(fixturePath));
  await mkdir(fixtureDir, { recursive: true });

  const fonts = createBundledFontSource();
  const headless = installHeadlessMeasureProvider(fonts);

  const startedAt = performance.now();
  const laidOut = await layoutDocxHeadless(await readFile(fixturePath), { pageGap: 0 });
  if (laidOut.isErr()) {
    return Result.err(laidOut.error);
  }

  // One display list, both arms: any difference below is backend divergence.
  const list = buildDisplayList({
    layout: laidOut.value.layout,
    blockLookup: laidOut.value.blockLookup,
  });

  const written = writePdf(list, {
    fonts: { load: (face) => fonts.load(toFontRequest(face)) },
    timestamp: FIXED_TIMESTAMP,
    producer: PRODUCER,
  });
  if (written.isErr()) {
    return Result.err(written.error);
  }
  const exportMs = performance.now() - startedAt;

  const pdfPath = path.join(fixtureDir, "document.pdf");
  await writeFile(pdfPath, written.value.bytes);

  // No `pageBackground` here: the builder already emitted the canvas as the
  // page's first primitive, so both backends paint it from the same
  // instruction. Passing it again would put a mark in one arm only.
  const domPages = renderDisplayListToDom(list, { doc: stubDocument() }).map(asStub);
  for (const domPage of domPages) {
    // oxlint-disable-next-line no-await-in-loop -- bytes are read back in page order
    await inlineBlobUrls(domPage);
  }
  const htmlPath = path.join(fixtureDir, "page.html");
  await writeFile(
    htmlPath,
    buildPageHtml({
      pages: list.pages,
      elements: domPages,
      // The list's own families, so a face the source resolved through the
      // fallback chain is declared under the name the DOM backend emits.
      fontFaceCss: bundledFontFaceCss({ families: list.fonts.map((face) => face.family) }),
    }),
  );

  const [pdfArm, domArm] = await Promise.all([
    rasterizePdf({ pdfPath, pagesDir: path.join(fixtureDir, "pdf-pages") }),
    screenshotDom({
      htmlPath,
      pagesDir: path.join(fixtureDir, "dom-pages"),
      pageCount: list.pages.length,
    }),
  ]);

  const shared = {
    fixture: path.relative(REPO_ROOT, fixturePath),
    pageCount: list.pages.length,
    pdfByteSize: written.value.bytes.byteLength,
    exportMs,
    unsupported: list.unsupported,
    layoutGaps: laidOut.value.unsupported,
    measurementSubstitutions: headless.substitutions(),
    embeddingSubstitutions: written.value.substitutions,
  };
  const unscored = {
    meanSimilarity: null,
    pixelWeightedScore: null,
    worstPage: null,
    pages: [],
  } as const;

  if (pdfArm.status === "skipped" || domArm.status === "skipped") {
    const skipped = [pdfArm, domArm].flatMap((arm) =>
      arm.status === "skipped" ? [arm.reason] : [],
    );
    return Result.ok({ ...shared, ...unscored, skipped });
  }

  const { comparison } = await comparePageRasters({
    referencePagePngs: [...pdfArm.pagePngs],
    folioPagePngs: [...domArm.pagePngs],
    outputDir: path.join(fixtureDir, "diff"),
  });
  if (comparison.status === "empty") {
    return Result.ok({
      ...shared,
      ...unscored,
      skipped: ["Neither arm produced a page raster."],
    });
  }

  const pages = comparison.pages.map(
    ({ page, status, similarity, diffPixels, totalPixels }): PageScore => ({
      page,
      status,
      similarity,
      diffPixels,
      totalPixels,
    }),
  );
  let worstPage: PageScore | null = null;
  for (const candidate of pages) {
    if (worstPage === null || candidate.similarity < worstPage.similarity) {
      worstPage = candidate;
    }
  }

  return Result.ok({
    ...shared,
    meanSimilarity: meanOf(pages.map((entry) => entry.similarity)),
    pixelWeightedScore: comparison.score,
    worstPage,
    pages,
    skipped: [],
  });
};

// ---------------------------------------------------------------------------
// Baseline
// ---------------------------------------------------------------------------

type Baseline = Record<string, number>;

const round = (value: number): number => Number(value.toFixed(BASELINE_DECIMALS));

const readBaseline = async (): Promise<Baseline | null> => {
  const file = Bun.file(BASELINE_PATH);
  if (!(await file.exists())) {
    return null;
  }
  const parsed: unknown = await file.json();
  if (typeof parsed !== "object" || parsed === null) {
    return null;
  }
  return Object.fromEntries(
    Object.entries(parsed).flatMap(([fixture, value]) =>
      typeof value === "number" ? [[fixture, value] as const] : [],
    ),
  );
};

type Regression = {
  readonly fixture: string;
  readonly baseline: number;
  readonly current: number;
  readonly delta: number;
};

const findRegressions = (
  reports: readonly FixtureReport[],
  baseline: Baseline,
): readonly Regression[] =>
  reports.flatMap((report) => {
    const previous = baseline[report.fixture];
    if (previous === undefined || report.meanSimilarity === null) {
      return [];
    }
    const delta = report.meanSimilarity - previous;
    return -delta > BASELINE_REGRESSION_TOLERANCE
      ? [
          {
            fixture: report.fixture,
            baseline: previous,
            current: report.meanSimilarity,
            delta,
          },
        ]
      : [];
  });

/**
 * Baseline fixtures this run did not cover.
 *
 * Only meaningful for the default corpus: a caller who named one fixture or a
 * directory did not ask for the others, so counting them as missing would fail
 * every targeted run and teach the same "ignore the exit code" habit the
 * baseline gate exists to avoid.
 */
const findMissing = (reports: readonly FixtureReport[], baseline: Baseline): readonly string[] => {
  const ran = new Set(reports.map((report) => report.fixture));
  return Object.keys(baseline).filter((fixture) => !ran.has(fixture));
};

// ---------------------------------------------------------------------------
// Reporting
// ---------------------------------------------------------------------------

const formatSimilarity = (value: number): string => value.toFixed(4);

/** Rounded before the sign is chosen, so a movement too small to print does
 * not render as a puzzling `-0.0000`. */
const formatDelta = (value: number): string => {
  const rounded = Number(value.toFixed(BASELINE_DECIMALS));
  return `${rounded < 0 ? "-" : "+"}${Math.abs(rounded).toFixed(BASELINE_DECIMALS)}`;
};

/** `0.9419 (baseline 0.9419, +0.0000)`, so movement is visible either way. */
const formatMean = (mean: number | null, baseline: number | undefined): string => {
  if (mean === null) {
    return "n/a";
  }
  if (baseline === undefined) {
    return `${formatSimilarity(mean)} (new)`;
  }
  return `${formatSimilarity(mean)} (baseline ${formatSimilarity(baseline)}, ${formatDelta(mean - baseline)})`;
};

const printReport = (report: FixtureReport, baseline: Baseline | null): void => {
  console.log(report.fixture);
  for (const page of report.pages) {
    console.log(
      `  page ${String(page.page)}  ${formatSimilarity(page.similarity)}  ` +
        `${page.diffPixels.toLocaleString("en")} / ${page.totalPixels.toLocaleString("en")} px  ${page.status}`,
    );
  }
  const worst =
    report.worstPage === null
      ? "n/a"
      : `page ${String(report.worstPage.page)} (${formatSimilarity(report.worstPage.similarity)})`;
  const mean = formatMean(report.meanSimilarity, baseline?.[report.fixture]);
  console.log(
    `  ${String(report.pageCount)} page(s), ${report.pdfByteSize.toLocaleString("en")} pdf bytes, ` +
      `${report.exportMs.toFixed(0)} ms export, worst ${worst}, mean ${mean}`,
  );
  for (const reason of report.skipped) {
    console.log(`  skipped: ${reason}`);
  }
  for (const gap of report.layoutGaps) {
    console.log(`  layout gap: ${gap.story} (${gap.detail})`);
  }
  for (const entry of report.unsupported) {
    console.log(
      `  unsupported: ${entry.construct} on page ${String(entry.pageIndex + 1)} (${entry.detail})`,
    );
  }
  for (const { requested, reason } of report.measurementSubstitutions) {
    const face = `${requested.family}${requested.bold ? " bold" : ""}${requested.italic ? " italic" : ""}`;
    console.log(`  substituted (measure): ${face} (${reason})`);
  }
  for (const substitution of report.embeddingSubstitutions) {
    console.log(`  substituted (embed): ${JSON.stringify(substitution)}`);
  }
};

// ---------------------------------------------------------------------------
// Entry
// ---------------------------------------------------------------------------

const resolveFixtures = async (target: string): Promise<readonly string[]> => {
  const absolute = path.resolve(target);
  const info = await stat(absolute);
  if (!info.isDirectory()) {
    return [absolute];
  }
  const entries = await readdir(absolute);
  return entries
    .filter((name) => name.toLowerCase().endsWith(DOCX_EXTENSION))
    .sort((a, b) => a.localeCompare(b, "en"))
    .map((name) => path.join(absolute, name));
};

const args = parseArgs(process.argv.slice(2));
if (!args) {
  console.error(USAGE);
  process.exit(2);
}
if (args.target !== null && !existsSync(path.resolve(args.target))) {
  console.error(`No such fixture or directory: ${args.target}`);
  process.exit(2);
}

const fixtures =
  args.target === null
    ? DEFAULT_FIXTURES.map((relative) => path.join(REPO_ROOT, relative))
    : await resolveFixtures(args.target);
if (fixtures.length === 0) {
  console.error(`No ${DOCX_EXTENSION} fixtures found.`);
  process.exit(2);
}
await mkdir(args.outputDir, { recursive: true });

const reports: FixtureReport[] = [];
for (const fixturePath of fixtures) {
  // oxlint-disable-next-line no-await-in-loop -- fixtures share one measure provider, so they run in sequence
  const report = await runFixture({ fixturePath, outputDir: args.outputDir });
  if (report.isErr()) {
    printFailure(path.relative(REPO_ROOT, fixturePath), report.error);
    process.exit(1);
  }
  reports.push(report.value);
}

const baseline = await readBaseline();
const regressions = baseline === null ? [] : findRegressions(reports, baseline);
// Only the default corpus is the set the baseline describes; see findMissing.
const missing = baseline === null || args.target !== null ? [] : findMissing(reports, baseline);
const newFixtures =
  baseline === null
    ? []
    : reports
        .filter((report) => baseline[report.fixture] === undefined)
        .map((report) => report.fixture);

if (args.updateBaseline) {
  const next: Baseline = Object.fromEntries(
    reports.flatMap((report) =>
      report.meanSimilarity === null ? [] : [[report.fixture, round(report.meanSimilarity)]],
    ),
  );
  await writeFile(BASELINE_PATH, `${JSON.stringify(next, null, 2)}\n`);
}

const scored = reports.flatMap((report) =>
  report.meanSimilarity === null ? [] : [report.meanSimilarity],
);
/** An absolute floor applies only when the caller named one. */
const belowThreshold =
  args.threshold === null
    ? []
    : reports.filter(
        (report) => report.meanSimilarity !== null && report.meanSimilarity < args.threshold,
      );

if (args.json) {
  console.log(
    JSON.stringify(
      {
        threshold: args.threshold,
        regressionTolerance: BASELINE_REGRESSION_TOLERANCE,
        outputDir: args.outputDir,
        timestamp: FIXED_TIMESTAMP,
        meanSimilarity: meanOf(scored),
        regressions,
        missingFromRun: missing,
        newFixtures,
        belowThreshold: belowThreshold.map((report) => report.fixture),
        fixtures: reports,
      },
      null,
      2,
    ),
  );
} else {
  for (const report of reports) {
    printReport(report, baseline);
  }
  for (const { fixture, baseline: previous, current, delta } of regressions) {
    console.log(
      `regressed: ${fixture} ${formatSimilarity(previous)} -> ${formatSimilarity(current)} (${formatDelta(delta)})`,
    );
  }
  for (const fixture of missing) {
    console.log(`missing: ${fixture} is in the baseline but did not run`);
  }
  for (const fixture of newFixtures) {
    console.log(`new: ${fixture} has no baseline entry; run --update-baseline to record it`);
  }
  for (const report of belowThreshold) {
    console.log(
      `below threshold: ${report.fixture} ${formatSimilarity(report.meanSimilarity ?? 0)} < ${formatSimilarity(args.threshold ?? 0)}`,
    );
  }
  const overall = meanOf(scored);
  const gate =
    args.threshold === null
      ? `baseline +/- ${String(BASELINE_REGRESSION_TOLERANCE)}`
      : `baseline +/- ${String(BASELINE_REGRESSION_TOLERANCE)} and threshold ${formatSimilarity(args.threshold)}`;
  console.log(
    `${String(reports.length)} fixture(s), mean ${overall === null ? "n/a" : formatSimilarity(overall)}, ` +
      `gate ${gate}, artifacts in ${args.outputDir}`,
  );
}

process.exit(regressions.length > 0 || missing.length > 0 || belowThreshold.length > 0 ? 1 : 0);
