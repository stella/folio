/**
 * Render a document with folio's own layout and painters, headlessly: PDF
 * from the PDF backend, HTML from the DOM backend, and PNG by screenshotting
 * that HTML in Chromium. Every format paints the one display list built from
 * one layout, measured with the `@fontsource` faces the CLI installs.
 *
 * PNG needs a browser. `playwright-core` is an optional peer dependency and
 * is imported only when a PNG is asked for.
 */

import { Result } from "better-result";
import { existsSync, readFileSync } from "node:fs";
import { createRequire } from "node:module";
import path from "node:path";

import { selectDisplayPages } from "@stll/folio-core/display-list/selectDisplayPages";
import { renderDisplayListToHtml } from "@stll/folio-core/display-list/html/renderDisplayListToHtml";
import type { DisplayList } from "@stll/folio-core/display-list/types";
import { buildDocxDisplayList, writeDisplayListPdf } from "@stll/folio-core/export-pdf";
import {
  createFontsourceFaces,
  type FontsourceFaces,
} from "@stll/folio-core/fonts/fontsourceFaces";

import type { LoadedFile } from "./document";
import { cliError, FOLIO_CLI_ERROR_CODES, type FolioCliError } from "./errors";

export const RENDER_FORMATS = ["pdf", "png", "html"] as const;

export type RenderFormat = (typeof RENDER_FORMATS)[number];

const requireFromHere = createRequire(import.meta.url);

/** The installed `@fontsource/<name>` directory, or `null` when it is not installed. */
const fontsourceDirectory = (packageName: string): string | null => {
  const resolved = Result.try(() =>
    requireFromHere.resolve(`@fontsource/${packageName}/package.json`),
  );
  return resolved.isOk() ? path.dirname(resolved.value) : null;
};

let faces: FontsourceFaces | undefined;

/** The CLI's `@fontsource` faces, shared by every render in the process. */
export const cliFontFaces = (): FontsourceFaces => {
  faces ??= createFontsourceFaces({
    read: (packageName, relativePath) => {
      const directory = fontsourceDirectory(packageName);
      if (directory === null) return null;
      const filePath = path.join(directory, relativePath);
      return existsSync(filePath) ? new Uint8Array(readFileSync(filePath)) : null;
    },
  });
  return faces;
};

const renderError = (message: string): FolioCliError =>
  cliError({ code: FOLIO_CLI_ERROR_CODES.invalidDocument, message });

const describe = (error: unknown): string =>
  error instanceof Error ? error.message : String(error);

/** The document's display list, measured with the CLI's faces. */
export const buildDisplayList = async (
  file: LoadedFile,
): Promise<Result<DisplayList, FolioCliError>> => {
  const built = await buildDocxDisplayList(file.bytes.slice().buffer, {
    fonts: cliFontFaces().source,
  });
  return built.isOk()
    ? Result.ok(built.value.list)
    : Result.err(renderError(`${file.path} could not be laid out: ${built.error.message}`));
};

/** Refuse page numbers the document does not have; `undefined` keeps every page. */
export const checkPages = (
  list: DisplayList,
  pages: readonly number[] | undefined,
): Result<DisplayList, FolioCliError> => {
  if (pages === undefined) return Result.ok(list);
  const missing = pages.filter(
    (page) => !Number.isInteger(page) || page < 1 || page > list.pages.length,
  );
  if (missing.length > 0) {
    return Result.err(
      cliError({
        code: FOLIO_CLI_ERROR_CODES.invalidInput,
        message: `The document has ${list.pages.length} page${list.pages.length === 1 ? "" : "s"}; no page ${missing.join(", ")}.`,
      }),
    );
  }
  return Result.ok(
    selectDisplayPages(
      list,
      pages.map((page) => page - 1),
    ),
  );
};

/** Standalone HTML for the list, with its faces inlined. */
export const displayListHtml = (list: DisplayList, title: string): string =>
  renderDisplayListToHtml(list, {
    title,
    fontFaceCss: cliFontFaces().fontFaceCss({ families: list.fonts.map((face) => face.family) }),
    pageGapPx: 16,
    canvasColor: "#e8e8e8",
  });

type RenderPdfOptions = {
  file: LoadedFile;
  pages: readonly number[] | undefined;
  timestamp: string;
};

export type RenderedOutput = { bytes: Uint8Array; pageCount: number };

export const renderPdf = async ({
  file,
  pages,
  timestamp,
}: RenderPdfOptions): Promise<Result<RenderedOutput, FolioCliError>> => {
  const list = await buildDisplayList(file);
  if (list.isErr()) return Result.err(list.error);
  const selected = checkPages(list.value, pages);
  if (selected.isErr()) return Result.err(selected.error);
  const exported = await writeDisplayListPdf(selected.value, {
    fonts: cliFontFaces().source,
    timestamp,
    producer: "folio",
  });
  return exported.isOk()
    ? Result.ok({ bytes: exported.value.bytes, pageCount: exported.value.pageCount })
    : Result.err(renderError(`${file.path} could not be exported: ${exported.error.message}`));
};

/** What `playwright-core` exposes that this module uses. */
type ChromiumLike = {
  launch: (options: { headless: boolean }) => Promise<{
    newPage: (options: { deviceScaleFactor: number }) => Promise<{
      setContent: (html: string, options: { waitUntil: "load" }) => Promise<void>;
      evaluate: (script: string) => Promise<unknown>;
      locator: (selector: string) => {
        first: () => { screenshot: (options: { type: "png" }) => Promise<Uint8Array> };
      };
    }>;
    close: () => Promise<void>;
  }>;
};

const hasChromium = (module: unknown): module is { chromium: ChromiumLike } =>
  typeof module === "object" &&
  module !== null &&
  "chromium" in module &&
  typeof module.chromium === "object" &&
  module.chromium !== null &&
  "launch" in module.chromium;

const rasterUnavailable = (detail: string): FolioCliError =>
  cliError({
    code: FOLIO_CLI_ERROR_CODES.rendererUnavailable,
    message: `PNG rendering needs Chromium through playwright-core: ${detail}`,
    hint: "Install it next to the CLI (npm install playwright-core) and its browser (npx playwright-core install chromium), or render a PDF.",
  });

type RenderPngOptions = { file: LoadedFile; page: number; scale: number };

/** One page as a PNG: the DOM backend's page, screenshotted in headless Chromium. */
export const renderPng = async ({
  file,
  page,
  scale,
}: RenderPngOptions): Promise<Result<RenderedOutput, FolioCliError>> => {
  const list = await buildDisplayList(file);
  if (list.isErr()) return Result.err(list.error);
  const selected = checkPages(list.value, [page]);
  if (selected.isErr()) return Result.err(selected.error);

  const imported = await Result.tryPromise(() => import("playwright-core"));
  if (imported.isErr()) return Result.err(rasterUnavailable("the package is not installed"));
  const module: unknown = imported.value;
  if (!hasChromium(module)) return Result.err(rasterUnavailable("the package has no chromium"));

  const shot = await Result.tryPromise({
    try: async () => {
      const browser = await module.chromium.launch({ headless: true });
      try {
        const tab = await browser.newPage({ deviceScaleFactor: scale });
        await tab.setContent(displayListHtml(selected.value, path.basename(file.path)), {
          waitUntil: "load",
        });
        await tab.evaluate("document.fonts.ready");
        return await tab.locator(".layout-page").first().screenshot({ type: "png" });
      } finally {
        await browser.close();
      }
    },
    catch: (error) => rasterUnavailable(describe(error)),
  });
  return shot.isOk() ? Result.ok({ bytes: shot.value, pageCount: 1 }) : Result.err(shot.error);
};
