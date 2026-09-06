/**
 * Lay a `.docx` package out to pages without a browser.
 *
 * The editor reaches `Layout` through `controller/layoutPipeline.ts`, which
 * needs an `EditorState`, a painter and a live DOM. Export, agents and CI need
 * the same pages with none of that, so this module walks the four pure stages
 * directly: parse, project to ProseMirror, flatten to flow blocks, measure and
 * paginate.
 *
 * ## What it does not do
 *
 * Header, footer and footnote stories are separate OOXML stories that the
 * editor lays out as their own ProseMirror views. This entry paginates the
 * body only, and says so through {@link HeadlessLayoutResult.unsupported}
 * rather than emitting pages that quietly lack their furniture.
 *
 * ## Measurement
 *
 * Layout is only as good as the installed `MeasureProvider`. The caller
 * installs one (`installHeadlessMeasureProvider` for a server, the canvas
 * backend in a browser); this module refuses to guess, because a silently
 * wrong provider produces a plausible layout that is wrong everywhere.
 */

import { Result, TaggedError } from "better-result";
import type { Node as PMNode } from "prosemirror-model";

import { parseDocx } from "./docx/parser";
import type { DocxInput } from "./utils/docxInput";
import { buildFontAlternates } from "./fonts/fontAlternates";
import { layoutDocument } from "./layout-engine/index";
import { getMeasureProvider } from "./layout-engine/measure/measureProvider";
import { measureBlocks } from "./layout-engine/measure/measureBlocks";
import type { FlowBlock, Layout, LayoutOptions, Measure } from "./layout-engine/types";
import type { BlockLookup } from "./layout-painter/index";
import { toFlowBlocks } from "./layout-bridge/convert/toFlowBlocks";
import type { ToFlowBlocksOptions } from "./layout-bridge/convert/toFlowBlocks";
import { getMargins, getPageSize, getPageNumbering } from "./paged-layout/sectionGeometry";
import { toProseDoc } from "./prosemirror/conversion/toProseDoc";
import type { Document } from "./types/document";

/** A story this entry does not paginate, named so a caller can see the hole. */
export type HeadlessLayoutGap = {
  readonly story: "header" | "footer" | "footnote" | "endnote";
  readonly detail: string;
};

export class HeadlessLayoutError extends TaggedError("HeadlessLayoutError")<{
  message: string;
  cause?: unknown;
}> {}

export type HeadlessLayoutOptions = {
  /** Gap between pages. Only a viewer cares; export passes 0. */
  readonly pageGap?: number;
};

export type HeadlessLayoutResult = {
  readonly layout: Layout;
  /** Block and measure per fragment `blockId`, as the painter expects them. */
  readonly blockLookup: BlockLookup;
  readonly document: Document;
  /**
   * The editable projection the pages were laid out from. Returned because a
   * fragment addresses its content by ProseMirror range, so reading a page's
   * text (`getPageTextFromLayout`) is impossible without it.
   */
  readonly proseDoc: PMNode;
  readonly unsupported: readonly HeadlessLayoutGap[];
};

/**
 * Stories the editor lays out separately. Listed once, as data, so the gap
 * reported to callers cannot drift from the gap that actually exists.
 */
const UNPAGINATED_STORIES = [
  {
    story: "header",
    detail: "Header stories are laid out by the editor pipeline; this entry paginates the body.",
  },
  {
    story: "footer",
    detail: "Footer stories are laid out by the editor pipeline; this entry paginates the body.",
  },
  {
    story: "footnote",
    detail: "Footnote heights come from the editor's footnote views; body lines reserve no space.",
  },
] as const satisfies readonly HeadlessLayoutGap[];

const hasStory = (document: Document, story: HeadlessLayoutGap["story"]): boolean => {
  const sections = document.package.document.sections ?? [];
  switch (story) {
    case "header":
      return sections.some((section) => (section.properties.headerReferences?.length ?? 0) > 0);
    case "footer":
      return sections.some((section) => (section.properties.footerReferences?.length ?? 0) > 0);
    case "footnote":
      return (document.package.footnotes?.length ?? 0) > 0;
    case "endnote":
      return (document.package.endnotes?.length ?? 0) > 0;
    default:
      story satisfies never;
      return false;
  }
};

const buildFlowOptions = (document: Document, pageContentHeight: number): ToFlowBlocksOptions => {
  const settings = document.package.settings;
  const options: ToFlowBlocksOptions = {
    pageContentHeight,
    fontAlternates: buildFontAlternates(document.package.fontTable),
  };
  const theme = document.package.theme;
  if (theme) {
    options.theme = theme;
  }
  if (settings?.defaultTabStop !== undefined) {
    options.defaultTabStopTwips = settings.defaultTabStop;
  }
  if (settings?.lineBreakRules) {
    options.lineBreakRules = settings.lineBreakRules;
  }
  if (settings?.autoHyphenation === true) {
    options.automaticHyphenation = {
      enabled: true,
      ...(settings.doNotHyphenateCaps === undefined
        ? {}
        : { doNotHyphenateCaps: settings.doNotHyphenateCaps }),
      ...(settings.consecutiveHyphenLimit === undefined
        ? {}
        : { consecutiveLineLimit: settings.consecutiveHyphenLimit }),
      ...(settings.hyphenationZoneTwips === undefined
        ? {}
        : { hyphenationZoneTwips: settings.hyphenationZoneTwips }),
    };
  }
  return options;
};

/**
 * `w:mirrorMargins` is parsed into `FolioDocumentSettings` but the shared
 * `Document` model types `settings` as the narrower `DocumentSettings`, so the
 * flag is present at runtime and invisible to the type. The React paged editor
 * narrows the same way; both should stop once the model carries the field.
 */
const readsMirrorMargins = (settings: Document["package"]["settings"]): boolean =>
  settings !== undefined && "mirrorMargins" in settings && settings.mirrorMargins === true;

const buildBlockLookup = (
  blocks: readonly FlowBlock[],
  measures: readonly Measure[],
): BlockLookup =>
  new Map(
    blocks.flatMap((block, index) => {
      const measure = measures.at(index);
      return measure === undefined ? [] : [[String(block.id), { block, measure }] as const];
    }),
  );

/**
 * Parse and paginate a package. Fails rather than defaulting when no
 * measurement backend is installed: a layout measured by the wrong provider
 * is wrong in a way no downstream check catches.
 */
export const layoutDocxHeadless = async (
  input: DocxInput,
  options: HeadlessLayoutOptions = {},
): Promise<Result<HeadlessLayoutResult, HeadlessLayoutError>> => {
  const providerCheck = Result.try(() => getMeasureProvider().getFontMetrics({ fontSize: 11 }));
  if (providerCheck.isErr()) {
    return Result.err(
      new HeadlessLayoutError({
        message:
          "No MeasureProvider is installed. Install one (installHeadlessMeasureProvider on a " +
          "server, installCanvasMeasureProvider in a browser) before laying out.",
        cause: providerCheck.error,
      }),
    );
  }

  const parsed = await Result.tryPromise({
    try: () => parseDocx(input, { preloadFonts: false }),
    catch: (cause) =>
      new HeadlessLayoutError({ message: "The package could not be parsed.", cause }),
  });
  if (parsed.isErr()) {
    return Result.err(parsed.error);
  }
  const document = parsed.value;

  const projected = Result.try(() =>
    toProseDoc(document, {
      ...(document.package.styles ? { styles: document.package.styles } : {}),
      ...(document.package.theme ? { theme: document.package.theme } : {}),
    }),
  );
  if (projected.isErr()) {
    return Result.err(
      new HeadlessLayoutError({
        message: "The document could not be projected to the editable model.",
        cause: projected.error,
      }),
    );
  }

  const sections = document.package.document.sections ?? [];
  const firstSection = sections.at(0)?.properties;
  const finalSection = sections.at(-1)?.properties ?? firstSection;
  const pageSize = getPageSize(firstSection);
  const margins = getMargins(firstSection);
  const contentWidth = pageSize.w - margins.left - margins.right;
  const pageContentHeight = pageSize.h - margins.top - margins.bottom;

  const laidOut = Result.try(() => {
    const blocks = toFlowBlocks(projected.value, buildFlowOptions(document, pageContentHeight));
    const measures = measureBlocks(blocks, contentWidth);
    const finalPageSize = getPageSize(finalSection);
    const finalMargins = getMargins(finalSection);
    const layoutOptions: LayoutOptions = {
      pageSize,
      margins,
      pageNumbering: getPageNumbering(firstSection),
      finalPageSize,
      finalMargins,
      finalPageNumbering: getPageNumbering(finalSection),
      pageGap: options.pageGap ?? 0,
      titlePage: firstSection?.titlePg === true,
      evenAndOddHeaders: document.package.settings?.evenAndOddHeaders === true,
      mirrorMargins: readsMirrorMargins(document.package.settings),
    };
    return {
      layout: layoutDocument(blocks, measures, layoutOptions),
      blockLookup: buildBlockLookup(blocks, measures),
    };
  });
  if (laidOut.isErr()) {
    return Result.err(
      new HeadlessLayoutError({
        message: "The document could not be paginated.",
        cause: laidOut.error,
      }),
    );
  }

  return Result.ok({
    layout: laidOut.value.layout,
    blockLookup: laidOut.value.blockLookup,
    document,
    proseDoc: projected.value,
    unsupported: UNPAGINATED_STORIES.filter((gap) => hasStory(document, gap.story)),
  });
};
