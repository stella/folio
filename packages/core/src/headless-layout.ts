/**
 * Lay a `.docx` package out to pages without a browser.
 *
 * The editor reaches `Layout` through `controller/layoutPipeline.ts`, which
 * needs an `EditorState`, a painter and a live DOM. Export, agents and CI need
 * the same pages with none of that, so this module walks the four pure stages
 * directly: parse, project to ProseMirror, flatten to flow blocks, measure and
 * paginate.
 *
 * ## The other stories
 *
 * Headers, footers and footnotes are separate OOXML stories that the editor
 * lays out as their own ProseMirror views, but neither conversion needs a view:
 * `convertHeaderFooterToContent` and `buildFootnoteContentMap` run the same
 * `toFlowBlocks → measureBlocks` chain the body runs. This entry therefore lays
 * them out too and returns them as {@link HeadlessLayoutResult.furniture}, so
 * an export paints the pages the editor paints rather than bare bodies.
 * Endnotes remain unpaginated and say so through
 * {@link HeadlessLayoutResult.unsupported}.
 *
 * What this entry still does not do is push the body down for a header taller
 * than its margin: the editor's margin clearing lives in the controller, and a
 * document whose header overflows its margin will overlap the body here.
 *
 * ## Measurement
 *
 * Layout is only as good as the installed `MeasureProvider`. The caller
 * installs one (`installHeadlessMeasureProvider` for a server, the canvas
 * backend in a browser); this module refuses to guess, because a silently
 * wrong provider produces a plausible layout that is wrong everywhere.
 */

import { createHash } from "node:crypto";

import { Result, TaggedError } from "better-result";
import type { Node as PMNode } from "prosemirror-model";

import { parseDocx } from "./docx/parser";
import { toArrayBuffer, type DocxInput } from "./utils/docxInput";
import { buildFontAlternates } from "./fonts/fontAlternates";
import { extractEmbeddedFonts, type EmbeddedFont } from "./fonts/embeddedFonts";
import { buildHeaderFooterFieldValues } from "./fields/resolveFieldValues";
import { layoutDocument } from "./layout-engine/index";
import { getMeasureProvider } from "./layout-engine/measure/measureProvider";
import { measureBlocks } from "./layout-engine/measure/measureBlocks";
import { resolveSectionHeaderFooterRefs } from "./layout-engine/headerFooterRefs";
import { FOOTNOTE_ENTRY_MARGIN_BOTTOM } from "./layout-engine/types";
import type {
  FlowBlock,
  FootnoteContent,
  HeaderFooterContent,
  Layout,
  LayoutOptions,
  Measure,
} from "./layout-engine/types";
import type { BlockLookup } from "./layout-painter/index";
import {
  buildFootnoteContentMap,
  collectEndnoteRefs,
  collectFootnoteRefs,
  computeNoteDisplayNumbers,
  remapNoteMarkerText,
  type ConvertFootnoteOptions,
} from "./layout-bridge/convert/footnoteLayout";
import {
  convertHeaderFooterToContent,
  type HeaderFooterMetrics,
} from "./layout-bridge/convert/headerFooterLayout";
import { toFlowBlocks } from "./layout-bridge/convert/toFlowBlocks";
import type { ToFlowBlocksOptions } from "./layout-bridge/convert/toFlowBlocks";
import { getMargins, getPageSize, getPageNumbering } from "./paged-layout/sectionGeometry";
import type { DocumentFeatures } from "./display-list/build/buildDisplayList";
import type { PageFurnitureInputs } from "./display-list/build/furniture";
import { toProseDoc } from "./prosemirror/conversion/toProseDoc";
import { getDocumentWatermark } from "./watermark/index";
import type { Document, HeaderFooter, Watermark } from "./types/document";

/**
 * Constructs the painter takes from render options rather than from `Layout`.
 *
 * Imported from the display-list producer rather than restated, so a feature
 * added there cannot silently go unreported here: the producer derives the
 * shape from its own gap list, and this module has to fill whatever that list
 * names.
 */
export type DocumentPaintFeatures = DocumentFeatures;

/** A story this entry does not paginate, named so a caller can see the hole. */
export type HeadlessLayoutGap = {
  readonly story: "endnote";
  readonly detail: string;
};

export class HeadlessLayoutError extends TaggedError("HeadlessLayoutError")<{
  message: string;
  cause?: unknown;
}> {}

export type HeadlessLayoutOptions = {
  /** Gap between pages. Only a viewer cares; export passes 0. */
  readonly pageGap?: number;
  /**
   * The instant `DATE` and `TIME` fields in a header or footer are measured
   * with. Defaults to the wall clock; pass a fixed instant when two runs over
   * one package must produce byte-identical output.
   */
  readonly now?: Date;
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
  /**
   * Whether the package authors constructs that reach the editor's painter
   * through render options rather than through `Layout`. A display-list
   * producer cannot see them, and needs to know whether it is reporting a real
   * gap or a construct the document never had.
   */
  readonly documentFeatures: DocumentPaintFeatures;
  /**
   * The page furniture, in the shapes `buildDisplayList` takes: page borders,
   * the watermark, the header and footer stories by relationship id, and the
   * footnote bodies by note id. Spread it into the builder's options and the
   * exported pages carry what the editor's pages carry.
   */
  readonly furniture: PageFurnitureInputs;
  /**
   * The package's own font faces, de-obfuscated. Empty when `fontTable.xml`
   * declares no embedded face, so a backend that finds nothing here knows the
   * document expects host fonts rather than that extraction was skipped.
   */
  readonly embeddedFonts: readonly EmbeddedFont[];
  readonly unsupported: readonly HeadlessLayoutGap[];
};

/**
 * Stories this entry still does not lay out, keyed by story so the map is total
 * over the union. `satisfies Record<…>` is what makes it total in both
 * directions: a story left unpaginated but unlisted is a compile error rather
 * than content that vanishes without a report, and a story that starts being
 * laid out cannot keep a stale entry, because removing it narrows the union.
 *
 * Endnotes are the only one left. They are collected at the end of the
 * document rather than reserved per page, so nothing in `Layout` and nothing in
 * the painter's render options places them yet; laying them out is a
 * pagination feature, not a conversion this entry could call.
 */
const UNPAGINATED_STORIES = {
  endnote: {
    story: "endnote",
    detail: "Endnotes are their own story and are not paginated with the body.",
  },
} as const satisfies Record<HeadlessLayoutGap["story"], HeadlessLayoutGap>;

const hasStory = (document: Document, story: HeadlessLayoutGap["story"]): boolean => {
  switch (story) {
    case "endnote":
      return (document.package.endnotes?.length ?? 0) > 0;
    default:
      story satisfies never;
      return false;
  }
};

const readPaintFeatures = (document: Document): DocumentPaintFeatures => ({
  pageBorders: (document.package.document.sections ?? []).some((section) =>
    Object.values(section.properties.pageBorders ?? {}).some((border) => border !== undefined),
  ),
  watermark: getDocumentWatermark(document) !== undefined,
});

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
 * What every non-body story is converted with: the same styles, theme, font
 * alternates, tab grid and line-breaking rules the body was flattened with.
 * Derived from the body's own options rather than re-read from settings, so a
 * header cannot come to break lines differently from the paragraph above it.
 */
type StoryOptions = Omit<ConvertFootnoteOptions, "measureBlocks">;

const buildStoryOptions = (document: Document, flowOptions: ToFlowBlocksOptions): StoryOptions => ({
  ...(document.package.styles ? { styles: document.package.styles } : {}),
  ...(flowOptions.theme === undefined ? {} : { theme: flowOptions.theme }),
  ...(flowOptions.fontAlternates === undefined
    ? {}
    : { fontAlternates: flowOptions.fontAlternates }),
  ...(flowOptions.defaultTabStopTwips === undefined
    ? {}
    : { defaultTabStopTwips: flowOptions.defaultTabStopTwips }),
  ...(flowOptions.lineBreakRules === undefined
    ? {}
    : { lineBreakRules: flowOptions.lineBreakRules }),
  ...(flowOptions.automaticHyphenation === undefined
    ? {}
    : { automaticHyphenation: flowOptions.automaticHyphenation }),
});

type ConvertStoriesOptions = {
  readonly parts: ReadonlyMap<string, HeaderFooter> | undefined;
  readonly contentWidth: number;
  readonly metrics: HeaderFooterMetrics;
  readonly storyOptions: StoryOptions;
  /**
   * The page count page-number fields are measured at. A header is measured
   * once and painted on every page, so `PAGE` reserves the width of the largest
   * number rather than of the page it happens to be measured for.
   */
  readonly pageCount: number;
  readonly now: Date;
};

const convertStories = ({
  parts,
  contentWidth,
  metrics,
  storyOptions,
  pageCount,
  now,
}: ConvertStoriesOptions): ReadonlyMap<string, HeaderFooterContent> => {
  const contentByRId = new Map<string, HeaderFooterContent>();
  for (const [rId, part] of parts ?? []) {
    const content = convertHeaderFooterToContent(part, contentWidth, metrics, {
      ...storyOptions,
      measureBlocks: (blocks, width) =>
        measureBlocks(
          blocks,
          width,
          undefined,
          undefined,
          buildHeaderFooterFieldValues(blocks, pageCount, now),
        ),
      rId,
    });
    if (content) {
      contentByRId.set(rId, content);
    }
  }
  return contentByRId;
};

/**
 * Watermarks by the header part that carries them. Word puts the shape in a
 * header, so a document with `w:titlePg` or even/odd headers can carry a
 * different watermark per page; the per-rId map is what lets the builder pick
 * the right one instead of painting one header's watermark behind every page.
 */
const collectWatermarks = (
  headers: ReadonlyMap<string, HeaderFooter> | undefined,
): ReadonlyMap<string, Watermark> => {
  const byRId = new Map<string, Watermark>();
  for (const [rId, header] of headers ?? []) {
    if (header.watermark) {
      byRId.set(rId, header.watermark);
    }
  }
  return byRId;
};

/**
 * A picture watermark's bytes, as a `data:` URL.
 *
 * Resolved through `imageTarget`, the package path the relationship pointed at
 * when the watermark was parsed: relationship ids are scoped per header part
 * and repeat across parts, so the path is the only unambiguous handle. A linked
 * (external) image resolves to nothing, and the builder reports it.
 */
const resolveWatermarkImageSrc = (
  document: Document,
  watermark: Watermark | undefined,
): string | undefined => {
  if (watermark?.kind !== "picture" || watermark.imageTargetExternal === true) {
    return undefined;
  }
  const target = watermark.imageTarget;
  return target === undefined ? undefined : document.package.media?.get(target)?.dataUrl;
};

/**
 * The scope every extracted face's family carries here.
 *
 * `getEmbeddedFontFaces` defaults it to a fresh random id, which would make two
 * layouts of one package differ in a field the caller compares. Deriving it
 * from the package's own bytes keeps two layouts of one package identical
 * while keeping two *different* packages apart: a DOM backend registers each
 * face under this name, so a fixed scope would let two documents on one page
 * define competing faces under one name and leave which one wins to font
 * matching.
 */
const embeddedFontNonce = (source: ArrayBuffer): string =>
  createHash("sha256").update(new Uint8Array(source)).digest("hex").slice(0, NONCE_LENGTH);

/** Enough of the digest to separate packages without bloating every family. */
const NONCE_LENGTH = 12;

/** Whether `fontTable.xml` declares any embedded face at all. */
const declaresEmbeddedFonts = (document: Document): boolean =>
  (document.package.fontTable?.fonts ?? []).some(
    (font) =>
      font.embedRegular !== undefined ||
      font.embedBold !== undefined ||
      font.embedItalic !== undefined ||
      font.embedBoldItalic !== undefined,
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

  const flowOptions = buildFlowOptions(document, pageContentHeight);
  const storyOptions = buildStoryOptions(document, flowOptions);

  const laidOut = Result.try(() => {
    const authored = toFlowBlocks(projected.value, flowOptions);

    // Body markers carry the raw `w:id` as their text; Word paints the
    // reference-order number. Remapping before measurement is what keeps the
    // marker's measured width, the painted digits and the number on the body in
    // the footnote band all the same number.
    const footnotes = document.package.footnotes ?? [];
    const endnotes = document.package.endnotes ?? [];
    const footnoteRefs = collectFootnoteRefs(authored);
    const footnoteNumbers = computeNoteDisplayNumbers(
      footnotes,
      footnoteRefs.map((ref) => ref.footnoteId),
    );
    const endnoteNumbers = computeNoteDisplayNumbers(
      endnotes,
      collectEndnoteRefs(authored).map((ref) => ref.endnoteId),
    );
    const blocks = remapNoteMarkerText(authored, { footnoteNumbers, endnoteNumbers });

    const measures = measureBlocks(blocks, contentWidth);

    const footnoteContentById =
      footnoteRefs.length === 0
        ? undefined
        : buildFootnoteContentMap(footnotes, footnoteRefs, contentWidth, {
            ...storyOptions,
            measureBlocks,
          });
    // The paginator reserves each note's band on the page its reference line
    // lands on, so it needs the heights before it places a line. The separator
    // slot is added once per note-bearing page by the paginator itself.
    const footnoteHeightById = new Map<number, number>();
    for (const [id, content] of footnoteContentById ?? []) {
      footnoteHeightById.set(id, content.height + FOOTNOTE_ENTRY_MARGIN_BOTTOM);
    }

    const sectionHeaderFooterRefs = resolveSectionHeaderFooterRefs(document);
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
      ...(footnoteHeightById.size === 0 ? {} : { footnoteHeightById }),
      ...(sectionHeaderFooterRefs === undefined ? {} : { sectionHeaderFooterRefs }),
    };
    return {
      layout: layoutDocument(blocks, measures, layoutOptions),
      blockLookup: buildBlockLookup(blocks, measures),
      footnoteContentById,
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

  const { layout, blockLookup, footnoteContentById } = laidOut.value;

  // Headers and footers are converted after pagination because a page-number
  // field measures against the final page count, and because nothing in a
  // header changes where the body's lines fell.
  const stories = Result.try(() => {
    const pageCount = layout.pages.length;
    const now = options.now ?? new Date();
    const shared = { contentWidth, storyOptions, pageCount, now } as const;
    return {
      headerContentByRId: convertStories({
        ...shared,
        parts: document.package.headers,
        metrics: { section: "header", pageSize, margins },
      }),
      footerContentByRId: convertStories({
        ...shared,
        parts: document.package.footers,
        metrics: { section: "footer", pageSize, margins },
      }),
    };
  });
  if (stories.isErr()) {
    return Result.err(
      new HeadlessLayoutError({
        message: "The header and footer stories could not be laid out.",
        cause: stories.error,
      }),
    );
  }

  const embeddedFonts = declaresEmbeddedFonts(document)
    ? await Result.tryPromise({
        try: async () => {
          const source = await toArrayBuffer(input);
          return extractEmbeddedFonts(source, embeddedFontNonce(source));
        },
        catch: (cause) =>
          new HeadlessLayoutError({
            message: "The package's embedded fonts could not be extracted.",
            cause,
          }),
      })
    : Result.ok<EmbeddedFont[], HeadlessLayoutError>([]);
  if (embeddedFonts.isErr()) {
    return Result.err(embeddedFonts.error);
  }

  const watermark = getDocumentWatermark(document);
  // `watermarkImageSrc` is one source for the whole document, so a package with
  // a different picture per header (a `w:titlePg` first page, or even/odd
  // headers) cannot be served by it: supplying it anyway would paint one
  // header's picture under another's watermark, silently. It is supplied only
  // when every picture watermark in the package resolves to the same source,
  // and withheld otherwise so the builder reports the picture it could not
  // paint. Serving them properly needs a source per relationship id.
  const watermarksByHeader = collectWatermarks(document.package.headers);
  const pictureSources = new Set(
    [watermark, ...watermarksByHeader.values()]
      .filter((candidate) => candidate !== undefined)
      .map((candidate) => resolveWatermarkImageSrc(document, candidate)),
  );
  const watermarkImageSrc = pictureSources.size === 1 ? [...pictureSources].at(0) : undefined;
  const pageBorders = firstSection?.pageBorders;

  return Result.ok({
    layout,
    blockLookup,
    document,
    proseDoc: projected.value,
    documentFeatures: readPaintFeatures(document),
    furniture: {
      // The painter takes one page-border spec and one theme per render, so the
      // first section's borders are the document's borders here too.
      ...(pageBorders === undefined ? {} : { pageBorders }),
      ...(document.package.theme === undefined ? {} : { theme: document.package.theme }),
      ...(watermark === undefined ? {} : { watermark }),
      watermarkByHeaderRId: watermarksByHeader,
      ...(watermarkImageSrc === undefined ? {} : { watermarkImageSrc }),
      headerContentByRId: stories.value.headerContentByRId,
      footerContentByRId: stories.value.footerContentByRId,
      titlePg: firstSection?.titlePg === true,
      ...(footnoteContentById === undefined
        ? {}
        : {
            footnoteContentById: footnoteContentById satisfies ReadonlyMap<number, FootnoteContent>,
          }),
    },
    embeddedFonts: embeddedFonts.value,
    unsupported: Object.values(UNPAGINATED_STORIES).filter((gap) => hasStory(document, gap.story)),
  });
};
