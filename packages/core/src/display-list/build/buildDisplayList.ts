/**
 * `Layout` + `BlockLookup` → `DisplayList`.
 *
 * The single source of truth for painting. Two backends consume the result and
 * neither may look behind it, so anything missing here is painted nowhere.
 *
 * Purity is a hard requirement, not a preference: no clock, no randomness, no
 * `globalThis.document`. The font and image tables intern in first-use page
 * order, and author colours are assigned per build, so two builds of one layout
 * are deeply equal. The one impure edge is the measure provider, which the
 * composition root installs and which is deterministic for a given font set;
 * that is the same seam line breaking already ran through, which is the point.
 */

import type { BlockLookup, BlockLookupEntry } from "../../layout-painter/index";
import type {
  Fragment,
  ImageRun,
  Layout,
  Page,
  ParagraphBlock,
  ParagraphBorders,
} from "../../layout-engine/types";
import type {
  DisplayColor,
  DisplayLink,
  DisplayLinkTarget,
  DisplayList,
  DisplayMetadata,
  DisplayOutlineEntry,
  DisplayPage,
  DisplayPrimitive,
} from "../types";
import { AuthorColorTable, type BuildContext } from "./buildContext";
import { DOC_CANVAS } from "./colors";
import { collectFloatingImages, pageGeometryOf } from "./floatingImages";
import { FontTable } from "./fontTable";
import { ImageTable, paintImageFragment } from "./imagePrimitives";
import {
  paintColumnSeparators,
  paintFootnoteArea,
  paintPageBackground,
  reportUnreachableFurniture,
} from "./pageFurniture";
import { paintParagraphFragment } from "./paragraphPrimitives";
import { paintTableFragment } from "./tablePrimitives";
import { paintTextBoxFragment } from "./textBoxPrimitives";
import { UnsupportedCollector, UNSUPPORTED_CONSTRUCT } from "./unsupported";

/** Document-wide gaps are attributed to the first page, the only one certain to exist. */
const FIRST_PAGE_INDEX = 0;

/**
 * Every feature a source document can have that a `Layout` does not record,
 * with the reason it paints nothing. The list is the single source of truth:
 * {@link DocumentFeatures} is derived from it, so a feature added here cannot
 * be left without a flag or without a reason.
 */
const DOCUMENT_FEATURE_GAPS = [
  [
    "pageBorders",
    "w:pgBorders reach the painter through render options, not through Layout, so the builder cannot paint them",
  ],
  [
    "watermark",
    "watermarks reach the painter through render options, not through Layout, so the builder cannot paint them",
  ],
] as const satisfies readonly (readonly [keyof typeof UNSUPPORTED_CONSTRUCT, string])[];

/**
 * Which of those features the source document actually has. The builder cannot
 * detect them, so the caller that parsed the document states it.
 */
export type DocumentFeatures = {
  readonly [Feature in (typeof DOCUMENT_FEATURE_GAPS)[number][0]]: boolean;
};

export type BuildDisplayListOptions = {
  readonly layout: Layout;
  readonly blockLookup: BlockLookup;
  readonly metadata?: DisplayMetadata;
  /** Page background. Defaults to opaque white: the PDF has no theme. */
  readonly pageBackground?: DisplayColor;
  /**
   * Omitted means unknown, and an unknown document is reported as having both:
   * a silent gap is worse than one the reader can dismiss.
   */
  readonly documentFeatures?: DocumentFeatures;
};

const paragraphTextOf = (block: ParagraphBlock): string =>
  block.runs
    .map((run) => (run.kind === "text" ? run.text : ""))
    .join("")
    .trim();

type FirstPass = {
  readonly bookmarkTargets: ReadonlyMap<string, DisplayLinkTarget>;
  readonly outline: readonly DisplayOutlineEntry[];
};

/**
 * Resolve everything that needs the whole document before any page is painted:
 * where each bookmark lands (a `#name` hyperlink cannot become a page target
 * until then) and the heading outline.
 */
const collectDocumentTargets = (layout: Layout, blockLookup: BlockLookup): FirstPass => {
  const bookmarkTargets = new Map<string, DisplayLinkTarget>();
  const outline: DisplayOutlineEntry[] = [];

  for (const [pageIndex, page] of layout.pages.entries()) {
    for (const fragment of page.fragments) {
      if (fragment.kind !== "paragraph") {
        continue;
      }
      const block = blockLookup.get(String(fragment.blockId))?.block;
      if (block?.kind !== "paragraph") {
        continue;
      }
      for (const name of block.bookmarks ?? []) {
        if (!bookmarkTargets.has(name)) {
          bookmarkTargets.set(name, { kind: "page", pageIndex, yPx: fragment.y });
        }
      }
      const level = block.attrs?.outlineLevel;
      if (level !== undefined && fragment.continuesFromPrev !== true) {
        const title = paragraphTextOf(block);
        if (title.length > 0) {
          outline.push({ title, level, pageIndex, yPx: fragment.y });
        }
      }
    }
  }

  return { bookmarkTargets, outline };
};

const bordersOfNeighbour = (entry: BlockLookupEntry | undefined): ParagraphBorders | undefined =>
  entry?.block.kind === "paragraph" ? entry.block.attrs?.borders : undefined;

type PaintFragmentOptions = {
  readonly fragment: Fragment;
  readonly entry: BlockLookupEntry;
  readonly context: BuildContext;
  readonly prevEntry: BlockLookupEntry | undefined;
  readonly nextEntry: BlockLookupEntry | undefined;
};

const paintFragment = ({
  fragment,
  entry,
  context,
  prevEntry,
  nextEntry,
}: PaintFragmentOptions): readonly DisplayPrimitive[] => {
  const { block, measure } = entry;
  const mismatch = (): readonly DisplayPrimitive[] => {
    context.unsupported.report(
      UNSUPPORTED_CONSTRUCT.measureMismatch,
      context.pageIndex,
      `${fragment.kind} fragment ${String(fragment.blockId)} has a ${block.kind} block and a ${measure.kind} measure`,
    );
    return [];
  };

  switch (fragment.kind) {
    case "paragraph": {
      if (block.kind !== "paragraph" || measure.kind !== "paragraph") {
        return mismatch();
      }
      const prevBorders = bordersOfNeighbour(prevEntry);
      const nextBorders = bordersOfNeighbour(nextEntry);
      return paintParagraphFragment({
        fragment,
        block,
        measure,
        context,
        ...(prevBorders === undefined ? {} : { prevBorders }),
        ...(nextBorders === undefined ? {} : { nextBorders }),
      });
    }
    case "table":
      return block.kind === "table" && measure.kind === "table"
        ? paintTableFragment({ fragment, block, measure, context })
        : mismatch();
    case "image":
      return block.kind === "image" ? paintImageFragment(fragment, block, context) : mismatch();
    case "textBox":
      return block.kind === "textBox" && measure.kind === "textBox"
        ? paintTextBoxFragment({ fragment, block, measure, context })
        : mismatch();
    default: {
      const unreachable: never = fragment;
      context.unsupported.report(
        UNSUPPORTED_CONSTRUCT.fragmentKind,
        context.pageIndex,
        `fragment kind ${JSON.stringify(unreachable)} has no builder`,
      );
      return [];
    }
  }
};

type BuildPageOptions = {
  readonly page: Page;
  readonly pageIndex: number;
  readonly totalPages: number;
  readonly blockLookup: BlockLookup;
  readonly pageBackground: DisplayColor;
  readonly bookmarkTargets: ReadonlyMap<string, DisplayLinkTarget>;
  readonly fonts: FontTable;
  readonly images: ImageTable;
  readonly unsupported: UnsupportedCollector;
  readonly authorColors: AuthorColorTable;
};

const buildPage = ({
  page,
  pageIndex,
  totalPages,
  blockLookup,
  pageBackground,
  bookmarkTargets,
  fonts,
  images,
  unsupported,
  authorColors,
}: BuildPageOptions): DisplayPage => {
  const links: DisplayLink[] = [];
  const context: BuildContext = {
    fonts,
    images,
    unsupported,
    authorColors,
    links,
    bookmarkTargets,
    pageIndex,
    pageNumber: page.logicalNumber,
    totalPages,
  };

  const entries = page.fragments.map((fragment) => blockLookup.get(String(fragment.blockId)));

  // Floating images are extracted before anything is painted, because
  // `behindDoc` artwork paints under the body and the rest paints over it: the
  // producer resolves that stacking so no backend has to.
  const geometry = pageGeometryOf(page);
  const placedFloats = new Set<ImageRun>();
  const behindFloats: DisplayPrimitive[] = [];
  const frontFloats: DisplayPrimitive[] = [];
  for (const [index, fragment] of page.fragments.entries()) {
    const block = entries[index]?.block;
    if (fragment.kind !== "paragraph" || block?.kind !== "paragraph") {
      continue;
    }
    for (const placement of collectFloatingImages({
      block,
      fragmentYPx: fragment.y,
      geometry,
      context,
      placed: placedFloats,
    })) {
      (placement.behindDoc ? behindFloats : frontFloats).push(...placement.primitives);
    }
  }

  // Back to front, as `renderPage.ts` appends: background, behind-document
  // floats, body fragments in `page.fragments` order, front floats, column
  // separators, then the footnote band.
  const primitives: DisplayPrimitive[] = [
    paintPageBackground(page, pageBackground),
    ...behindFloats,
  ];

  for (const [index, fragment] of page.fragments.entries()) {
    const entry = entries[index];
    if (!entry) {
      unsupported.report(
        UNSUPPORTED_CONSTRUCT.missingBlock,
        pageIndex,
        `no block lookup entry for ${fragment.kind} fragment ${String(fragment.blockId)}`,
      );
      continue;
    }
    primitives.push(
      ...paintFragment({
        fragment,
        entry,
        context,
        prevEntry: entries[index - 1],
        nextEntry: entries[index + 1],
      }),
    );
  }

  primitives.push(...frontFloats);
  primitives.push(...paintColumnSeparators(page));
  primitives.push(...paintFootnoteArea(page, context));
  reportUnreachableFurniture(page, context);

  return {
    pageNumber: page.number,
    widthPx: page.size.w,
    heightPx: page.size.h,
    orientation: page.orientation ?? (page.size.w > page.size.h ? "landscape" : "portrait"),
    primitives,
    links,
  };
};

export const buildDisplayList = ({
  layout,
  blockLookup,
  metadata,
  pageBackground,
  documentFeatures,
}: BuildDisplayListOptions): DisplayList => {
  const fonts = new FontTable();
  const images = new ImageTable();
  const unsupported = new UnsupportedCollector();
  const authorColors = new AuthorColorTable();
  const { bookmarkTargets, outline } = collectDocumentTargets(layout, blockLookup);

  // Page borders and watermarks reach the DOM painter through
  // `RenderPageOptions`; nothing in a `Layout` records either, so the gap is
  // stated once per document rather than per page. A `DisplayUnsupported`
  // carries a page index, so a document with no page has nothing to attribute
  // the gap to, and nothing was painted there to be missing from.
  if (layout.pages.length > 0) {
    for (const [feature, detail] of DOCUMENT_FEATURE_GAPS) {
      if (documentFeatures !== undefined && !documentFeatures[feature]) {
        continue;
      }
      unsupported.report(UNSUPPORTED_CONSTRUCT[feature], FIRST_PAGE_INDEX, detail);
    }
  }

  const pages = layout.pages.map((page, pageIndex) =>
    buildPage({
      page,
      pageIndex,
      totalPages: layout.pages.length,
      blockLookup,
      pageBackground: pageBackground ?? DOC_CANVAS,
      bookmarkTargets,
      fonts,
      images,
      unsupported,
      authorColors,
    }),
  );

  return {
    pages,
    fonts: fonts.snapshot(),
    images: images.snapshot(),
    outline,
    metadata: metadata ?? {},
    unsupported: unsupported.snapshot(),
  };
};
