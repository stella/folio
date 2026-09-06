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

import type { EmbeddedFont } from "../../fonts/embeddedFonts";
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
  DisplayHitRegionKind,
  DisplayLink,
  DisplayLinkTarget,
  DisplayList,
  DisplayMetadata,
  DisplayOutlineEntry,
  DisplayPage,
  DisplayPrimitive,
} from "../types";
import { HIT_REGION_KINDS } from "../primitives";
import { AuthorColorTable, type BuildContext } from "./buildContext";
import { DOC_CANVAS } from "./colors";
import { collectFloatingImages, pageGeometryOf } from "./floatingImages";
import { FontTable } from "./fontTable";
import { paintPageFurniture, type PageFurnitureInputs } from "./furniture";
import { ImageTable, paintImageFragment } from "./imagePrimitives";
import { paintColumnSeparators, paintPageBackground } from "./pageFurniture";
import { paintParagraphFragment } from "./paragraphPrimitives";
import { blockRegion, createPageComposer, type PageComposer } from "./regions";
import { paintTableFragment } from "./tablePrimitives";
import { paintTextBoxFragment } from "./textBoxPrimitives";
import { UnsupportedCollector, UNSUPPORTED_CONSTRUCT } from "./unsupported";

/** Document-wide gaps are attributed to the first page, the only one certain to exist. */
const FIRST_PAGE_INDEX = 0;

/**
 * Every feature a source document can have that a `Layout` does not record and
 * that nothing on a page betrays, with the reason it paints nothing when the
 * caller supplies neither the construct nor a flag denying it. The list is the
 * single source of truth: {@link DocumentFeatures} is derived from it, so a
 * feature added here cannot be left without a flag or without a reason.
 *
 * Footnote bodies and header/footer stories are deliberately absent: a page
 * that carries them says so itself (`footnoteIds`, `headerFooterRefs`), so
 * their gaps are reported per page from the layout rather than from a
 * document-wide flag the caller would have to remember to set.
 */
const DOCUMENT_FEATURE_GAPS = [
  [
    "pageBorders",
    "w:pgBorders reach the painter through render options, not through Layout, and none were supplied to the builder",
  ],
  [
    "watermark",
    "watermarks reach the painter through render options, not through Layout, and none was supplied to the builder",
  ],
] as const satisfies readonly (readonly [keyof typeof UNSUPPORTED_CONSTRUCT, string])[];

/**
 * Which of those features the source document actually has. The builder cannot
 * detect them, so the caller that parsed the document states it.
 */
export type DocumentFeatures = {
  readonly [Feature in (typeof DOCUMENT_FEATURE_GAPS)[number][0]]: boolean;
};

export type BuildDisplayListOptions = PageFurnitureInputs & {
  readonly layout: Layout;
  readonly blockLookup: BlockLookup;
  readonly metadata?: DisplayMetadata;
  /** Page background. Defaults to opaque white: the PDF has no theme. */
  readonly pageBackground?: DisplayColor;
  /**
   * States, for each construct in {@link DOCUMENT_FEATURE_GAPS}, whether the
   * document has one at all. Omitted means unknown, and an unknown document is
   * reported as having both: a silent gap is worse than one the reader can
   * dismiss. A construct supplied through {@link PageFurnitureInputs} is
   * painted and never reported, whatever this says.
   */
  readonly documentFeatures?: DocumentFeatures;
  /**
   * The package's own font faces (`fonts/embeddedFonts.ts`). Absent: a face the
   * measurer resolved to an embedded family travels as a name, and a backend
   * that cannot find that name on the host paints the document in a substitute.
   */
  readonly embeddedFonts?: readonly EmbeddedFont[];
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
  readonly composer: PageComposer;
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
  composer,
  prevEntry,
  nextEntry,
}: PaintFragmentOptions): void => {
  const { block, measure } = entry;
  const mismatch = (): void => {
    context.unsupported.report(
      UNSUPPORTED_CONSTRUCT.measureMismatch,
      context.pageIndex,
      `${fragment.kind} fragment ${String(fragment.blockId)} has a ${block.kind} block and a ${measure.kind} measure`,
    );
  };

  switch (fragment.kind) {
    case "paragraph": {
      if (block.kind !== "paragraph" || measure.kind !== "paragraph") {
        mismatch();
        return;
      }
      const prevBorders = bordersOfNeighbour(prevEntry);
      const nextBorders = bordersOfNeighbour(nextEntry);
      paintParagraphFragment({
        fragment,
        block,
        measure,
        context,
        composer,
        ...(prevBorders === undefined ? {} : { prevBorders }),
        ...(nextBorders === undefined ? {} : { nextBorders }),
      });
      return;
    }
    case "table":
      if (block.kind === "table" && measure.kind === "table") {
        paintTableFragment({ fragment, block, measure, context, composer });
        return;
      }
      mismatch();
      return;
    case "image":
      if (block.kind === "image") {
        composer.push(paintImageFragment(fragment, block, context));
        return;
      }
      mismatch();
      return;
    case "textBox":
      if (block.kind === "textBox" && measure.kind === "textBox") {
        paintTextBoxFragment({ composer, fragment, block, measure, context });
        return;
      }
      mismatch();
      return;
    default: {
      const unreachable: never = fragment;
      context.unsupported.report(
        UNSUPPORTED_CONSTRUCT.fragmentKind,
        context.pageIndex,
        `fragment kind ${JSON.stringify(unreachable)} has no builder`,
      );
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
  readonly furniture: PageFurnitureInputs;
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
  furniture,
}: BuildPageOptions): DisplayPage => {
  const links: DisplayLink[] = [];
  const context: BuildContext = {
    fonts,
    images,
    unsupported,
    authorColors,
    links,
    bookmarkTargets,
    story: { kind: "body" },
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

  const { behind, above } = paintPageFurniture({ page, furniture, context });

  // Back to front, as `renderPage.ts` appends: background, a `zOrder="back"`
  // page border, the watermark, behind-document floats, body fragments in
  // `page.fragments` order, front floats, column separators, the footnote band,
  // the header, the footer, and last a `zOrder="front"` page border.
  const composer = createPageComposer();
  composer.push([paintPageBackground(page, pageBackground), ...behind, ...behindFloats]);

  // The body's own box. A click outside it is in a margin, where the caret
  // belongs to whichever line is nearest rather than to the page.
  composer.region(
    {
      kind: HIT_REGION_KINDS.pageContent,
      rect: {
        xPx: page.margins.left,
        yPx: page.margins.top,
        widthPx: page.size.w - page.margins.left - page.margins.right,
        heightPx: page.size.h - page.margins.top - page.margins.bottom,
      },
    },
    () => {
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
        composer.region(
          blockRegion({ fragment, kind: FRAGMENT_REGION_KIND[fragment.kind], context }),
          () => {
            paintFragment({
              fragment,
              entry,
              context,
              composer,
              prevEntry: entries[index - 1],
              nextEntry: entries[index + 1],
            });
          },
        );
      }
    },
  );

  composer.push(frontFloats);
  composer.push(paintColumnSeparators(page));
  above(composer);

  return {
    pageNumber: page.number,
    widthPx: page.size.w,
    heightPx: page.size.h,
    orientation: page.orientation ?? (page.size.w > page.size.h ? "landscape" : "portrait"),
    primitives: composer.primitives(),
    regions: composer.regions(),
    links,
  };
};

/** Which kind of region a fragment paints into. */
const FRAGMENT_REGION_KIND = {
  paragraph: HIT_REGION_KINDS.paragraph,
  table: HIT_REGION_KINDS.table,
  image: HIT_REGION_KINDS.image,
  textBox: HIT_REGION_KINDS.textBox,
} as const satisfies Record<Fragment["kind"], DisplayHitRegionKind>;

/** Whether the caller handed over the construct this document-wide gap names. */
const SUPPLIED_BY = {
  pageBorders: (furniture: PageFurnitureInputs) => furniture.pageBorders !== undefined,
  watermark: (furniture: PageFurnitureInputs) =>
    furniture.watermark !== undefined || furniture.watermarkByHeaderRId !== undefined,
} as const satisfies Record<
  (typeof DOCUMENT_FEATURE_GAPS)[number][0],
  (furniture: PageFurnitureInputs) => boolean
>;

/**
 * A document's display list, one page at a time.
 *
 * The tables a page refers to (faces, images), the bookmark targets a link can
 * resolve against, and the outline are properties of the whole document, so
 * they are computed once; a page's primitives are not, and an editor paints a
 * window of pages rather than all of them. Building every page to paint three
 * of them is the difference between work proportional to the screen and work
 * proportional to the document, which on a long one is the difference a reader
 * feels on every keystroke.
 *
 * A page is built at most once. `snapshot` returns what has been built so far,
 * so a caller that wants the whole list asks for every page first — which is
 * what {@link buildDisplayList} does.
 */
export type DisplayListBuilder = {
  readonly pageCount: number;
  /** The page at that index, built on first request and kept. */
  readonly pageAt: (pageIndex: number) => DisplayPage | undefined;
  /**
   * The list as far as it has been built. The tables grow as pages are built,
   * so a backend reads them after the page it is about to paint, never before.
   */
  readonly snapshot: () => DisplayList;
};

export const createDisplayListBuilder = ({
  layout,
  blockLookup,
  metadata,
  pageBackground,
  documentFeatures,
  embeddedFonts,
  ...furniture
}: BuildDisplayListOptions): DisplayListBuilder => {
  const fonts = new FontTable(embeddedFonts ?? []);
  const images = new ImageTable();
  const unsupported = new UnsupportedCollector();
  const authorColors = new AuthorColorTable();
  const { bookmarkTargets, outline } = collectDocumentTargets(layout, blockLookup);

  // Nothing on a page betrays a page border or a watermark the caller withheld,
  // so the gap is stated once per document rather than per page. A
  // `DisplayUnsupported` carries a page index, so a document with no page has
  // nothing to attribute the gap to, and nothing was painted there to be
  // missing from.
  if (layout.pages.length > 0) {
    for (const [feature, detail] of DOCUMENT_FEATURE_GAPS) {
      if (SUPPLIED_BY[feature](furniture)) {
        continue;
      }
      if (documentFeatures !== undefined && !documentFeatures[feature]) {
        continue;
      }
      unsupported.report(UNSUPPORTED_CONSTRUCT[feature], FIRST_PAGE_INDEX, detail);
    }
  }

  const built = new Map<number, DisplayPage>();

  const pageAt = (pageIndex: number): DisplayPage | undefined => {
    const existing = built.get(pageIndex);
    if (existing !== undefined) {
      return existing;
    }
    const page = layout.pages[pageIndex];
    if (page === undefined) {
      return undefined;
    }
    const display = buildPage({
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
      furniture,
    });
    built.set(pageIndex, display);
    return display;
  };

  return {
    pageCount: layout.pages.length,
    pageAt,
    snapshot: () => ({
      // In layout order whichever order they were built in: a page's index is
      // its place in the document, not its place in a paint sequence.
      pages: [...built.entries()].sort(([left], [right]) => left - right).map(([, page]) => page),
      fonts: fonts.snapshot(),
      images: images.snapshot(),
      outline,
      metadata: metadata ?? {},
      unsupported: unsupported.snapshot(),
    }),
  };
};

/**
 * The whole document's display list. Every page is built, so the tables, the
 * outline and the gap report are complete; an editor that paints a window of
 * pages uses {@link createDisplayListBuilder} instead.
 */
export const buildDisplayList = (options: BuildDisplayListOptions): DisplayList => {
  const builder = createDisplayListBuilder(options);
  for (let pageIndex = 0; pageIndex < builder.pageCount; pageIndex += 1) {
    builder.pageAt(pageIndex);
  }
  return builder.snapshot();
};
