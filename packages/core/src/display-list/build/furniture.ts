/**
 * The four constructs a `Layout` does not carry, and how they reach a page.
 *
 * Page borders, the watermark, the footnote bodies and the header/footer
 * stories are all inputs the caller already holds: the DOM painter takes them
 * from `RenderPageOptions`, and the builder takes the same values in the same
 * shapes. Sharing the shapes is the point — a per-page selection made twice is
 * a selection that can come out two ways, so the painter's own
 * `applySectionHeaderFooterOptions` decides which header, footer and watermark
 * this page gets, and this module only paints what it chose.
 */

import type { FootnoteContent, HeaderFooterContent, Page } from "../../layout-engine/types";
import {
  applySectionHeaderFooterOptions,
  selectSectionHeaderFooterRIds,
  type RenderPageOptions,
} from "../../layout-painter/renderPage";
import type { Theme, Watermark } from "../../types/document";
import type { DisplayPrimitive } from "../types";
import type { BuildContext } from "./buildContext";
import { paintHeaderFooter } from "./headerFooterPrimitives";
import { paintPageBorders, type PageBorders } from "./pageBorderPrimitives";
import { paintFootnoteArea } from "./pageFurniture";
import type { PageComposer } from "./regions";
import { UNSUPPORTED_CONSTRUCT } from "./unsupported";
import { paintWatermark } from "./watermarkPrimitives";

/**
 * Everything the builder cannot derive from a `Layout`. Every field is
 * optional, and every one of them means the same thing when absent: the
 * construct is not painted, and it is reported wherever the layout or the
 * caller's `documentFeatures` shows the document had one.
 */
export type PageFurnitureInputs = {
  /**
   * `w:pgBorders` for the section. Absent: no border is painted, and the gap is
   * reported unless `documentFeatures` says the document has none.
   */
  readonly pageBorders?: PageBorders;
  /** Resolves `w:themeColor` on a page border. Absent: authored colours only. */
  readonly theme?: Theme | null;
  /**
   * The watermark for every page with no more specific match in
   * {@link watermarkByHeaderRId}. Absent with that map also absent: no
   * watermark is painted, and the gap is reported unless `documentFeatures`
   * says the document has none.
   */
  readonly watermark?: Watermark;
  /**
   * Per-header-rId watermarks, authoritative when present: a page whose active
   * header carries none paints none, rather than inheriting {@link watermark}.
   */
  readonly watermarkByHeaderRId?: ReadonlyMap<string, Watermark>;
  /**
   * Resolved `data:` source for a picture watermark. Absent: the picture is
   * reported, because the relationship id → asset mapping belongs to the
   * package layer and the builder does no I/O.
   */
  readonly watermarkImageSrc?: string;
  /**
   * The header for every page, or for pages 2+ when `titlePg` is set. Absent
   * with the per-rId map also absent: no header is painted, and a page whose
   * `headerFooterRefs` name one says so through `unsupported`.
   */
  readonly headerContent?: HeaderFooterContent;
  /** The footer counterpart of {@link headerContent}. */
  readonly footerContent?: HeaderFooterContent;
  /** Headers by part relationship id, for section-scoped and even/odd headers. */
  readonly headerContentByRId?: ReadonlyMap<string, HeaderFooterContent>;
  /** Footers by part relationship id. */
  readonly footerContentByRId?: ReadonlyMap<string, HeaderFooterContent>;
  /** The first page's header when `titlePg` is set and the layout carries no section refs. */
  readonly firstPageHeaderContent?: HeaderFooterContent;
  /** The first page's footer when `titlePg` is set and the layout carries no section refs. */
  readonly firstPageFooterContent?: HeaderFooterContent;
  /** `w:titlePg` for the first section. Absent: the first page is not special-cased. */
  readonly titlePg?: boolean;
  /** Overrides `w:pgMar`'s `w:header`. Absent: the page's own margin is used. */
  readonly headerDistancePx?: number;
  /** Overrides `w:pgMar`'s `w:footer`. Absent: the page's own margin is used. */
  readonly footerDistancePx?: number;
  /**
   * Footnote bodies by `w:footnote` id. Absent, or missing an id a page
   * carries: the band still opens with its rule and reserves the paginator's
   * height, and the missing bodies are reported.
   */
  readonly footnoteContentById?: ReadonlyMap<number, FootnoteContent>;
};

const suppliesHeaderFooter = (furniture: PageFurnitureInputs): boolean =>
  furniture.headerContent !== undefined ||
  furniture.footerContent !== undefined ||
  furniture.headerContentByRId !== undefined ||
  furniture.footerContentByRId !== undefined ||
  furniture.firstPageHeaderContent !== undefined ||
  furniture.firstPageFooterContent !== undefined;

const suppliesWatermark = (furniture: PageFurnitureInputs): boolean =>
  furniture.watermark !== undefined || furniture.watermarkByHeaderRId !== undefined;

/**
 * The header, footer and watermark this page paints, chosen by the painter's
 * own rule (`w:titlePg` first-page swap, `w:evenAndOddHeaders`, section-scoped
 * parts) rather than by a second implementation of it.
 */
const selectForPage = (page: Page, furniture: PageFurnitureInputs): RenderPageOptions => {
  const source: RenderPageOptions = {
    ...(furniture.headerContent === undefined ? {} : { headerContent: furniture.headerContent }),
    ...(furniture.footerContent === undefined ? {} : { footerContent: furniture.footerContent }),
    ...(furniture.headerContentByRId === undefined
      ? {}
      : { headerContentByRId: furniture.headerContentByRId }),
    ...(furniture.footerContentByRId === undefined
      ? {}
      : { footerContentByRId: furniture.footerContentByRId }),
    ...(furniture.watermark === undefined ? {} : { watermark: furniture.watermark }),
    ...(furniture.watermarkByHeaderRId === undefined
      ? {}
      : { watermarkByHeaderRId: furniture.watermarkByHeaderRId }),
  };

  const selected: RenderPageOptions = { ...source };
  if (applySectionHeaderFooterOptions(page, selected, source)) {
    return selected;
  }
  if (furniture.titlePg !== true || page.number !== 1) {
    return selected;
  }
  if (furniture.firstPageHeaderContent === undefined) {
    delete selected.headerContent;
  } else {
    selected.headerContent = furniture.firstPageHeaderContent;
  }
  if (furniture.firstPageFooterContent === undefined) {
    delete selected.footerContent;
  } else {
    selected.footerContent = furniture.firstPageFooterContent;
  }
  return selected;
};

/**
 * Name a header or footer the page asks for and the caller did not supply. A
 * backend cannot tell a page with no header from a page whose header never
 * reached the producer.
 */
const reportMissingStories = (
  page: Page,
  furniture: PageFurnitureInputs,
  selected: RenderPageOptions,
  context: BuildContext,
): void => {
  const refs = page.headerFooterRefs;
  if (refs === undefined) {
    return;
  }
  if (!suppliesHeaderFooter(furniture)) {
    context.unsupported.report(
      UNSUPPORTED_CONSTRUCT.headerFooterContent,
      context.pageIndex,
      "the page names header and footer parts, but no story content was supplied to the builder",
    );
    return;
  }
  // Only a part the page actually *selects* can be missing. A page that names
  // a first-page header and is not the first page selects none, and reporting
  // that as an absent story would be a gap the document does not have.
  const { headerRId, footerRId } = selectSectionHeaderFooterRIds(page);
  const namesHeader = headerRId !== undefined;
  const namesFooter = footerRId !== undefined;
  // A watermark from that same part proves the part reached the producer and
  // was read. Word's own watermark lives in a header that holds nothing else,
  // so a header with no story content is that document's normal shape rather
  // than a story that went missing on the way here.
  const headerCarriedAWatermark =
    headerRId !== undefined && furniture.watermarkByHeaderRId?.has(headerRId) === true;
  if (namesHeader && !headerCarriedAWatermark && selected.headerContent === undefined) {
    context.unsupported.report(
      UNSUPPORTED_CONSTRUCT.headerFooterContent,
      context.pageIndex,
      "the header part this page selects was not among the supplied stories",
    );
  }
  if (namesFooter && selected.footerContent === undefined) {
    context.unsupported.report(
      UNSUPPORTED_CONSTRUCT.headerFooterContent,
      context.pageIndex,
      "the footer part this page selects was not among the supplied stories",
    );
  }
};

/**
 * Furniture split by where it sits in the page's paint order: `behind` goes
 * under the body, `above` over it. `renderPage.ts` establishes both.
 */
export type PageFurniturePrimitives = {
  /** Painted under the body: a back page border, the watermark. */
  readonly behind: readonly DisplayPrimitive[];
  /** Painted over it, into the composer, so its regions nest. */
  readonly above: (composer: PageComposer) => void;
};

export type PaintPageFurnitureOptions = {
  readonly page: Page;
  readonly furniture: PageFurnitureInputs;
  readonly context: BuildContext;
};

export const paintPageFurniture = ({
  page,
  furniture,
  context,
}: PaintPageFurnitureOptions): PageFurniturePrimitives => {
  const selected = selectForPage(page, furniture);
  reportMissingStories(page, furniture, selected, context);

  const borders =
    furniture.pageBorders === undefined
      ? []
      : paintPageBorders({
          page,
          borders: furniture.pageBorders,
          ...(furniture.theme === undefined ? {} : { theme: furniture.theme }),
          context,
        });
  const paintsBorderBehind = furniture.pageBorders?.zOrder === "back";

  const watermark =
    !suppliesWatermark(furniture) || selected.watermark === undefined
      ? []
      : paintWatermark({
          page,
          watermark: selected.watermark,
          ...(furniture.watermarkImageSrc === undefined
            ? {}
            : { imageSrc: furniture.watermarkImageSrc }),
          context,
        });

  // What goes over the body is painted after it, through the composer, because
  // the note band and the header and footer slots are stories a click can land
  // in and their regions have to nest around what fills them.
  const above = (composer: PageComposer): void => {
    paintFootnoteArea({
      composer,
      page,
      context,
      ...(furniture.footnoteContentById === undefined
        ? {}
        : { contentById: furniture.footnoteContentById }),
    });
    if (selected.headerContent !== undefined) {
      paintHeaderFooter({
        composer,
        page,
        section: "header",
        content: selected.headerContent,
        ...(furniture.headerDistancePx === undefined
          ? {}
          : { distancePx: furniture.headerDistancePx }),
        context,
      });
    }
    if (selected.footerContent !== undefined) {
      paintHeaderFooter({
        composer,
        page,
        section: "footer",
        content: selected.footerContent,
        ...(furniture.footerDistancePx === undefined
          ? {}
          : { distancePx: furniture.footerDistancePx }),
        context,
      });
    }
    if (!paintsBorderBehind) {
      composer.push(borders);
    }
  };

  return { behind: [...(paintsBorderBehind ? borders : []), ...watermark], above };
};
