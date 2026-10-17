import type {
  FlowBlock,
  HeaderFooterContent,
  PageHeaderFooterRefs,
  PageMargins,
} from "../layout-engine/types";

type BodyMarginClearanceOptions = {
  authoredMargins: PageMargins;
  preparedHeader: HeaderFooterContent | undefined;
  preparedFooter: HeaderFooterContent | undefined;
};

/** Derive body margins that clear the visible header and footer stories. */
export const bodyMarginsClearHeaderFooter = ({
  authoredMargins,
  preparedHeader,
  preparedFooter,
}: BodyMarginClearanceOptions): PageMargins => {
  const headerBottom = preparedHeader
    ? (authoredMargins.header ?? 0) + (preparedHeader.marginPushBottom ?? preparedHeader.height)
    : authoredMargins.top;
  const footerClearance = preparedFooter
    ? (authoredMargins.footer ?? 0) + (preparedFooter.marginPushBottom ?? preparedFooter.height)
    : authoredMargins.bottom;
  const top = Math.max(authoredMargins.top, headerBottom);
  const clearedTop = Math.max(top, preparedHeader?.bodyTopClearance ?? 0);
  const bottom = Math.max(authoredMargins.bottom, footerClearance);
  if (clearedTop === authoredMargins.top && bottom === authoredMargins.bottom) {
    return authoredMargins;
  }
  return { ...authoredMargins, top: clearedTop, bottom };
};

/** The header/footer part set a page shows (ECMA-376 §17.10.6, §17.10.1). */
export type HeaderFooterVariant = "default" | "even" | "first";

const VARIANT_REFS = {
  default: { header: "headerDefault", footer: "footerDefault" },
  even: { header: "headerEven", footer: "footerEven" },
  first: { header: "headerFirst", footer: "footerFirst" },
} as const satisfies Record<
  HeaderFooterVariant,
  { header: keyof PageHeaderFooterRefs; footer: keyof PageHeaderFooterRefs }
>;

const sectionShowsVariant = (refs: PageHeaderFooterRefs, variant: HeaderFooterVariant): boolean => {
  switch (variant) {
    case "default":
      return true;
    case "even":
      return refs.evenAndOddHeaders === true;
    case "first":
      return refs.titlePg === true;
    default: {
      const unhandled: never = variant;
      return unhandled;
    }
  }
};

type SectionVariantMarginOptions = {
  sectionHeaderFooterRefs: PageHeaderFooterRefs[] | undefined;
  /** Authored margins per section, in section order. */
  sectionMargins: readonly PageMargins[];
  /** Margins for a section with no authored entry. */
  fallbackMargins: PageMargins;
  headerContentByRId: ReadonlyMap<string, HeaderFooterContent> | undefined;
  footerContentByRId: ReadonlyMap<string, HeaderFooterContent> | undefined;
};

/**
 * Per section, the body margins that clear this variant's header and footer;
 * `undefined` for a section that does not show the variant. The editor
 * pipeline and headless layout both derive variant margins here.
 */
export const sectionVariantMargins = (
  variant: HeaderFooterVariant,
  {
    sectionHeaderFooterRefs,
    sectionMargins,
    fallbackMargins,
    headerContentByRId,
    footerContentByRId,
  }: SectionVariantMarginOptions,
): (PageMargins | undefined)[] | undefined =>
  sectionHeaderFooterRefs?.map((refs, index) => {
    if (!sectionShowsVariant(refs, variant)) return undefined;
    const headerRId = refs[VARIANT_REFS[variant].header];
    const footerRId = refs[VARIANT_REFS[variant].footer];
    return bodyMarginsClearHeaderFooter({
      authoredMargins: sectionMargins[index] ?? fallbackMargins,
      preparedHeader: headerRId ? headerContentByRId?.get(headerRId) : undefined,
      preparedFooter: footerRId ? footerContentByRId?.get(footerRId) : undefined,
    });
  });

type SectionHeaderFooterClearanceOptions = {
  authoredMargins: PageMargins;
  sectionHeaderFooterRefs: PageHeaderFooterRefs[] | undefined;
  headerContentByRId: ReadonlyMap<string, HeaderFooterContent> | undefined;
  footerContentByRId: ReadonlyMap<string, HeaderFooterContent> | undefined;
};

/** Apply each section's default header/footer clearance to its section-break block. */
export const bodyBlocksClearSectionHeaderFooter = (
  blocks: FlowBlock[],
  {
    authoredMargins,
    sectionHeaderFooterRefs,
    headerContentByRId,
    footerContentByRId,
  }: SectionHeaderFooterClearanceOptions,
): FlowBlock[] => {
  if (!sectionHeaderFooterRefs || (!headerContentByRId && !footerContentByRId)) {
    return blocks;
  }

  let sectionIndex = 0;
  let currentAuthoredMargins = authoredMargins;
  let changed = false;
  const nextBlocks = blocks.map((block) => {
    if (block.kind !== "sectionBreak") {
      return block;
    }

    currentAuthoredMargins = block.margins ?? currentAuthoredMargins;
    const refs = sectionHeaderFooterRefs[sectionIndex];
    sectionIndex += 1;
    const effectiveMargins = bodyMarginsClearHeaderFooter({
      authoredMargins: currentAuthoredMargins,
      preparedHeader: refs?.headerDefault ? headerContentByRId?.get(refs.headerDefault) : undefined,
      preparedFooter: refs?.footerDefault ? footerContentByRId?.get(refs.footerDefault) : undefined,
    });
    if (effectiveMargins === currentAuthoredMargins && block.margins === currentAuthoredMargins) {
      return block;
    }

    changed = true;
    return { ...block, margins: effectiveMargins };
  });

  return changed ? nextBlocks : blocks;
};
