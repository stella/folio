/**
 * Header/footer document-model operations, extracted from the React
 * `useHeaderFooterEditor` hook. Pure functions over the folio `Document`
 * model: resolving which header/footer renders for each slot, and building the
 * new `Document` produced by creating, saving, or removing a header/footer.
 *
 * ProseMirror stays in the adapter: the save path serialises the live hidden
 * header/footer PM doc to `BlockContent[]` (`proseDocToBlocks`) and passes the
 * blocks in here.
 */

import type {
  BlockContent,
  Document,
  DocumentBody,
  DocxPackage,
  HeaderFooter,
  SectionProperties,
} from "../types/document";
import { clearHeaderFooterVerbatimXml } from "../docx/headerFooterVerbatim";

export type HeaderFooterPosition = "header" | "footer";

/**
 * The resolved header/footer slots for the active document, plus the
 * relationship ids that save/remove must target (the *displayed* slot's rId,
 * not necessarily `finalSectionProperties`).
 */
export type HeaderFooterResolution = {
  headerContent: HeaderFooter | null;
  footerContent: HeaderFooter | null;
  firstPageHeaderContent: HeaderFooter | null;
  firstPageFooterContent: HeaderFooter | null;
  hasTitlePg: boolean;
  activeHeaderRId: string | null;
  activeFooterRId: string | null;
  activeFirstHeaderRId: string | null;
  activeFirstFooterRId: string | null;
};

const EMPTY_RESOLUTION: HeaderFooterResolution = {
  headerContent: null,
  footerContent: null,
  firstPageHeaderContent: null,
  firstPageFooterContent: null,
  hasTitlePg: false,
  activeHeaderRId: null,
  activeFooterRId: null,
  activeFirstHeaderRId: null,
  activeFirstFooterRId: null,
};

/**
 * Resolve the header/footer content (and the active relationship ids) for a
 * parsed package. Returns the empty resolution when no package is loaded.
 */
export const resolveHeaderFooterContent = (
  pkg: DocxPackage | undefined,
): HeaderFooterResolution => {
  if (!pkg) {
    return EMPTY_RESOLUTION;
  }

  const finalProps = pkg.document.finalSectionProperties;
  const sections = pkg.document.sections;
  const headers = pkg.headers;
  const footers = pkg.footers;

  // Collect all section properties (inline sections + final).
  const allSectionProps: SectionProperties[] = [];
  if (sections) {
    for (const s of sections) {
      allSectionProps.push(s.properties);
    }
  } else if (finalProps) {
    allSectionProps.push(finalProps);
  }

  let header: HeaderFooter | null = null;
  let footer: HeaderFooter | null = null;
  let firstHeader: HeaderFooter | null = null;
  let firstFooter: HeaderFooter | null = null;
  let titlePg = false;

  // Resolve default headers/footers: use the last section that defines them
  // (typically the final section properties).
  // Same title-page-section preference as footers below — pages 2+ should use
  // the body section's default header, not a later signature section's
  // empty/different one.
  let resolvedHeaderRId: string | null = null;
  let resolvedFooterRId: string | null = null;
  let resolvedFirstHeaderRId: string | null = null;
  let resolvedFirstFooterRId: string | null = null;

  if (headers) {
    let primaryHeaderFromTitleSection: HeaderFooter | null = null;
    let primaryHeaderRId: string | null = null;
    let lastHeader: HeaderFooter | null = null;
    let lastHeaderRId: string | null = null;
    for (const sp of allSectionProps) {
      if (!sp.headerReferences) {
        continue;
      }
      const defaultRef = sp.headerReferences.find((r) => r.type === "default");
      if (!defaultRef?.rId) {
        continue;
      }
      const candidate = headers.get(defaultRef.rId);
      if (!candidate) {
        continue;
      }
      lastHeader = candidate;
      lastHeaderRId = defaultRef.rId;
      // Word only honors `first` references when `<w:titlePg/>` is set on the
      // section (ECMA-376 §17.10.6). A stale `first` reference on a section
      // without titlePg should be ignored — gating on `sp.titlePg` keeps this
      // resolution consistent with the first-page resolution below and with
      // Word's behavior.
      const hasFirstRef =
        sp.titlePg === true && sp.headerReferences.some((r) => r.type === "first");
      if (hasFirstRef && !primaryHeaderFromTitleSection) {
        primaryHeaderFromTitleSection = candidate;
        primaryHeaderRId = defaultRef.rId;
      }
    }
    header = primaryHeaderFromTitleSection ?? lastHeader ?? header;
    resolvedHeaderRId = primaryHeaderRId ?? lastHeaderRId;
  }

  if (footers) {
    // Per ECMA-376 §17.10, each section has its own header/footer references.
    // Folio's HF model is currently flat (one default per document) — the
    // closest spec-conformant approximation is "the section that hosts the
    // title page's first-page references". That's section 1 in NVCA-style
    // multi-section docs (sec 1: title page + body, sec 2..N: signature pages
    // with different / empty footers). Picking the LAST section's default (the
    // previous behavior) silently dropped the body footer's PAGE field on
    // pages 2+ when later sections override `default` with a stripped-down
    // footer.
    //
    // Algorithm: pick the default from the FIRST section that has both a
    // first-page reference AND a default — that's the title-page section,
    // whose default applies to pages 2+ within that section (and, since
    // folio's HF is flat, to pages 2+ globally). Fall back to the last default
    // if no section has both.
    let primaryFooterFromTitleSection: HeaderFooter | null = null;
    let primaryFooterRId: string | null = null;
    let lastFooter: HeaderFooter | null = null;
    let lastFooterRId: string | null = null;
    for (const sp of allSectionProps) {
      if (!sp.footerReferences) {
        continue;
      }
      const defaultRef = sp.footerReferences.find((r) => r.type === "default");
      if (!defaultRef?.rId) {
        continue;
      }
      const candidate = footers.get(defaultRef.rId);
      if (!candidate) {
        continue;
      }
      lastFooter = candidate;
      lastFooterRId = defaultRef.rId;
      // Same titlePg gate as headers above — ignore `first` refs in sections
      // that don't enable title-page mode.
      const hasFirstRef =
        sp.titlePg === true && sp.footerReferences.some((r) => r.type === "first");
      if (hasFirstRef && !primaryFooterFromTitleSection) {
        primaryFooterFromTitleSection = candidate;
        primaryFooterRId = defaultRef.rId;
      }
    }
    footer = primaryFooterFromTitleSection ?? lastFooter ?? footer;
    resolvedFooterRId = primaryFooterRId ?? lastFooterRId;
  }

  // Resolve first-page headers/footers: find the first section with titlePg.
  for (const sp of allSectionProps) {
    if (sp.titlePg) {
      titlePg = true;
      if (headers && sp.headerReferences) {
        const firstRef = sp.headerReferences.find((r) => r.type === "first");
        if (firstRef?.rId) {
          firstHeader = headers.get(firstRef.rId) ?? null;
          resolvedFirstHeaderRId = firstRef.rId;
        }
      }
      if (footers && sp.footerReferences) {
        const firstRef = sp.footerReferences.find((r) => r.type === "first");
        if (firstRef?.rId) {
          firstFooter = footers.get(firstRef.rId) ?? null;
          resolvedFirstFooterRId = firstRef.rId;
        }
      }
      break; // first section with titlePg wins
    }
  }

  // Fallback: if no section has titlePg, check finalSectionProperties for
  // first-page refs (they won't be used as first-page without titlePg, but
  // keep them so the "only first headers exist" fallback below works).
  if (!titlePg && headers) {
    const refs = finalProps?.headerReferences;
    const firstRef = refs?.find((r) => r.type === "first");
    if (firstRef?.rId) {
      firstHeader = headers.get(firstRef.rId) ?? null;
      resolvedFirstHeaderRId = firstRef.rId;
    }
  }
  if (!titlePg && footers) {
    const refs = finalProps?.footerReferences;
    const firstRef = refs?.find((r) => r.type === "first");
    if (firstRef?.rId) {
      firstFooter = footers.get(firstRef.rId) ?? null;
      resolvedFirstFooterRId = firstRef.rId;
    }
  }

  // When titlePg is not set but only 'first' headers exist, use them as
  // default. Mirror the rId fallback so save/remove targets the rId actually
  // rendered — otherwise the active default rId stays null and edits to the
  // displayed header/footer silently no-op.
  if (!titlePg) {
    if (!header && firstHeader) {
      header = firstHeader;
      resolvedHeaderRId = resolvedFirstHeaderRId;
    }
    if (!footer && firstFooter) {
      footer = firstFooter;
      resolvedFooterRId = resolvedFirstFooterRId;
    }
  }

  return {
    headerContent: header,
    footerContent: footer,
    firstPageHeaderContent: firstHeader,
    firstPageFooterContent: firstFooter,
    hasTitlePg: titlePg,
    // Active rIds for the *displayed* H/F. Save/remove must target these — not
    // finalSectionProperties — otherwise edits to a multi-section doc's
    // title-page footer end up in the (hidden) final section's rId and the
    // visible footer never updates.
    activeHeaderRId: resolvedHeaderRId,
    activeFooterRId: resolvedFooterRId,
    activeFirstHeaderRId: resolvedFirstHeaderRId,
    activeFirstFooterRId: resolvedFirstFooterRId,
  };
};

/**
 * The section properties that drive the rendered chrome, with `titlePg` merged
 * in when title-page mode was inferred from header/footer references.
 */
export const resolveEffectiveSectionProperties = (
  documentBody: DocumentBody | undefined,
  hasTitlePg: boolean,
): SectionProperties | undefined => {
  const firstContentSection = documentBody?.sections?.find((section) => section.content.length > 0);
  const base = firstContentSection?.properties ?? documentBody?.finalSectionProperties;
  if (!hasTitlePg || base?.titlePg) {
    return base;
  }
  return base ? { ...base, titlePg: true } : base;
};

/** Pick the relationship id for the slot a save/remove targets. */
export const pickActiveHeaderFooterRId = (
  resolution: HeaderFooterResolution,
  position: HeaderFooterPosition,
  isFirstPage: boolean,
): string | null => {
  if (isFirstPage) {
    return position === "header" ? resolution.activeFirstHeaderRId : resolution.activeFirstFooterRId;
  }
  return position === "header" ? resolution.activeHeaderRId : resolution.activeFooterRId;
};

/**
 * Build a Document with a fresh empty header/footer for the given slot, wired
 * into `finalSectionProperties` via a new reference. Returns null when the
 * document has no final section properties to attach the reference to.
 */
export const createEmptyHeaderFooter = (
  document: Document,
  position: HeaderFooterPosition,
  isFirstPage: boolean,
): Document | null => {
  const pkg = document.package;
  const sectionProps = pkg.document.finalSectionProperties;
  if (!sectionProps) {
    return null;
  }

  const hdrFtrType: "default" | "first" = isFirstPage ? "first" : "default";
  const rId = `rId_new_${position}_${hdrFtrType}`;
  const emptyHf: HeaderFooter = {
    type: position,
    hdrFtrType,
    content: [{ type: "paragraph", content: [] }],
  };

  const mapKey = position === "header" ? "headers" : "footers";
  const newMap = new Map(pkg[mapKey]);
  newMap.set(rId, emptyHf);

  const refKey = position === "header" ? "headerReferences" : "footerReferences";
  const existingRefs = sectionProps[refKey] ?? [];
  const newRef = { type: hdrFtrType, rId };

  return {
    ...document,
    package: {
      ...pkg,
      [mapKey]: newMap,
      document: {
        ...pkg.document,
        finalSectionProperties: {
          ...sectionProps,
          [refKey]: [...existingRefs, newRef],
        },
      },
    },
  };
};

export type SaveHeaderFooterArgs = {
  document: Document;
  position: HeaderFooterPosition;
  isFirstPage: boolean;
  /** The relationship id the displayed slot resolves to (`pickActiveHeaderFooterRId`). */
  activeRId: string;
  /** The serialised header/footer body (from the live PM doc). */
  blocks: BlockContent[];
};

/**
 * Serialise edited header/footer content into a new Document. A brand-new
 * HeaderFooter inside a brand-new Map keeps every earlier history entry's
 * Document untouched so undo can step back to the pre-edit state. Returns null
 * when the target header/footer cannot be found.
 */
export const saveHeaderFooterContent = ({
  document,
  position,
  isFirstPage,
  activeRId,
  blocks,
}: SaveHeaderFooterArgs): Document | null => {
  const pkg = document.package;
  const mapKey = position === "header" ? "headers" : "footers";
  const map = pkg[mapKey];
  const existing = map?.get(activeRId);
  if (!map || !existing) {
    return null;
  }

  const updated: HeaderFooter = {
    ...existing,
    type: position,
    hdrFtrType: isFirstPage ? "first" : "default",
    content: blocks,
  };
  clearHeaderFooterVerbatimXml(updated);
  const newMap = new Map(map);
  newMap.set(activeRId, updated);

  return {
    ...document,
    package: {
      ...pkg,
      [mapKey]: newMap,
    },
  };
};

export type RemoveHeaderFooterArgs = {
  document: Document;
  position: HeaderFooterPosition;
  /** The relationship id the displayed slot resolves to (`pickActiveHeaderFooterRId`). */
  activeRId: string;
};

/**
 * Remove a header/footer and drop its reference from every section that points
 * at it, returning the new Document.
 */
export const removeHeaderFooter = ({
  document,
  position,
  activeRId,
}: RemoveHeaderFooterArgs): Document => {
  const pkg = document.package;
  const refKey = position === "header" ? "headerReferences" : "footerReferences";
  const mapKey = position === "header" ? "headers" : "footers";

  const newMap = new Map(pkg[mapKey]);
  newMap.delete(activeRId);

  const stripRef = (sp: SectionProperties): SectionProperties => {
    const refs = sp[refKey];
    if (!refs?.some((r) => r.rId === activeRId)) {
      return sp;
    }
    return {
      ...sp,
      [refKey]: refs.filter((r) => r.rId !== activeRId),
    };
  };

  const oldDoc = pkg.document;
  const newSections = oldDoc.sections?.map((s) => ({
    ...s,
    properties: stripRef(s.properties),
  }));
  const newFinalProps = oldDoc.finalSectionProperties
    ? stripRef(oldDoc.finalSectionProperties)
    : oldDoc.finalSectionProperties;

  return {
    ...document,
    package: {
      ...pkg,
      [mapKey]: newMap,
      document: {
        ...oldDoc,
        ...(newSections !== undefined ? { sections: newSections } : {}),
        ...(newFinalProps !== undefined ? { finalSectionProperties: newFinalProps } : {}),
      },
    },
  };
};
