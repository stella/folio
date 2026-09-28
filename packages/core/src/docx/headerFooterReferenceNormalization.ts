import { PARSE_WARNING_CODES } from "@stll/docx-core/model";

import type {
  BlockContent,
  DocumentBody,
  HeaderFooter,
  SectionProperties,
  Table,
} from "../types/document";

/** The codes this normalisation is reported under, owned here, not at the caller. */
export const DANGLING_HEADER_REFERENCE_WARNING = PARSE_WARNING_CODES.danglingHeaderReference;
export const DANGLING_FOOTER_REFERENCE_WARNING = PARSE_WARNING_CODES.danglingFooterReference;

type NormalizeHeaderFooterReferencesInput = {
  documentBody: DocumentBody;
  headers?: Map<string, HeaderFooter>;
  footers?: Map<string, HeaderFooter>;
};

type NormalizeHeaderFooterReferencesResult = {
  removedDanglingHeaderReferences: number;
  removedDanglingFooterReferences: number;
};

export const normalizeHeaderFooterReferences = ({
  documentBody,
  headers,
  footers,
}: NormalizeHeaderFooterReferencesInput): NormalizeHeaderFooterReferencesResult => {
  let removedDanglingHeaderReferences = 0;
  let removedDanglingFooterReferences = 0;

  const normalizeSectionProperties = (sectionProperties: SectionProperties): void => {
    const headerResult = removeDanglingReferences(sectionProperties.headerReferences, headers);
    if (headerResult.changed) {
      removedDanglingHeaderReferences += headerResult.removed;
      if (headerResult.references.length > 0) {
        sectionProperties.headerReferences = headerResult.references;
      } else {
        delete sectionProperties.headerReferences;
      }
    }

    const footerResult = removeDanglingReferences(sectionProperties.footerReferences, footers);
    if (footerResult.changed) {
      removedDanglingFooterReferences += footerResult.removed;
      if (footerResult.references.length > 0) {
        sectionProperties.footerReferences = footerResult.references;
      } else {
        delete sectionProperties.footerReferences;
      }
    }
  };

  forEachSectionProperties(documentBody, normalizeSectionProperties);

  return {
    removedDanglingHeaderReferences,
    removedDanglingFooterReferences,
  };
};

/**
 * Visit every distinct section record once, in section order: the paragraph
 * carriers in document order, then the body's final `w:sectPr`.
 */
const forEachSectionProperties = (
  documentBody: DocumentBody,
  visit: (sectionProperties: SectionProperties) => void,
): void => {
  const seen = new Set<SectionProperties>();
  const visitOnce = (sectionProperties: SectionProperties | undefined): void => {
    if (!sectionProperties || seen.has(sectionProperties)) {
      return;
    }
    seen.add(sectionProperties);
    visit(sectionProperties);
  };

  const visitBlocks = (blocks: BlockContent[]): void => {
    for (const block of blocks) {
      if (block.type === "paragraph") {
        visitOnce(block.sectionProperties);
      } else if (block.type === "table") {
        visitTable(block);
      } else if (block.type === "blockSdt" || block.type === "blockCustomXml") {
        visitBlocks(block.content);
      }
    }
  };
  const visitTable = (table: Table): void => {
    for (const row of table.rows) {
      for (const cell of row.cells) {
        visitBlocks(cell.content);
      }
    }
  };

  visitBlocks(documentBody.content);
  visitOnce(documentBody.finalSectionProperties);
  for (const section of documentBody.sections ?? []) {
    visitOnce(section.properties);
  }
};

type AssignHeaderFooterRolesInput = {
  documentBody: DocumentBody;
  headers?: Map<string, HeaderFooter>;
  footers?: Map<string, HeaderFooter>;
};

/**
 * Give each header and footer part the role a section reference states for
 * it.
 *
 * A part says nothing about which pages it serves: the `w:type` of the
 * `w:headerReference` / `w:footerReference` naming it does (ECMA-376 Part 1
 * §17.10.5, `ST_HdrFtr`). The parts are read from the document's
 * relationships, before any section is consulted, so each is read as a
 * default part and takes its role here, from the first section that
 * references it. A part no section references stays default; one several
 * sections reference in different roles keeps the first, and a reader that
 * needs every role reads the section references themselves.
 */
export const assignHeaderFooterRoles = ({
  documentBody,
  headers,
  footers,
}: AssignHeaderFooterRolesInput): void => {
  const assigned = new Set<HeaderFooter>();
  const assign = (
    references: readonly { type: HeaderFooter["hdrFtrType"]; rId: string }[] | undefined,
    parts: Map<string, HeaderFooter> | undefined,
  ): void => {
    for (const { type, rId } of references ?? []) {
      const part = parts?.get(rId);
      if (!part || assigned.has(part)) {
        continue;
      }
      assigned.add(part);
      part.hdrFtrType = type;
    }
  };
  forEachSectionProperties(documentBody, (sectionProperties) => {
    assign(sectionProperties.headerReferences, headers);
    assign(sectionProperties.footerReferences, footers);
  });
};

type HeaderFooterReference = {
  rId: string;
};

type RemoveDanglingReferencesResult<T extends HeaderFooterReference> = {
  changed: boolean;
  references: T[];
  removed: number;
};

const removeDanglingReferences = <T extends HeaderFooterReference>(
  references: T[] | undefined,
  validParts: Map<string, HeaderFooter> | undefined,
): RemoveDanglingReferencesResult<T> => {
  if (!references || !validParts) {
    return { changed: false, references: references ?? [], removed: 0 };
  }

  const kept = references.filter((reference) => validParts.has(reference.rId));
  return {
    changed: kept.length !== references.length,
    references: kept,
    removed: references.length - kept.length,
  };
};
