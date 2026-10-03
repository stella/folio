/** Adapter command boundaries allocate identities; journaled operations remain deterministic. */
import { DOCUMENT_OP_TYPES } from "@stll/docx-core/ops";
import type {
  CreateHeaderFooterOp,
  FormattingPatch,
  RemoveHeaderFooterOp,
  SetSectionPropsOp,
} from "@stll/docx-core/ops";
import { visitDocxParagraphs } from "../docx/paragraphTraversal";
import type {
  BlockContent,
  Document,
  HeaderFooterType,
  SectionProperties,
} from "../types/document";
import { generateHexId, isValidHexId, MAX_HEX_ID_EXCLUSIVE } from "../utils/hexId";

export { DOCUMENT_OP_TYPES } from "@stll/docx-core/ops";
export type { DocumentOp, OpStory, TextPosition, FormattingPatch } from "@stll/docx-core/ops";

export const canonicalFinalSectionIndex = (document: Document): number =>
  document.package.document.content.filter(
    (block) => block.type === "paragraph" && block.sectionProperties !== undefined,
  ).length;

export const withCanonicalParagraphIds = (
  content: BlockContent[],
  document: Document,
): BlockContent[] => {
  const taken = new Set<string>();
  const { document: documentBody, headers, footers, footnotes, endnotes } = document.package;
  visitDocxParagraphs({ documentBody, headers, footers, footnotes, endnotes }, (paragraph) => {
    if (paragraph.paraId) taken.add(paragraph.paraId.toUpperCase());
  });
  const owned = structuredClone(content);
  visitDocxParagraphs({ documentBody: { content: owned } }, (paragraph) => {
    if (
      paragraph.paraId &&
      isValidHexId(paragraph.paraId) &&
      Number.parseInt(paragraph.paraId, 16) > 0 &&
      Number.parseInt(paragraph.paraId, 16) < MAX_HEX_ID_EXCLUSIVE &&
      !taken.has(paragraph.paraId.toUpperCase())
    ) {
      taken.add(paragraph.paraId.toUpperCase());
      return;
    }
    let paraId = generateHexId();
    while (taken.has(paraId)) paraId = generateHexId();
    taken.add(paraId);
    paragraph.paraId = paraId;
  });
  return owned;
};

type CreateCanonicalHeaderFooterOptions = {
  document: Document;
  position: "header" | "footer";
  referenceType: HeaderFooterType;
  sectionIndex?: number;
};

export const createCanonicalHeaderFooterOperation = ({
  document,
  position,
  referenceType,
  sectionIndex = canonicalFinalSectionIndex(document),
}: CreateCanonicalHeaderFooterOptions): CreateHeaderFooterOp => {
  const occupied = new Set([
    ...(document.package.headers?.keys() ?? []),
    ...(document.package.footers?.keys() ?? []),
    ...(document.package.relationships?.keys() ?? []),
  ]);
  const base = `rId_new_${position}_${referenceType}`;
  let rId = base;
  let suffix = 2;
  while (occupied.has(rId)) rId = `${base}_${suffix++}`;
  return {
    type: DOCUMENT_OP_TYPES.CREATE_HEADER_FOOTER,
    sectionIndex,
    story: { kind: position, rId },
    referenceType,
    content: withCanonicalParagraphIds([{ type: "paragraph", content: [] }], document),
  };
};

type RemoveCanonicalHeaderFooterOptions = {
  document: Document;
  position: "header" | "footer";
  rId: string;
};

/** Remove each explicit binding, leaving inherited sections to inherit the remaining parts. */
export const removeCanonicalHeaderFooterOperations = ({
  document,
  position,
  rId,
}: RemoveCanonicalHeaderFooterOptions): RemoveHeaderFooterOp[] => {
  const properties: SectionProperties[] = [];
  for (const block of document.package.document.content) {
    if (block.type === "paragraph" && block.sectionProperties)
      properties.push(block.sectionProperties);
  }
  properties.push(document.package.document.finalSectionProperties ?? {});
  const operations: RemoveHeaderFooterOp[] = [];
  for (const [sectionIndex, section] of properties.entries()) {
    const references = position === "header" ? section.headerReferences : section.footerReferences;
    for (const reference of references ?? []) {
      if (reference.rId !== rId) continue;
      operations.push({
        type: DOCUMENT_OP_TYPES.REMOVE_HEADER_FOOTER,
        sectionIndex,
        story: { kind: position, rId },
        referenceType: reference.type,
      });
    }
  }
  return operations;
};

export const createCanonicalSectionPropertiesOperation = (
  document: Document,
  patch: FormattingPatch<SectionProperties>,
): SetSectionPropsOp => ({
  type: DOCUMENT_OP_TYPES.SET_SECTION_PROPS,
  sectionIndex: canonicalFinalSectionIndex(document),
  patch,
});
