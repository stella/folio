/** Explicit story lifecycle and section property edits, with exact JSON-safe inverses. */
import { Result } from "better-result";
import { DEFAULT_TAB_STOP_TWIPS } from "../model/document";
import type {
  Document,
  DocumentBody,
  HeaderFooter,
  SectionProperties,
  Section,
} from "../model/document";
import { withBodyContent, storyParagraphs } from "./blocks";
import { contractViolation } from "./contract";
import type { DocumentEdit } from "./edits";
import { structurallyEqual } from "./equality";
import { documentStories, storyBody } from "./stories";
import { applyFormattingPatch } from "./patch";
import { DOCUMENT_OP_REFUSAL_REASONS, DocumentOpRefusal } from "./refusal";
import {
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  type StoryParts,
  type CreateHeaderFooterOp,
  type RemoveHeaderFooterOp,
  type SetSectionPropsOp,
  type RestoreStoryPartsOp,
} from "./types";

type LifecycleOp =
  | CreateHeaderFooterOp
  | RemoveHeaderFooterOp
  | SetSectionPropsOp
  | RestoreStoryPartsOp;
const refuse = (op: DocumentOp, message: string) =>
  Result.err(
    new DocumentOpRefusal({
      opType: op.type,
      reason: DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
      message,
    }),
  );

export const captureStoryParts = (document: Document): StoryParts => {
  const { sections, ...body } = document.package.document;
  const pkg = document.package;
  return {
    body,
    ...(sections === undefined
      ? {}
      : {
          sections: sections.map(({ headers, footers, ...section }) =>
            Object.assign(
              {},
              section,
              headers === undefined ? {} : { headers: [...headers] },
              footers === undefined ? {} : { footers: [...footers] },
            ),
          ),
        }),
    ...(pkg.headers === undefined ? {} : { headers: [...pkg.headers] }),
    ...(pkg.footers === undefined ? {} : { footers: [...pkg.footers] }),
    ...(pkg.footnotes === undefined ? {} : { footnotes: pkg.footnotes }),
    ...(pkg.endnotes === undefined ? {} : { endnotes: pkg.endnotes }),
    ...(pkg.settings === undefined ? {} : { settings: pkg.settings }),
  };
};

const restoreParts = (document: Document, parts: StoryParts): Document => {
  const current = captureStoryParts(document);
  const pkg = { ...document.package };
  const body: DocumentBody = { ...parts.body };
  const existingBlocks = document.package.document.content;
  body.content = body.content.map((block, index) => {
    const previous = existingBlocks.at(index);
    return previous && structurallyEqual(block, previous) ? previous : block;
  });
  if (parts.sections !== undefined)
    body.sections = parts.sections.map(({ headers, footers, ...section }) => ({
      ...section,
      ...(headers === undefined ? {} : { headers: new Map(headers) }),
      ...(footers === undefined ? {} : { footers: new Map(footers) }),
    }));
  if (
    !structurallyEqual(current.body, parts.body) ||
    !structurallyEqual(current.sections, parts.sections)
  )
    pkg.document = body;
  if (!structurallyEqual(current.headers, parts.headers)) {
    if (parts.headers === undefined) delete pkg.headers;
    else
      pkg.headers = new Map(
        parts.headers.map(([rId, part]) => {
          const previous = document.package.headers?.get(rId);
          return [rId, previous && structurallyEqual(previous, part) ? previous : part] as const;
        }),
      );
  }
  if (!structurallyEqual(current.footers, parts.footers)) {
    if (parts.footers === undefined) delete pkg.footers;
    else
      pkg.footers = new Map(
        parts.footers.map(([rId, part]) => {
          const previous = document.package.footers?.get(rId);
          return [rId, previous && structurallyEqual(previous, part) ? previous : part] as const;
        }),
      );
  }
  if (!structurallyEqual(current.footnotes, parts.footnotes)) {
    if (parts.footnotes === undefined) delete pkg.footnotes;
    else
      pkg.footnotes = parts.footnotes.map((note) => {
        const previous = document.package.footnotes?.find(({ id }) => id === note.id);
        return previous && structurallyEqual(previous, note) ? previous : note;
      });
  }
  if (!structurallyEqual(current.endnotes, parts.endnotes)) {
    if (parts.endnotes === undefined) delete pkg.endnotes;
    else
      pkg.endnotes = parts.endnotes.map((note) => {
        const previous = document.package.endnotes?.find(({ id }) => id === note.id);
        return previous && structurallyEqual(previous, note) ? previous : note;
      });
  }
  if (!structurallyEqual(current.settings, parts.settings)) {
    if (parts.settings === undefined) delete pkg.settings;
    else pkg.settings = parts.settings;
  }
  return { ...document, package: pkg };
};

export const storyLifecycleEdit = (
  before: Document,
  after: Document,
  op: DocumentOp,
): Result<DocumentEdit, DocumentOpRefusal> => {
  const invalid = contractViolation(after);
  if (invalid)
    return Result.err(
      new DocumentOpRefusal({ opType: op.type, reason: invalid.reason, message: invalid.message }),
    );
  const prior = captureStoryParts(before);
  const next = captureStoryParts(after);
  const paragraphs = (document: Document) =>
    new Map(
      documentStories(document)
        .flatMap((story) => storyParagraphs(storyBody(document, story)))
        .map(({ paragraph }) => [paragraph.paraId ?? "", paragraph]),
    );
  const oldParagraphs = paragraphs(before);
  const newParagraphs = paragraphs(after);
  const oldIds = new Set(oldParagraphs.keys());
  const newIds = new Set(newParagraphs.keys());
  return Result.ok({
    document: after,
    inverse: [
      {
        type: DOCUMENT_OP_TYPES.RESTORE_STORY_PARTS,
        expected: structuredClone(next),
        parts: structuredClone(prior),
      },
    ],
    touched: {
      modified: [...newIds].filter(
        (id) => oldIds.has(id) && !structurallyEqual(oldParagraphs.get(id), newParagraphs.get(id)),
      ),
      inserted: [...newIds].filter((id) => !oldIds.has(id)),
      removed: [...oldIds].filter((id) => !newIds.has(id)),
    },
  });
};

/** Ordered section breaks; the final section follows every paragraph sectPr. */
const sectionEnds = (body: DocumentBody) =>
  body.content.flatMap((block, index) =>
    block.type === "paragraph" && block.sectionProperties !== undefined
      ? [{ index, properties: block.sectionProperties }]
      : [],
  );
export const sectionPropertiesAt = (
  document: Document,
  sectionIndex: number,
): SectionProperties | undefined => {
  if (!Number.isInteger(sectionIndex) || sectionIndex < 0) return undefined;
  const ends = sectionEnds(document.package.document);
  if (sectionIndex > ends.length) return undefined;
  return (
    ends.at(sectionIndex)?.properties ??
    document.package.document.finalSectionProperties ??
    document.package.document.sections?.at(sectionIndex)?.properties ??
    {}
  );
};

const withSectionProperties = (
  document: Document,
  sectionIndex: number,
  properties: SectionProperties,
): Document => {
  const body = document.package.document;
  const end = sectionEnds(body).at(sectionIndex);
  const content = [...body.content];
  if (end) {
    const paragraph = content.at(end.index);
    if (paragraph?.type === "paragraph")
      content[end.index] = { ...paragraph, sectionProperties: properties };
  }
  const next = withBodyContent(body, content);
  if (!end) next.finalSectionProperties = properties;
  if (next.sections)
    next.sections = next.sections.map((section, index) =>
      index === sectionIndex ? { ...section, properties } : section,
    );
  return { ...document, package: { ...document.package, document: next } };
};

export const applyStoryLifecycle = (
  document: Document,
  op: LifecycleOp,
): Result<DocumentEdit, DocumentOpRefusal> => {
  if (op.type === DOCUMENT_OP_TYPES.RESTORE_STORY_PARTS) {
    if (!structurallyEqual(captureStoryParts(document), op.expected))
      return Result.err(
        new DocumentOpRefusal({
          opType: op.type,
          reason: DOCUMENT_OP_REFUSAL_REASONS.STALE,
          message: "Story lifecycle inverse is stale.",
        }),
      );
    return storyLifecycleEdit(document, restoreParts(document, structuredClone(op.parts)), op);
  }
  const properties = sectionPropertiesAt(document, op.sectionIndex);
  if (!properties) return refuse(op, "The section does not exist.");
  if (op.type === DOCUMENT_OP_TYPES.SET_SECTION_PROPS) {
    const nextProps = applyFormattingPatch(properties, op.patch) ?? {};
    let next = withSectionProperties(document, op.sectionIndex, nextProps);
    if (
      op.patch.evenAndOddHeaders !== undefined &&
      (op.patch.evenAndOddHeaders !== null || next.package.settings !== undefined)
    ) {
      const settings = {
        ...(next.package.settings ?? { defaultTabStop: DEFAULT_TAB_STOP_TWIPS }),
      };
      if (op.patch.evenAndOddHeaders === null) delete settings.evenAndOddHeaders;
      else settings.evenAndOddHeaders = op.patch.evenAndOddHeaders;
      next = { ...next, package: { ...next.package, settings } };
    }
    return storyLifecycleEdit(document, next, op);
  }
  const kind = op.story.kind;
  const refs = kind === "header" ? properties.headerReferences : properties.footerReferences;
  const parts = kind === "header" ? document.package.headers : document.package.footers;
  const nextParts = new Map(parts);
  let nextRefs = [...(refs ?? [])];
  const nextProps = { ...properties };
  if (op.type === DOCUMENT_OP_TYPES.CREATE_HEADER_FOOTER) {
    if (
      !op.story.rId ||
      document.package.relationships?.has(op.story.rId) ||
      document.package.headers?.has(op.story.rId) ||
      document.package.footers?.has(op.story.rId) ||
      refs?.some(({ type }) => type === op.referenceType)
    )
      return refuse(op, "The header/footer identity or section variant is already used.");
    if (op.content.length === 0 || storyParagraphs({ content: op.content }).length === 0)
      return refuse(op, "A header/footer needs an addressable paragraph.");
    const part: HeaderFooter = {
      type: kind,
      hdrFtrType: op.referenceType,
      content: structuredClone(op.content),
    };
    nextParts.set(op.story.rId, part);
    nextRefs.push({ type: op.referenceType, rId: op.story.rId });
    if (op.referenceType === "first") nextProps.titlePg = true;
    if (op.referenceType === "even") nextProps.evenAndOddHeaders = true;
  } else {
    if (
      !parts?.has(op.story.rId) ||
      !refs?.some(({ type, rId }) => type === op.referenceType && rId === op.story.rId)
    )
      return refuse(op, "The section does not own this header/footer variant.");
    nextRefs = nextRefs.filter(({ type }) => type !== op.referenceType);
    // Shared parts survive while another section still references them.
    const usedElsewhere = [
      ...sectionEnds(document.package.document).map(({ properties: sectionProps }) => sectionProps),
      document.package.document.finalSectionProperties ?? {},
    ].some(
      (sectionProps, index) =>
        index !== op.sectionIndex &&
        (kind === "header" ? sectionProps.headerReferences : sectionProps.footerReferences)?.some(
          ({ rId }) => rId === op.story.rId,
        ),
    );
    if (!usedElsewhere && !nextRefs.some(({ rId }) => rId === op.story.rId))
      nextParts.delete(op.story.rId);
  }
  if (kind === "header") nextProps.headerReferences = nextRefs;
  else nextProps.footerReferences = nextRefs;
  let next = withSectionProperties(document, op.sectionIndex, nextProps);
  next =
    kind === "header"
      ? { ...next, package: { ...next.package, headers: nextParts } }
      : { ...next, package: { ...next.package, footers: nextParts } };
  if (op.type === DOCUMENT_OP_TYPES.CREATE_HEADER_FOOTER && op.referenceType === "even")
    next = {
      ...next,
      package: {
        ...next.package,
        settings: {
          ...(next.package.settings ?? { defaultTabStop: DEFAULT_TAB_STOP_TWIPS }),
          evenAndOddHeaders: true,
        },
      },
    };
  if (next.package.document.sections) {
    const sections = next.package.document.sections.map((section, index): Section => {
      if (index !== op.sectionIndex) return section;
      const bound = new Map(kind === "header" ? section.headers : section.footers);
      const part = nextParts.get(op.story.rId);
      if (op.type === DOCUMENT_OP_TYPES.CREATE_HEADER_FOOTER && part)
        bound.set(op.referenceType, part);
      else bound.delete(op.referenceType);
      return kind === "header" ? { ...section, headers: bound } : { ...section, footers: bound };
    });
    next = {
      ...next,
      package: { ...next.package, document: { ...next.package.document, sections } },
    };
  }
  return storyLifecycleEdit(document, next, op);
};
