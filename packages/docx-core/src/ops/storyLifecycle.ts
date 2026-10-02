/** Explicit story lifecycle and section property edits, with exact JSON-safe inverses. */
import { panic, Result } from "better-result";
import { DEFAULT_TAB_STOP_TWIPS } from "../model/document";
import type {
  Document,
  DocumentBody,
  HeaderFooter,
  SectionProperties,
  Paragraph,
  BlockContent,
} from "../model/document";
import { withBodyContent, storyParagraphs } from "./blocks";
import { contractViolation } from "./contract";
import type { DocumentEdit } from "./edits";
import { structurallyEqual } from "./equality";
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

/** This total projection forces an ownership decision when the body model gains a field. */
const bodyValues = (body: DocumentBody) =>
  ({
    content: body.content,
    background: body.background ?? null,
    finalSectionProperties: body.finalSectionProperties ?? null,
    comments: body.comments ?? null,
  }) satisfies {
    [Key in keyof Omit<DocumentBody, "sections">]-?: Exclude<DocumentBody[Key], undefined> | null;
  };

/** Capture only fields owned by the lifecycle delta; absent values survive JSON as null. */
export const captureStoryParts = (document: Document, owned: StoryParts): StoryParts => {
  const body = document.package.document;
  const values = bodyValues(body);
  const pkg = document.package;
  return {
    ...(owned.body === undefined
      ? {}
      : {
          body: {
            ...(owned.body.content === undefined ? {} : { content: values.content }),
            ...(owned.body.background === undefined ? {} : { background: values.background }),
            ...(owned.body.finalSectionProperties === undefined
              ? {}
              : {
                  finalSectionProperties: values.finalSectionProperties,
                }),
            ...(owned.body.comments === undefined ? {} : { comments: values.comments }),
          },
        }),
    ...(owned.sections === undefined
      ? {}
      : {
          sections: owned.sections.map(({ index, properties, headers, footers }) => {
            const section = body.sections?.at(index);
            return {
              index,
              ...(properties === undefined ? {} : { properties: section?.properties }),
              ...(headers === undefined
                ? {}
                : { headers: section?.headers ? [...section.headers] : null }),
              ...(footers === undefined
                ? {}
                : { footers: section?.footers ? [...section.footers] : null }),
            };
          }),
        }),
    ...(owned.headers === undefined ? {} : { headers: pkg.headers ? [...pkg.headers] : null }),
    ...(owned.footers === undefined ? {} : { footers: pkg.footers ? [...pkg.footers] : null }),
    ...(owned.footnotes === undefined ? {} : { footnotes: pkg.footnotes ?? null }),
    ...(owned.endnotes === undefined ? {} : { endnotes: pkg.endnotes ?? null }),
    ...(owned.settings === undefined ? {} : { settings: pkg.settings ?? null }),
  };
};

/** Immutable operations share unowned fields; section content is a derived mirror. */
const changedParts = (before: Document, after: Document): StoryParts => {
  const oldBody = before.package.document;
  const newBody = after.package.document;
  const body = {
    ...(oldBody.content === newBody.content ? {} : { content: oldBody.content }),
    ...(oldBody.background === newBody.background
      ? {}
      : { background: oldBody.background ?? null }),
    ...(oldBody.finalSectionProperties === newBody.finalSectionProperties
      ? {}
      : {
          finalSectionProperties: oldBody.finalSectionProperties ?? null,
        }),
    ...(oldBody.comments === newBody.comments ? {} : { comments: oldBody.comments ?? null }),
  };
  const sections =
    oldBody.sections === newBody.sections
      ? []
      : (oldBody.sections ?? []).flatMap((section, index) => {
          const next = newBody.sections?.at(index);
          if (section === next) return [];
          const changes = {
            index,
            ...(section.properties === next?.properties ? {} : { properties: section.properties }),
            ...(section.headers === next?.headers
              ? {}
              : { headers: section.headers ? [...section.headers] : null }),
            ...(section.footers === next?.footers
              ? {}
              : { footers: section.footers ? [...section.footers] : null }),
          };
          return Object.keys(changes).length === 1 ? [] : [changes];
        });
  return {
    ...(Object.keys(body).length === 0 ? {} : { body }),
    ...(sections.length === 0 ? {} : { sections }),
    ...(before.package.headers === after.package.headers
      ? {}
      : {
          headers: before.package.headers ? [...before.package.headers] : null,
        }),
    ...(before.package.footers === after.package.footers
      ? {}
      : {
          footers: before.package.footers ? [...before.package.footers] : null,
        }),
    ...(before.package.footnotes === after.package.footnotes
      ? {}
      : { footnotes: before.package.footnotes ?? null }),
    ...(before.package.endnotes === after.package.endnotes
      ? {}
      : { endnotes: before.package.endnotes ?? null }),
    ...(before.package.settings === after.package.settings
      ? {}
      : { settings: before.package.settings ?? null }),
  };
};

const restoreParts = (document: Document, parts: StoryParts): Document => {
  const pkg = { ...document.package };
  let body = document.package.document;
  if (parts.body !== undefined) {
    body = { ...body };
    const fields = parts.body;
    if (fields.content !== undefined) {
      const existingBlocks = body.content;
      body = withBodyContent(
        body,
        fields.content.map((block, index) => {
          const previous = existingBlocks.at(index);
          return previous && structurallyEqual(block, previous) ? previous : block;
        }),
      );
    }
    if (fields.background !== undefined) {
      if (fields.background === null) delete body.background;
      else body.background = fields.background;
    }
    if (fields.finalSectionProperties !== undefined) {
      if (fields.finalSectionProperties === null) delete body.finalSectionProperties;
      else body.finalSectionProperties = fields.finalSectionProperties;
    }
    if (fields.comments !== undefined) {
      if (fields.comments === null) delete body.comments;
      else body.comments = fields.comments;
    }
  }
  if (parts.headers !== undefined) {
    if (parts.headers === null) delete pkg.headers;
    else
      pkg.headers = new Map(
        parts.headers.map(([rId, part]) => {
          const previous = pkg.headers?.get(rId);
          return [rId, previous && structurallyEqual(previous, part) ? previous : part] as const;
        }),
      );
  }
  if (parts.footers !== undefined) {
    if (parts.footers === null) delete pkg.footers;
    else
      pkg.footers = new Map(
        parts.footers.map(([rId, part]) => {
          const previous = pkg.footers?.get(rId);
          return [rId, previous && structurallyEqual(previous, part) ? previous : part] as const;
        }),
      );
  }
  if (parts.sections !== undefined && body.sections !== undefined) {
    const sections = [...body.sections];
    for (const fields of parts.sections) {
      const previous = sections.at(fields.index);
      if (!previous) return panic("Owned lifecycle section is missing after staleness validation.");
      const section = { ...previous };
      if (fields.properties !== undefined) section.properties = fields.properties;
      if (fields.headers !== undefined) {
        if (fields.headers === null) delete section.headers;
        else
          section.headers = new Map(
            fields.headers.map(([variant, part]) => {
              const rId = section.properties.headerReferences?.find(
                ({ type }) => type === variant,
              )?.rId;
              return [
                variant,
                (rId === undefined ? undefined : pkg.headers?.get(rId)) ?? part,
              ] as const;
            }),
          );
      }
      if (fields.footers !== undefined) {
        if (fields.footers === null) delete section.footers;
        else
          section.footers = new Map(
            fields.footers.map(([variant, part]) => {
              const rId = section.properties.footerReferences?.find(
                ({ type }) => type === variant,
              )?.rId;
              return [
                variant,
                (rId === undefined ? undefined : pkg.footers?.get(rId)) ?? part,
              ] as const;
            }),
          );
      }
      sections[fields.index] = section;
    }
    body = { ...body, sections };
  }
  pkg.document = body;
  if (parts.footnotes !== undefined) {
    if (parts.footnotes === null) delete pkg.footnotes;
    else {
      const current = new Map(pkg.footnotes?.map((note) => [note.id, note]));
      pkg.footnotes = parts.footnotes.map((note) => {
        const previous = current.get(note.id);
        return previous && structurallyEqual(previous, note) ? previous : note;
      });
    }
  }
  if (parts.endnotes !== undefined) {
    if (parts.endnotes === null) delete pkg.endnotes;
    else {
      const current = new Map(pkg.endnotes?.map((note) => [note.id, note]));
      pkg.endnotes = parts.endnotes.map((note) => {
        const previous = current.get(note.id);
        return previous && structurallyEqual(previous, note) ? previous : note;
      });
    }
  }
  if (parts.settings !== undefined) {
    if (parts.settings === null) delete pkg.settings;
    else pkg.settings = parts.settings;
  }
  return { ...document, package: pkg };
};

/** Traverse changed blocks only; shared records cannot add, remove or modify an ID. */
const changedParagraphs = (before: Document, after: Document) => {
  const oldParagraphs = new Map<string, Paragraph>();
  const newParagraphs = new Map<string, Paragraph>();
  const collect = (
    oldContent: readonly BlockContent[] = [],
    newContent: readonly BlockContent[] = [],
  ) => {
    if (oldContent === newContent) return;
    const oldBlocks = new Set(oldContent);
    const newBlocks = new Set(newContent);
    for (const [blocks, shared, output] of [
      [oldContent, newBlocks, oldParagraphs],
      [newContent, oldBlocks, newParagraphs],
    ] as const) {
      for (const block of blocks) {
        if (shared.has(block)) continue;
        for (const { paragraph } of storyParagraphs({ content: [block] }))
          output.set(paragraph.paraId ?? "", paragraph);
      }
    }
  };
  collect(before.package.document.content, after.package.document.content);
  for (const key of ["headers", "footers"] as const) {
    const oldParts = before.package[key];
    const newParts = after.package[key];
    if (oldParts === newParts) continue;
    for (const rId of new Set([...(oldParts?.keys() ?? []), ...(newParts?.keys() ?? [])]))
      collect(oldParts?.get(rId)?.content, newParts?.get(rId)?.content);
  }
  for (const key of ["footnotes", "endnotes"] as const) {
    const oldNotes = before.package[key];
    const newNotes = after.package[key];
    if (oldNotes === newNotes) continue;
    const oldParts = new Map(oldNotes?.map((note) => [note.id, note]));
    const newParts = new Map(newNotes?.map((note) => [note.id, note]));
    for (const id of new Set([...oldParts.keys(), ...newParts.keys()]))
      collect(oldParts.get(id)?.content, newParts.get(id)?.content);
  }
  return { oldParagraphs, newParagraphs };
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
  const prior = changedParts(before, after);
  const next = captureStoryParts(after, prior);
  const { oldParagraphs, newParagraphs } = changedParagraphs(before, after);
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
  let next = { ...body };
  if (end) {
    const content = [...body.content];
    const paragraph = content.at(end.index);
    if (paragraph?.type === "paragraph")
      content[end.index] = { ...paragraph, sectionProperties: properties };
    next = withBodyContent(body, content);
  } else next.finalSectionProperties = properties;
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
    if (!structurallyEqual(captureStoryParts(document, op.expected), op.expected))
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
