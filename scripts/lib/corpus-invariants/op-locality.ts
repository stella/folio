import path from "node:path";
import { parseRelationships, RELATIONSHIP_TYPES } from "@stll/folio-core/docx/relsParser";
import {
  parseXmlDocument,
  getChildElements,
  getLocalName,
  getAttributes,
  getAttribute,
  getNamespaceUri,
} from "@stll/folio-core/docx/xmlParser";
import type {
  BlockContent,
  Document,
  Relationship,
  SectionProperties,
  DocumentBody,
  HeaderFooter,
} from "../../../packages/docx-core/src/model/document";
import {
  COMMENT_PART_RELATIONSHIPS,
  DEFAULT_TAB_STOP_TWIPS,
} from "../../../packages/docx-core/src/model/document";
import { Result } from "better-result";
import {
  documentStories,
  storyBody,
  type DocumentOp,
  DOCUMENT_OP_TYPES,
  OP_STORIES,
  sameStory,
  isEmptyWatermarkHostParagraph,
  type OpStory,
} from "../../../packages/docx-core/src/ops/documentOps";
import type { PackageResources, StoryParts } from "../../../packages/docx-core/src/ops/types";
import { storyParagraphs } from "../../../packages/docx-core/src/ops/blocks";
import { leafSpans, isCommentAnchor } from "../../../packages/docx-core/src/ops/leaves";
import type { CommentOp } from "../../../packages/docx-core/src/ops/types";
import { idKey } from "../../../packages/docx-core/src/ops/ids";
import { failureFromAssertion } from "../corpus-signature";
import {
  type CorpusInvariantInput,
  type CorpusInvariantOutcome,
  EXTENDED_CORPUS_INVARIANTS,
  timeStage,
} from "./contract";
import {
  generateOpSequence,
  OP_SEQUENCE_SEEDS,
  prepareOpDocument,
  sameOpModel,
  seedFromBytes,
  serializeOpDocument,
  serializedOpParts,
  type OpSequenceStep,
  type OpSequence,
} from "./op-sequences";
import { firstDifferingOpPart } from "./op-part-difference";

import { opErrorOutcome } from "./op-outcome";

const INVARIANT = EXTENDED_CORPUS_INVARIANTS.opLocality;

/** Remove declared paragraphs and their empty ancestors, preserving every other field. */
const untouchedBlocks = (
  blocks: readonly BlockContent[],
  touched: ReadonlySet<string>,
): BlockContent[] => {
  const out: BlockContent[] = [];
  for (const block of blocks) {
    switch (block.type) {
      case "paragraph":
        if (block.paraId === undefined || !touched.has(idKey(block.paraId))) out.push(block);
        break;
      case "table": {
        const rows = block.rows.flatMap((row) => {
          const cells = row.cells.flatMap((cell) => {
            const content = untouchedBlocks(cell.content, touched);
            return content.length === 0 ? [] : [{ ...cell, content }];
          });
          return cells.length === 0 ? [] : [{ ...row, cells }];
        });
        if (rows.length > 0) out.push({ ...block, rows });
        break;
      }
      case "blockSdt":
      case "blockCustomXml": {
        const content = untouchedBlocks(block.content, touched);
        if (content.length > 0) out.push({ ...block, content });
        break;
      }
      case "preservedBlock":
      case "bookmarkStart":
      case "bookmarkEnd":
        out.push(block);
        break;
      default: {
        const unreachable: never = block;
        return unreachable;
      }
    }
  }
  return out;
};

const lifecycleOwnership = (op: DocumentOp) => {
  switch (op.type) {
    case DOCUMENT_OP_TYPES.CREATE_HEADER_FOOTER:
      return {
        sectionIndex: op.sectionIndex,
        sectionKeys: [
          op.story.kind === "header" ? "headerReferences" : "footerReferences",
          ...(op.referenceType === "first" ? ["titlePg"] : []),
          ...(op.referenceType === "even" ? ["evenAndOddHeaders"] : []),
        ],
        story: op.story,
        settingsEven: op.referenceType === "even",
      };
    case DOCUMENT_OP_TYPES.REMOVE_HEADER_FOOTER:
      return {
        sectionIndex: op.sectionIndex,
        sectionKeys: [op.story.kind === "header" ? "headerReferences" : "footerReferences"],
        story: op.story,
        settingsEven: false,
      };
    case DOCUMENT_OP_TYPES.SET_SECTION_PROPS:
      return {
        sectionIndex: op.sectionIndex,
        sectionKeys: Object.entries(op.patch)
          .filter(([, value]) => value !== undefined)
          .map(([key]) => key),
        story: undefined,
        settingsEven: op.patch.evenAndOddHeaders !== undefined,
      };
    case DOCUMENT_OP_TYPES.ADD_NOTE:
      return {
        sectionIndex: undefined,
        sectionKeys: [],
        story: { kind: op.note.type, id: op.note.id },
        settingsEven: false,
      };
    case DOCUMENT_OP_TYPES.REMOVE_NOTE:
      return { sectionIndex: undefined, sectionKeys: [], story: op.story, settingsEven: false };
    default:
      return undefined;
  }
};

/** Explicit lifecycle payloads may change a header relationship; watermark decorations may not. */
const explicitHeaderRelationshipIds = (op: DocumentOp): string[] => {
  switch (op.type) {
    case DOCUMENT_OP_TYPES.CREATE_HEADER_FOOTER:
    case DOCUMENT_OP_TYPES.REMOVE_HEADER_FOOTER:
      return op.story.kind === "header" ? [op.story.rId] : [];
    case DOCUMENT_OP_TYPES.RESTORE_STORY_PARTS: {
      if (op.parts.headers === undefined) return [];
      const prior = new Map(op.expected.headers ?? []);
      const next = new Map(op.parts.headers ?? []);
      return [...new Set([...prior.keys(), ...next.keys()])].filter(
        (rId) => !sameOpModel(prior.get(rId), next.get(rId)),
      );
    }
    default:
      return [];
  }
};

const watermarkHostId = (header: HeaderFooter) => {
  if (header.watermarkBlockIndex === undefined) return undefined;
  const host = header.content.at(header.watermarkBlockIndex);
  return host?.type === "paragraph" && isEmptyWatermarkHostParagraph(host)
    ? host.paraId
    : undefined;
};

type WithoutOwnedRecordsOptions = {
  document: Document;
  original: Document;
  counterpart: Document;
  op: DocumentOp;
};

/** Restoration owns precisely the present payload fields, including fields restoring absence. */
const withoutRestoredRecords = (document: Document, parts: StoryParts): Document => {
  const out = structuredClone(document);
  const body = out.package.document;
  const stripBody = {
    content: () => {
      if (parts.body?.content !== undefined) body.content = [];
    },
    background: () => {
      if (parts.body?.background !== undefined) delete body.background;
    },
    finalSectionProperties: () => {
      if (parts.body?.finalSectionProperties !== undefined) delete body.finalSectionProperties;
    },
    comments: () => {
      if (parts.body?.comments !== undefined) delete body.comments;
    },
  } satisfies Record<keyof NonNullable<StoryParts["body"]>, () => void>;
  for (const strip of Object.values(stripBody)) strip();
  const stripPackage = {
    headers: () => {
      if (parts.headers !== undefined) delete out.package.headers;
    },
    footers: () => {
      if (parts.footers !== undefined) delete out.package.footers;
    },
    footnotes: () => {
      if (parts.footnotes !== undefined) delete out.package.footnotes;
    },
    endnotes: () => {
      if (parts.endnotes !== undefined) delete out.package.endnotes;
    },
    settings: () => {
      if (parts.settings !== undefined) delete out.package.settings;
    },
  } satisfies Record<keyof Omit<StoryParts, "body" | "sections" | "undefinedFields">, () => void>;
  for (const strip of Object.values(stripPackage)) strip();
  const restoredSections = new Set(
    (parts.sections ?? []).flatMap(({ index, properties }) =>
      properties === undefined ? [] : [index],
    ),
  );
  let sectionIndex = 0;
  for (const block of body.content) {
    if (block.type !== "paragraph" || block.sectionProperties === undefined) continue;
    if (restoredSections.has(sectionIndex)) delete block.sectionProperties;
    sectionIndex += 1;
  }
  if (restoredSections.has(sectionIndex)) delete body.finalSectionProperties;
  // Section content and mounted story maps are derived mirrors, not independent ownership.
  delete body.sections;
  return out;
};

const isCommentOp = (op: DocumentOp): op is CommentOp => {
  switch (op.type) {
    case DOCUMENT_OP_TYPES.CREATE_COMMENT:
    case DOCUMENT_OP_TYPES.UPDATE_COMMENT_CONTENT:
    case DOCUMENT_OP_TYPES.SET_COMMENT_RESOLUTION:
    case DOCUMENT_OP_TYPES.DELETE_COMMENT:
    case DOCUMENT_OP_TYPES.RESTORE_COMMENT_STATE:
      return true;
    default:
      return false;
  }
};
const ownedCommentIds = (op: CommentOp, original: Document): Set<number> => {
  if (op.type === DOCUMENT_OP_TYPES.CREATE_COMMENT) return new Set([op.comment.id]);
  if (op.type === DOCUMENT_OP_TYPES.RESTORE_COMMENT_STATE) return new Set(op.ids);
  const ids = new Set([op.id]);
  if (op.type !== DOCUMENT_OP_TYPES.DELETE_COMMENT) return ids;
  let previous = 0;
  while (previous !== ids.size) {
    previous = ids.size;
    for (const comment of original.package.document.comments ?? [])
      if (comment.parentId !== undefined && ids.has(comment.parentId)) ids.add(comment.id);
  }
  return ids;
};
const commentStories = (op: CommentOp, original: Document): OpStory[] => {
  if (op.type === DOCUMENT_OP_TYPES.RESTORE_COMMENT_STATE)
    return [...op.expected.anchors, ...op.state.anchors].map(({ story }) => story);
  if (
    op.type === DOCUMENT_OP_TYPES.UPDATE_COMMENT_CONTENT ||
    op.type === DOCUMENT_OP_TYPES.SET_COMMENT_RESOLUTION
  )
    return [];
  if (op.type === DOCUMENT_OP_TYPES.CREATE_COMMENT) {
    switch (op.anchor.kind) {
      case "point":
        return [op.anchor.at.story];
      case "range":
        return [op.anchor.from.story];
      case "revision":
        return [op.anchor.story];
      case "reply":
        break;
    }
  }
  const ids =
    op.type === DOCUMENT_OP_TYPES.CREATE_COMMENT && op.anchor.kind === "reply"
      ? new Set([op.anchor.parentId])
      : ownedCommentIds(op, original);
  return documentStories(original).filter((story) =>
    storyParagraphs(storyBody(original, story)).some(({ paragraph }) =>
      leafSpans(paragraph.content).some(
        ({ node }) => isCommentAnchor(node) && "id" in node && ids.has(node.id),
      ),
    ),
  );
};

const withoutOwnedRecords = ({
  document,
  original,
  counterpart,
  op,
}: WithoutOwnedRecordsOptions): Document => {
  if (isCommentOp(op)) {
    const out = structuredClone(document);
    const ids = ownedCommentIds(op, original);
    const comments = out.package.document.comments ?? [];
    if (
      op.type === DOCUMENT_OP_TYPES.UPDATE_COMMENT_CONTENT ||
      op.type === DOCUMENT_OP_TYPES.SET_COMMENT_RESOLUTION
    ) {
      for (const comment of comments) {
        if (!ids.has(comment.id)) continue;
        if (op.type === DOCUMENT_OP_TYPES.UPDATE_COMMENT_CONTENT) {
          comment.content = [];
          for (const key of Object.keys(op.patch ?? {})) Reflect.deleteProperty(comment, key);
        } else delete comment.done;
      }
    } else {
      out.package.document.comments = comments.filter(({ id }) => !ids.has(id));
      if (out.package.document.comments.length === 0) delete out.package.document.comments;
    }
    if (out.package.relationships !== undefined) {
      out.package.relationships = new Map(
        [...out.package.relationships].filter(
          ([, relation]) =>
            counterpart.package.relationships?.has(relation.id) ||
            !Object.values(COMMENT_PART_RELATIONSHIPS).some(({ type }) => type === relation.type),
        ),
      );
      if (out.package.relationships.size === 0) delete out.package.relationships;
    }
    return out;
  }
  if (op.type === DOCUMENT_OP_TYPES.SET_PACKAGE_RESOURCES) {
    const out = structuredClone(document);
    for (const [key, resource] of Object.entries(op.resources))
      if (!sameOpModel(resource, Reflect.get(op.expected, key)))
        Reflect.deleteProperty(out.package, key);
    // Keyed package maps own exactly the entries the operation carries, or the whole map
    // when its presence changes.
    const maps = { relationships: op.relationships, media: op.media } as const;
    for (const [field, change] of Object.entries(maps)) {
      if (change.expected !== change.next) {
        Reflect.deleteProperty(out.package, field);
        continue;
      }
      const entries: unknown = Reflect.get(out.package, field);
      if (!(entries instanceof Map)) continue;
      for (const { key } of change.entries) entries.delete(key);
    }
    return out;
  }
  if (
    op.type === DOCUMENT_OP_TYPES.CREATE_NUMBERING_INSTANCE ||
    op.type === DOCUMENT_OP_TYPES.DELETE_NUMBERING_INSTANCE
  ) {
    const out = structuredClone(document);
    if (out.package.numbering !== undefined) {
      out.package.numbering.nums = out.package.numbering.nums.filter(
        ({ numId }) => numId !== op.num.numId,
      );
      if (op.abstractNum !== undefined)
        out.package.numbering.abstractNums = out.package.numbering.abstractNums.filter(
          ({ abstractNumId }) => abstractNumId !== op.abstractNum?.abstractNumId,
        );
      if (
        (original.package.numbering === undefined ||
          (op.type === DOCUMENT_OP_TYPES.DELETE_NUMBERING_INSTANCE &&
            op.restore?.type !== "definitions")) &&
        out.package.numbering.nums.length === 0 &&
        out.package.numbering.abstractNums.length === 0
      ) {
        if (Object.hasOwn(original.package, "numbering")) out.package.numbering = undefined;
        else delete out.package.numbering;
      }
    }
    return out;
  }
  if (op.type === DOCUMENT_OP_TYPES.SET_SECTION_ENDPOINT) {
    const out = structuredClone(document);
    if (op.endpoint.type === "final") delete out.package.document.finalSectionProperties;
    else {
      const endpoint = op.endpoint;
      for (const block of out.package.document.content)
        if (
          block.type === "paragraph" &&
          block.paraId !== undefined &&
          idKey(block.paraId) === idKey(endpoint.blockId)
        )
          delete block.sectionProperties;
    }
    delete out.package.document.sections;
    return out;
  }
  if (op.type === DOCUMENT_OP_TYPES.SET_DOCUMENT_WATERMARK) {
    const out = structuredClone(document);
    const created = new Set(op.coverage.map(({ rId }) => rId));
    if (out.package.headers !== undefined) {
      for (const [rId, header] of out.package.headers) {
        if (created.has(rId)) {
          out.package.headers.delete(rId);
          continue;
        }
        const source = original.package.headers?.get(rId);
        const hostId = op.change.kind === "remove" && source ? watermarkHostId(source) : undefined;
        const insertedHost = op.hosts.find((host) => host.rId === rId)?.paraId;
        const ownedHosts = new Set([hostId, insertedHost].filter((id) => id !== undefined));
        header.content = header.content.filter(
          (block) =>
            block.type !== "paragraph" ||
            block.paraId === undefined ||
            !ownedHosts.has(block.paraId),
        );
        delete header.watermark;
        delete header.rawWatermarkXml;
        delete header.watermarkBlockIndex;
      }
      if (out.package.headers.size === 0 && original.package.headers === undefined)
        delete out.package.headers;
    }
    const stripReferences = (
      properties: SectionProperties | undefined,
      source: SectionProperties | undefined,
    ) => {
      if (!properties?.headerReferences) return;
      properties.headerReferences = properties.headerReferences.filter(
        ({ rId }) => !created.has(rId),
      );
      if (properties.headerReferences.length === 0 && source?.headerReferences === undefined)
        delete properties.headerReferences;
    };
    for (const block of out.package.document.content) {
      if (block.type !== "paragraph") continue;
      const source = original.package.document.content.find(
        (candidate) => candidate.type === "paragraph" && candidate.paraId === block.paraId,
      );
      stripReferences(
        block.sectionProperties,
        source?.type === "paragraph" ? source.sectionProperties : undefined,
      );
    }
    stripReferences(
      out.package.document.finalSectionProperties,
      original.package.document.finalSectionProperties,
    );
    if (
      out.package.document.finalSectionProperties !== undefined &&
      Object.keys(out.package.document.finalSectionProperties).length === 0 &&
      original.package.document.finalSectionProperties === undefined
    )
      delete out.package.document.finalSectionProperties;
    if (out.package.relationships !== undefined) {
      for (const rId of created) out.package.relationships.delete(rId);
      if (out.package.relationships.size === 0 && original.package.relationships === undefined)
        delete out.package.relationships;
    }
    delete out.package.document.sections;
    return out;
  }
  if (op.type === DOCUMENT_OP_TYPES.RESTORE_STORY_PARTS)
    return withoutRestoredRecords(document, op.parts);
  const ownership = lifecycleOwnership(op);
  if (!ownership) return document;
  const out = structuredClone(document);
  const stripProperties = (
    properties: SectionProperties | undefined,
  ): SectionProperties | undefined => {
    if (!properties) return undefined;
    const remaining = Object.fromEntries(
      Object.entries(properties).filter(
        ([key]) => !ownership.sectionKeys.some((ownedKey) => ownedKey === key),
      ),
    );
    if (
      op.type === DOCUMENT_OP_TYPES.CREATE_HEADER_FOOTER ||
      op.type === DOCUMENT_OP_TYPES.REMOVE_HEADER_FOOTER
    ) {
      const referenceKey = op.story.kind === "header" ? "headerReferences" : "footerReferences";
      const references = properties[referenceKey]?.filter(
        ({ type, rId }) => type !== op.referenceType || rId !== op.story.rId,
      );
      if (references && references.length > 0) remaining[referenceKey] = references;
    }

    return Object.keys(remaining).length === 0 ? undefined : remaining;
  };
  let sectionIndex = 0;
  for (const block of out.package.document.content) {
    if (block.type !== "paragraph" || block.sectionProperties === undefined) continue;
    if (sectionIndex === ownership.sectionIndex) {
      const properties = stripProperties(block.sectionProperties);
      if (properties === undefined) delete block.sectionProperties;
      else block.sectionProperties = properties;
    }
    sectionIndex += 1;
  }
  if (sectionIndex === ownership.sectionIndex) {
    const properties = stripProperties(out.package.document.finalSectionProperties);
    if (properties === undefined) delete out.package.document.finalSectionProperties;
    else out.package.document.finalSectionProperties = properties;
  }
  const story = ownership.story;
  if (story?.kind === "header") {
    out.package.headers = new Map(out.package.headers);
    out.package.headers.delete(story.rId);
  }
  if (story?.kind === "footer") {
    out.package.footers = new Map(out.package.footers);
    out.package.footers.delete(story.rId);
  }
  if (story?.kind === "footnote")
    out.package.footnotes = (out.package.footnotes ?? []).filter(({ id }) => id !== story.id);
  if (story?.kind === "endnote")
    out.package.endnotes = (out.package.endnotes ?? []).filter(({ id }) => id !== story.id);
  if (ownership.settingsEven && out.package.settings !== undefined) {
    const settings = { ...out.package.settings };
    delete settings.evenAndOddHeaders;
    // Creating settings owns only the model-required default; authored values remain compared.
    if (
      original.package.settings === undefined &&
      settings.defaultTabStop === DEFAULT_TAB_STOP_TWIPS &&
      Object.keys(settings).length === 1
    )
      delete out.package.settings;
    else out.package.settings = settings;
  }
  // The model's section content/maps mirror owned package stories; body data remains the oracle.
  delete out.package.document.sections;
  return out;
};

const projected = (document: Document, touched: ReadonlySet<string>): Document => {
  const projectContent = <Part extends Pick<DocumentBody, "content">>(part: Part) => ({
    ...part,
    content: untouchedBlocks(part.content, touched),
  });
  const body = {
    ...projectContent(document.package.document),
  };
  delete body.sections;
  const pkg = { ...document.package, document: body };
  if (pkg.headers !== undefined)
    pkg.headers = new Map([...pkg.headers].map(([id, part]) => [id, projectContent(part)]));
  if (pkg.footers !== undefined)
    pkg.footers = new Map([...pkg.footers].map(([id, part]) => [id, projectContent(part)]));
  if (pkg.footnotes !== undefined) pkg.footnotes = pkg.footnotes.map(projectContent);
  if (pkg.endnotes !== undefined) pkg.endnotes = pkg.endnotes.map(projectContent);
  return { ...document, package: pkg };
};

const addressedStory = (op: DocumentOp) => {
  if ("at" in op && typeof op.at === "object" && "story" in op.at) return op.at.story;
  if ("from" in op) return op.from.story;
  if ("story" in op) return op.story;
  return undefined;
};

type StoryOwnershipOptions = { op: DocumentOp; story: OpStory; original: Document };
const storyOwnership = ({
  op,
  story,
  original,
}: StoryOwnershipOptions): "content" | "fields" | "none" => {
  if (isCommentOp(op))
    return commentStories(op, original).some((owned) => sameStory(owned, story))
      ? "content"
      : "none";
  if (op.type === DOCUMENT_OP_TYPES.SET_DOCUMENT_WATERMARK)
    return story !== OP_STORIES.MAIN &&
      story.kind === "header" &&
      op.coverage.some(({ rId }) => rId === story.rId)
      ? "content"
      : "none";

  const addressed = addressedStory(op);
  if (addressed !== undefined && sameStory(addressed, story)) return "content";
  const lifecycle = lifecycleOwnership(op);
  if (lifecycle !== undefined) {
    if (lifecycle.story !== undefined && sameStory(lifecycle.story, story)) return "content";
    if (lifecycle.sectionIndex !== undefined && story === OP_STORIES.MAIN) return "fields";
  }
  if (op.type !== DOCUMENT_OP_TYPES.RESTORE_STORY_PARTS) return "none";
  if (story === OP_STORIES.MAIN) {
    if (op.parts.body?.content !== undefined) return "content";
    return op.parts.sections !== undefined ? "fields" : "none";
  }
  switch (story.kind) {
    case "header":
      return op.parts.headers !== undefined ? "content" : "none";
    case "footer":
      return op.parts.footers !== undefined ? "content" : "none";
    case "footnote":
      return op.parts.footnotes !== undefined ? "content" : "none";
    case "endnote":
      return op.parts.endnotes !== undefined ? "content" : "none";
    default: {
      const unreachable: never = story;
      return unreachable;
    }
  }
};

/** Check the producer's declared touched set, never infer it from observed differences. */
export const localityStepFailures = ({ before, op, edit }: OpSequenceStep): string[] => {
  const failures: string[] = [];
  const modified = new Set(edit.touched.modified.map(idKey));
  const inserted = new Set(edit.touched.inserted.map(idKey));
  const removed = new Set(edit.touched.removed.map(idKey));
  const touched = new Set([...modified, ...inserted, ...removed]);
  const watermarkOwnsTouched = (story: OpStory, paragraphId: string) => {
    if (op.type !== DOCUMENT_OP_TYPES.SET_DOCUMENT_WATERMARK) return false;
    const id = idKey(paragraphId);
    if (story === OP_STORIES.MAIN)
      return before.package.document.content.some(
        (block) =>
          block.type === "paragraph" &&
          block.sectionProperties !== undefined &&
          block.paraId !== undefined &&
          idKey(block.paraId) === id,
      );
    if (story.kind !== "header") return false;
    if (
      [...op.coverage, ...op.hosts].some(
        ({ rId, paraId }) => rId === story.rId && idKey(paraId) === id,
      )
    )
      return true;
    if (op.change.kind !== "remove") return false;
    const header = before.package.headers?.get(story.rId);
    const hostId = header ? watermarkHostId(header) : undefined;
    return hostId !== undefined && idKey(hostId) === id;
  };
  for (const document of [before, edit.document]) {
    for (const story of documentStories(document)) {
      if (
        op.type !== DOCUMENT_OP_TYPES.SET_DOCUMENT_WATERMARK &&
        storyOwnership({ op, story, original: before }) !== "none"
      )
        continue;
      if (
        storyParagraphs(storyBody(document, story)).some(
          ({ paragraph }) =>
            paragraph.paraId !== undefined &&
            touched.has(idKey(paragraph.paraId)) &&
            !watermarkOwnsTouched(story, paragraph.paraId),
        )
      )
        failures.push(`${op.type} declared a touched block outside its addressed story`);
    }
  }
  const beforeParagraphs = new Map(
    documentStories(before)
      .flatMap((story) => storyParagraphs(storyBody(before, story)))
      .flatMap(({ paragraph }) =>
        paragraph.paraId === undefined ? [] : [[idKey(paragraph.paraId), paragraph] as const],
      ),
  );
  const afterParagraphs = new Map(
    documentStories(edit.document)
      .flatMap((story) => storyParagraphs(storyBody(edit.document, story)))
      .flatMap(({ paragraph }) =>
        paragraph.paraId === undefined ? [] : [[idKey(paragraph.paraId), paragraph] as const],
      ),
  );
  for (const [id, paragraph] of beforeParagraphs) {
    const after = afterParagraphs.get(id);
    if (after === undefined) {
      if (!removed.has(id)) failures.push(`${op.type} removed an undeclared block`);
    } else if (!touched.has(id) && !sameOpModel(paragraph, after)) {
      failures.push(`${op.type} changed an untouched paragraph`);
    }
  }
  for (const id of afterParagraphs.keys()) {
    if (!beforeParagraphs.has(id) && !inserted.has(id))
      failures.push(`${op.type} inserted an undeclared block`);
  }
  const scopedBefore = withoutOwnedRecords({
    document: before,
    original: before,
    counterpart: edit.document,
    op,
  });
  const scopedAfter = withoutOwnedRecords({
    document: edit.document,
    original: before,
    counterpart: before,
    op,
  });
  // Section-field ownership does not grant ownership of the paragraph carrying it.
  const ownsMainContent =
    storyOwnership({ op, story: OP_STORIES.MAIN, original: before }) === "content";
  const projectedTouched = new Set(touched);
  if (!ownsMainContent)
    for (const document of [before, edit.document])
      for (const { paragraph } of storyParagraphs(document.package.document))
        if (paragraph.paraId !== undefined) projectedTouched.delete(idKey(paragraph.paraId));
  const originalUntouched = projected(scopedBefore, projectedTouched);
  const editedUntouched = projected(scopedAfter, projectedTouched);
  if (!sameOpModel(originalUntouched, editedUntouched))
    failures.push(`${op.type} changed records outside its declared story and section fields`);
  if (
    !sameOpModel(
      originalUntouched.package.document.content,
      editedUntouched.package.document.content,
    )
  )
    failures.push(`${op.type} changed an untouched block or container`);
  if (serializeOpDocument(originalUntouched) !== serializeOpDocument(editedUntouched))
    failures.push(`${op.type} changed serialization outside its touched blocks`);
  return [...new Set(failures)];
};

type SerializedLocalityOptions = {
  sequence: Pick<OpSequence, "steps">;
  control: Map<string, Uint8Array>;
  edited: Map<string, Uint8Array>;
  documentPart: string;
};
export const serializedLocalityFailures = ({
  sequence,
  control,
  edited,
  documentPart,
}: SerializedLocalityOptions): string[] => {
  const failures: string[] = [];
  const ownedPaths = new Set<string>();
  const ownedRelationshipIds = new Set<string>();
  const ownedRelationshipTypes = new Set<string>();
  let lifecycle = false;
  for (const { op, before } of sequence.steps) {
    if (op.type === DOCUMENT_OP_TYPES.SET_DOCUMENT_WATERMARK) {
      lifecycle = true;
      ownedPaths.add(documentPart);
      // Watermarks own header decorations; unrelated header content stays guarded by model locality.
      for (const rId of before.package.headers?.keys() ?? []) ownedRelationshipIds.add(rId);
      for (const { rId } of op.coverage) ownedRelationshipIds.add(rId);
    }
    if (isCommentOp(op)) {
      lifecycle = true;
      ownedPaths.add("word/comments.xml");
      ownedPaths.add("word/commentsExtended.xml");
      for (const { type } of Object.values(COMMENT_PART_RELATIONSHIPS))
        ownedRelationshipTypes.add(type);
      for (const ownedStory of commentStories(op, before)) {
        if (ownedStory === OP_STORIES.MAIN) ownedPaths.add(documentPart);
        else if (ownedStory.kind === "header" || ownedStory.kind === "footer")
          ownedRelationshipIds.add(ownedStory.rId);
        else if (ownedStory.kind === "footnote")
          ownedRelationshipTypes.add(RELATIONSHIP_TYPES.footnotes);
        else if (ownedStory.kind === "endnote")
          ownedRelationshipTypes.add(RELATIONSHIP_TYPES.endnotes);
      }
    }
    if (op.type === DOCUMENT_OP_TYPES.SET_PACKAGE_RESOURCES) {
      const resourceRelationships = {
        styles: RELATIONSHIP_TYPES.styles,
        numbering: RELATIONSHIP_TYPES.numbering,
      } satisfies Record<keyof PackageResources, string>;
      for (const [key, relationshipType] of Object.entries(resourceRelationships)) {
        if (sameOpModel(Reflect.get(op.expected, key), Reflect.get(op.resources, key))) continue;
        lifecycle = true;
        ownedRelationshipTypes.add(relationshipType);
      }
      if (op.media.expected !== op.media.next || op.media.entries.length > 0) {
        lifecycle = true;
        ownedRelationshipTypes.add(RELATIONSHIP_TYPES.image);
      }
      if (
        op.relationships.expected !== op.relationships.next ||
        op.relationships.entries.length > 0
      )
        lifecycle = true;
      for (const { key } of op.relationships.entries) ownedRelationshipIds.add(key);
    }
    if (
      op.type === DOCUMENT_OP_TYPES.CREATE_NUMBERING_INSTANCE ||
      op.type === DOCUMENT_OP_TYPES.DELETE_NUMBERING_INSTANCE
    ) {
      lifecycle = true;
      ownedRelationshipTypes.add(RELATIONSHIP_TYPES.numbering);
    }
    if (op.type === DOCUMENT_OP_TYPES.RESTORE_STORY_PARTS) {
      if (op.parts.body !== undefined || op.parts.sections !== undefined)
        ownedPaths.add(documentPart);
      const restoredRelationships = {
        headers: RELATIONSHIP_TYPES.header,
        footers: RELATIONSHIP_TYPES.footer,
        footnotes: RELATIONSHIP_TYPES.footnotes,
        endnotes: RELATIONSHIP_TYPES.endnotes,
        settings: RELATIONSHIP_TYPES.settings,
      } satisfies Record<keyof Omit<StoryParts, "body" | "sections" | "undefinedFields">, string>;
      for (const [key, relationshipType] of Object.entries(restoredRelationships)) {
        if (
          Object.entries(op.parts).some(
            ([ownedKey, value]) => ownedKey === key && value !== undefined,
          )
        ) {
          ownedRelationshipTypes.add(relationshipType);
          lifecycle = true;
        }
      }
      if (op.parts.body?.comments !== undefined) {
        ownedRelationshipTypes.add(RELATIONSHIP_TYPES.comments);
        lifecycle = true;
      }
    }
    if (op.type === DOCUMENT_OP_TYPES.SET_SECTION_ENDPOINT) ownedPaths.add(documentPart);
    const story = addressedStory(op);
    if (story === OP_STORIES.MAIN) ownedPaths.add(documentPart);
    if (story !== undefined && story !== OP_STORIES.MAIN) {
      if (story.kind === "header" || story.kind === "footer") ownedRelationshipIds.add(story.rId);
      if (story.kind === "footnote") ownedRelationshipTypes.add(RELATIONSHIP_TYPES.footnotes);
      if (story.kind === "endnote") ownedRelationshipTypes.add(RELATIONSHIP_TYPES.endnotes);
    }
    const ownership = lifecycleOwnership(op);
    if (!ownership) continue;
    if (ownership.sectionIndex !== undefined) ownedPaths.add(documentPart);
    lifecycle = true;
    const ownedStory = ownership.story;
    if (ownedStory?.kind === "header" || ownedStory?.kind === "footer")
      ownedRelationshipIds.add(ownedStory.rId);
    if (ownedStory?.kind === "footnote") ownedRelationshipTypes.add(RELATIONSHIP_TYPES.footnotes);
    if (ownedStory?.kind === "endnote") ownedRelationshipTypes.add(RELATIONSHIP_TYPES.endnotes);
    if (ownership.settingsEven) ownedRelationshipTypes.add(RELATIONSHIP_TYPES.settings);
  }
  const relationshipsPath = path.posix.join(
    path.posix.dirname(documentPart),
    "_rels",
    `${path.posix.basename(documentPart)}.rels`,
  );
  const decode = (parts: Map<string, Uint8Array>, part: string) =>
    new TextDecoder().decode(parts.get(part) ?? new Uint8Array());
  const controlRelationships = parseRelationships(decode(control, relationshipsPath));
  const editedRelationships = parseRelationships(decode(edited, relationshipsPath));
  const relationshipAttributes = (parts: Map<string, Uint8Array>) =>
    new Map(
      getChildElements(parseXmlDocument(decode(parts, relationshipsPath))).flatMap((element) => {
        const id = getAttribute(element, null, "Id");
        return id ? [[id, getAttributes(element)] as const] : [];
      }),
    );
  if (sequence.steps.some(({ op }) => isCommentOp(op))) {
    const baselineAttributes = relationshipAttributes(control);
    const editedAttributes = relationshipAttributes(edited);
    for (const [id, relation] of controlRelationships) {
      if (
        !editedRelationships.has(id) ||
        !Object.values(COMMENT_PART_RELATIONSHIPS).some(({ type }) => type === relation.type)
      )
        continue;
      if (!sameOpModel(baselineAttributes.get(id), editedAttributes.get(id)))
        failures.push(
          `sequence changed an existing comment relationship payload: ${relationshipsPath}#${id}`,
        );
    }
  }
  if (sequence.steps.some(({ op }) => op.type === DOCUMENT_OP_TYPES.SET_DOCUMENT_WATERMARK)) {
    const lifecycleHeaderIds = new Set(
      sequence.steps.flatMap(({ op }) => explicitHeaderRelationshipIds(op)),
    );
    const baselineAttributes = relationshipAttributes(control);
    const editedAttributes = relationshipAttributes(edited);
    for (const [id, relationship] of controlRelationships) {
      if (relationship.type !== RELATIONSHIP_TYPES.header || lifecycleHeaderIds.has(id)) continue;
      if (!sameOpModel(baselineAttributes.get(id), editedAttributes.get(id)))
        failures.push(
          `sequence changed an existing header relationship payload: ${relationshipsPath}#${id}`,
        );
    }
  }
  const ownsRelationship = (relationship: Relationship) =>
    ownedRelationshipIds.has(relationship.id) || ownedRelationshipTypes.has(relationship.type);
  for (const relationship of [...controlRelationships.values(), ...editedRelationships.values()]) {
    if (ownsRelationship(relationship) && relationship.targetMode !== "External")
      ownedPaths.add(
        path.posix.resolve("/", path.posix.dirname(documentPart), relationship.target).slice(1),
      );
  }
  if (lifecycle) {
    const remaining = (relationships: typeof controlRelationships) =>
      new Map([...relationships].filter(([, relationship]) => !ownsRelationship(relationship)));
    if (!sameOpModel(remaining(controlRelationships), remaining(editedRelationships)))
      failures.push(`sequence changed unrelated package relationships: ${relationshipsPath}`);
    const contentTypesPath = "[Content_Types].xml";
    const unownedContentTypes = (parts: Map<string, Uint8Array>) => {
      const root = parseXmlDocument(decode(parts, contentTypesPath));
      if (
        !root ||
        getLocalName(root.name) !== "Types" ||
        getNamespaceUri(root) !== "http://schemas.openxmlformats.org/package/2006/content-types"
      ) {
        failures.push("sequence has invalid package content types: [Content_Types].xml");
        return undefined;
      }
      return getChildElements(root)
        .filter((element) => {
          const partName = getAttribute(element, null, "PartName");
          return (
            partName === null ||
            partName === undefined ||
            !ownedPaths.has(partName.replace(/^\//u, ""))
          );
        })
        .map((element) => ({
          name: getLocalName(element.name),
          attributes: getAttributes(element),
        }))
        .sort((left, right) => JSON.stringify(left).localeCompare(JSON.stringify(right)));
    };
    if (!sameOpModel(unownedContentTypes(control), unownedContentTypes(edited)))
      failures.push("sequence changed unrelated package content types: [Content_Types].xml");
    ownedPaths.add(relationshipsPath);
    ownedPaths.add(contentTypesPath);
  }
  for (const part of [...new Set([...control.keys(), ...edited.keys()])].toSorted()) {
    if (ownedPaths.has(part)) continue;
    if (!sameOpModel(control.get(part), edited.get(part))) {
      failures.push(`sequence changed unrelated serialized part: ${part}`);
      break;
    }
  }
  return failures;
};

type SerializedLocalityStepOptions = Omit<SerializedLocalityOptions, "sequence"> & {
  step: OpSequenceStep;
};

/** Attribute package locality to one forward operation, with only that operation's ownership. */
export const serializedLocalityStepFailures = ({
  step,
  ...options
}: SerializedLocalityStepOptions): string[] => {
  const first = serializedLocalityFailures({
    ...options,
    sequence: { steps: [step] },
  })
    .map((message) => ({ message, part: message.slice(message.indexOf(": ") + 2) }))
    .toSorted((left, right) => (left.part < right.part ? -1 : Number(left.part > right.part)))
    .at(0);
  return first === undefined ? [] : [`${step.op.type} ${first.message.replace(/^sequence /u, "")}`];
};

export const runOpLocalityInvariant = async (
  input: CorpusInvariantInput,
): Promise<CorpusInvariantOutcome> => {
  const timings = {};
  const outcome = await timeStage(timings, "sequences", () =>
    Result.tryPromise({
      try: async () => {
        const document = await prepareOpDocument(input);
        const seed = seedFromBytes(input.bytes);
        const failures: string[] = [];
        const control = await serializedOpParts(document);
        for (const salt of OP_SEQUENCE_SEEDS) {
          const sequence = generateOpSequence(document, seed ^ salt);
          failures.push(...sequence.mutations.map((type) => `${type} mutated its input document`));
          // oxlint-disable-next-line no-await-in-loop -- each sequence is checked against one shared control save
          const edited = await serializedOpParts(sequence.document);
          const compoundFailures = serializedLocalityFailures({
            sequence,
            control,
            edited,
            documentPart: input.documentPart,
          });
          let classified = false;
          for (const step of sequence.steps) {
            const modelFailures = localityStepFailures(step);
            if (compoundFailures.length === 0 && modelFailures.length === 0) continue;
            // oxlint-disable-next-line no-await-in-loop -- extra package saves classify only observed failures
            const [beforeParts, editedParts] = await Promise.all([
              serializedOpParts(step.before),
              serializedOpParts(step.edit.document),
            ]);
            const stepFailures = serializedLocalityStepFailures({
              step,
              control: beforeParts,
              edited: editedParts,
              documentPart: input.documentPart,
            });
            classified ||= stepFailures.length > 0;
            failures.push(...stepFailures);
            const part = firstDifferingOpPart({ control: beforeParts, edited: editedParts });
            failures.push(
              ...modelFailures.map((message) => `${message}; part: ${part ?? "model-only"}`),
            );
          }
          if (!classified)
            for (const type of new Set(sequence.steps.map(({ op }) => op.type)))
              failures.push(
                ...compoundFailures.map(
                  (message) => `${type} ${message.replace(/^sequence /u, "composition ")}`,
                ),
              );
        }
        return [...new Set(failures)];
      },
      catch: (cause: unknown) => cause,
    }),
  );
  if (outcome.isErr())
    return opErrorOutcome({ invariant: INVARIANT, error: outcome.error, timings });
  return {
    status: "evaluated",
    timings,
    failures: outcome.value.map((message) => failureFromAssertion(INVARIANT, message)),
  };
};
