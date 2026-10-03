import { panic } from "better-result";
/** Reproducible, bounded operation sequences over seeded corpus models. */
import { isDeepStrictEqual } from "node:util";
import type { Document, Paragraph } from "../../../packages/docx-core/src/model/document";
import {
  applyDocumentOp,
  normalizeForOps,
  validateOpsDocument,
  DOCUMENT_OP_TYPES,
  OP_STORIES,
  INHERIT_RUN_PROPS,
  REVISION_DECISIONS,
  type DocumentOp,
  type AppliedDocumentOp,
  type NoteStory,
  sectionPropertiesAt,
} from "../../../packages/docx-core/src/ops/documentOps";
import { cloneDocumentWithParagraphPropertySources } from "@stll/folio-core/docx/document-clone";
import { ensureParaIds } from "@stll/folio-core/docx/ensureParaIds";
import { repackDocx } from "@stll/folio-core/docx/rezip";
import { unzipDocx } from "@stll/folio-core/docx/unzip";
import { parseDocx } from "@stll/folio-core/docx/parser";
import { serializeDocument } from "@stll/folio-core/docx/serializer/documentSerializer";
import { storyParagraphs, blockListAt } from "../../../packages/docx-core/src/ops/blocks";
import { leafSpans } from "../../../packages/docx-core/src/ops/leaves";
import { paragraphLength } from "../../../packages/docx-core/src/ops/offsets";
import {
  packageIdentityKeys,
  packageParagraphIds,
  idKey,
} from "../../../packages/docx-core/src/ops/ids";
import type { CorpusInvariantInput } from "./contract";

export const OP_SEQUENCE_SEEDS = [0x17a3, 0x5b91, 0xcf27] as const;

/** Exact records, including captures and binary contents; undefined is wire absence. */
export const exactOpModel = (value: unknown): unknown => {
  if (value instanceof Uint8Array) return value.slice();
  if (value instanceof ArrayBuffer) return value.slice(0);
  if (value instanceof Date) return new Date(value);
  if (Array.isArray(value)) return value.map(exactOpModel);
  if (value instanceof Map)
    return new Map([...value].map(([key, entry]) => [key, exactOpModel(entry)]));
  if (
    typeof value !== "object" ||
    value === null ||
    Object.getPrototypeOf(value) !== Object.prototype
  )
    return value;
  return Object.fromEntries(
    Object.entries(value)
      .filter(([, entry]) => entry !== undefined)
      .map(([key, entry]) => [key, exactOpModel(entry)]),
  );
};

export const sameOpModel = (left: unknown, right: unknown): boolean =>
  isDeepStrictEqual(exactOpModel(left), exactOpModel(right));

export const serializeOpDocument = (document: Document): string => serializeDocument(document);

/** ZIP metadata/compression is transport; compare all uncompressed part bytes. */
export const serializedOpParts = async (document: Document): Promise<Map<string, Uint8Array>> => {
  const saved = await repackDocx(document, {
    updateModifiedDate: false,
    onDiagnostic: ({ type, part }) => panic(`${type}: ${part}`),
  });
  const raw = await unzipDocx(saved, { extractAllXml: true });
  return new Map(
    await Promise.all(
      Object.entries(raw.originalZip.files)
        .filter(([, file]) => !file.dir)
        .map(async ([path, file]) => [path, await file.async("uint8array")] as const),
    ),
  );
};

export const seedFromBytes = (bytes: Uint8Array): number => {
  let hash = 2166136261;
  for (const byte of bytes) hash = Math.imul(hash ^ byte, 16777619);
  return hash >>> 0;
};

const prepared = new WeakMap<Document, Promise<Document>>();
/** Establish the public operations contract before measuring edits, once per parse. */
export const prepareOpDocument = (input: CorpusInvariantInput): Promise<Document> => {
  const existing = prepared.get(input.parsed);
  if (existing !== undefined) return existing;
  const pending = (async () => {
    const seeded = await ensureParaIds(input.bytes);
    const parsed = seeded.alreadyComplete
      ? input.parsed
      : await parseDocx(Uint8Array.from(seeded.docx).buffer, { preloadFonts: false });
    const document = normalizeForOps(parsed);
    const valid = validateOpsDocument(document);
    if (valid.isErr()) throw valid.error;
    return document;
  })();
  prepared.set(input.parsed, pending);
  return pending;
};

const randomFor = (seed: number) => {
  let state = seed >>> 0;
  return (limit: number): number => {
    state = (Math.imul(state, 1664525) + 1013904223) >>> 0;
    return limit === 0 ? 0 : state % limit;
  };
};

export const OP_SEQUENCE_FAMILIES = [
  DOCUMENT_OP_TYPES.INSERT_TEXT,
  DOCUMENT_OP_TYPES.INSERT_CONTENT,
  DOCUMENT_OP_TYPES.DELETE_RANGE,
  DOCUMENT_OP_TYPES.SET_RUN_PROPS,
  DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS,
  DOCUMENT_OP_TYPES.SPLIT_BLOCK,
  DOCUMENT_OP_TYPES.JOIN_BLOCKS,
  DOCUMENT_OP_TYPES.INSERT_BLOCKS,
  DOCUMENT_OP_TYPES.DELETE_BLOCKS,
  DOCUMENT_OP_TYPES.INSERT_TABLE,
  DOCUMENT_OP_TYPES.DELETE_TABLE,
  DOCUMENT_OP_TYPES.SET_CONTAINER_BLOCKS,
  DOCUMENT_OP_TYPES.INSERT_ROW,
  DOCUMENT_OP_TYPES.DELETE_ROW,
  DOCUMENT_OP_TYPES.RESOLVE_REVISION,
  DOCUMENT_OP_TYPES.CREATE_HEADER_FOOTER,
  DOCUMENT_OP_TYPES.REMOVE_HEADER_FOOTER,
  DOCUMENT_OP_TYPES.ADD_NOTE,
  DOCUMENT_OP_TYPES.REMOVE_NOTE,
  DOCUMENT_OP_TYPES.SET_SECTION_PROPS,
] as const;
export const OP_SEQUENCE_LENGTH = OP_SEQUENCE_FAMILIES.length;

/** Every schema member needs a generator decision when the operations API grows. */
export const OP_GENERATOR_ROLES = {
  [DOCUMENT_OP_TYPES.CREATE_HEADER_FOOTER]: "generated",
  [DOCUMENT_OP_TYPES.REMOVE_HEADER_FOOTER]: "generated",
  [DOCUMENT_OP_TYPES.ADD_NOTE]: "generated",
  [DOCUMENT_OP_TYPES.REMOVE_NOTE]: "generated",
  [DOCUMENT_OP_TYPES.SET_SECTION_PROPS]: "generated",
  [DOCUMENT_OP_TYPES.RESTORE_STORY_PARTS]: "inverse",
  [DOCUMENT_OP_TYPES.INSERT_TEXT]: "generated",
  [DOCUMENT_OP_TYPES.INSERT_CONTENT]: "generated",
  [DOCUMENT_OP_TYPES.DELETE_RANGE]: "generated",
  [DOCUMENT_OP_TYPES.SET_RUN_PROPS]: "generated",
  [DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS]: "generated",
  [DOCUMENT_OP_TYPES.SPLIT_BLOCK]: "generated",
  [DOCUMENT_OP_TYPES.JOIN_BLOCKS]: "generated",
  [DOCUMENT_OP_TYPES.INSERT_BLOCKS]: "generated",
  [DOCUMENT_OP_TYPES.DELETE_BLOCKS]: "generated",
  [DOCUMENT_OP_TYPES.INSERT_TABLE]: "generated",
  [DOCUMENT_OP_TYPES.DELETE_TABLE]: "generated",
  [DOCUMENT_OP_TYPES.SET_CONTAINER_BLOCKS]: "generated",
  [DOCUMENT_OP_TYPES.INSERT_ROW]: "generated",
  [DOCUMENT_OP_TYPES.DELETE_ROW]: "generated",
  [DOCUMENT_OP_TYPES.RESOLVE_REVISION]: "generated",
  [DOCUMENT_OP_TYPES.SPLIT_INLINE]: "inverse",
  [DOCUMENT_OP_TYPES.JOIN_INLINE]: "inverse",
  [DOCUMENT_OP_TYPES.REPLACE_BLOCKS]: "inverse",
  [DOCUMENT_OP_TYPES.SET_PARAGRAPH_REVIEW]: "inverse",
  [DOCUMENT_OP_TYPES.REPLACE_INLINE]: "inverse",
  [DOCUMENT_OP_TYPES.SET_TABLE_ROWS]: "inverse",
} as const satisfies Record<DocumentOp["type"], "generated" | "inverse">;

type CandidateOptions = {
  document: Document;
  choose: (limit: number) => number;
  family: (typeof OP_SEQUENCE_FAMILIES)[number];
  revisions: readonly number[];
  tracked: boolean;
};
const candidate = ({
  document,
  choose,
  family,
  revisions,
  tracked,
}: CandidateOptions): DocumentOp | undefined => {
  const paragraphs = storyParagraphs(document.package.document);
  const location = paragraphs.at(choose(paragraphs.length));
  if (location === undefined || location.paragraph.paraId === undefined) return undefined;
  const paragraph = location.paragraph;
  const blockId = location.paragraph.paraId;
  const length = paragraphLength(paragraph);
  const at = { story: OP_STORIES.MAIN, blockId, offset: choose(length + 1) };
  const ids = new Set(packageParagraphIds(document.package).map(idKey));
  let paraId = 1;
  while (ids.has(idKey(paraId.toString(16).padStart(8, "0")))) paraId += 1;
  const newBlockId = paraId.toString(16).padStart(8, "0");
  const taken = new Set(packageIdentityKeys(document.package));
  const freshRevision: number[] = [];
  const freshControl: number[] = [];
  for (let id = 1; freshRevision.length < 64 || freshControl.length < 64; id += 1) {
    if (freshRevision.length < 64 && !taken.has(`revision:${id}`)) freshRevision.push(id);
    if (freshControl.length < 64 && !taken.has(`control:${id}`)) freshControl.push(id);
  }
  const revisionId = freshRevision.at(0);
  if (revisionId === undefined) return undefined;
  const review = tracked
    ? {
        revision: { id: revisionId, author: "Corpus", date: "2026-01-01T00:00:00Z" },
      }
    : {};
  const newIds = { revision: freshRevision.slice(1), control: freshControl };
  const freshParagraph = {
    type: "paragraph",
    paraId: newBlockId,
    content: [{ type: "run", content: [{ type: "text", text: "‸" }] }],
  } as const satisfies Paragraph;
  const end = { ...at, offset: Math.min(length, at.offset + 1 + choose(3)) };
  const sectionCount =
    document.package.document.content.filter(
      (block) => block.type === "paragraph" && block.sectionProperties !== undefined,
    ).length + 1;
  switch (family) {
    case "createHeaderFooter": {
      const sectionIndex = choose(sectionCount);
      const properties = sectionPropertiesAt(document, sectionIndex);
      if (!properties) return undefined;
      const kind = choose(2) === 0 ? "header" : "footer";
      const references =
        kind === "header" ? properties.headerReferences : properties.footerReferences;
      const availableVariants = (["default", "first", "even"] as const).filter(
        (variant) => !references?.some(({ type }) => type === variant),
      );
      const referenceType = availableVariants.at(choose(availableVariants.length));
      if (!referenceType) return undefined;
      let identity = 1;
      const used = (rId: string) =>
        document.package.headers?.has(rId) ||
        document.package.footers?.has(rId) ||
        document.package.relationships?.has(rId);
      while (used(`rIdCorpus${identity}`)) identity += 1;
      return {
        type: DOCUMENT_OP_TYPES.CREATE_HEADER_FOOTER,
        sectionIndex,
        story: { kind, rId: `rIdCorpus${identity}` },
        referenceType,
        content: [freshParagraph],
      };
    }
    case "removeHeaderFooter": {
      const candidates = Array.from({ length: sectionCount }, (_, sectionIndex) => {
        const properties = sectionPropertiesAt(document, sectionIndex);
        const headers =
          properties?.headerReferences
            ?.filter(({ rId }) => document.package.headers?.has(rId))
            .map(
              ({ type, rId }) =>
                ({ sectionIndex, story: { kind: "header", rId }, referenceType: type }) as const,
            ) ?? [];
        const footers =
          properties?.footerReferences
            ?.filter(({ rId }) => document.package.footers?.has(rId))
            .map(
              ({ type, rId }) =>
                ({ sectionIndex, story: { kind: "footer", rId }, referenceType: type }) as const,
            ) ?? [];
        return [...headers, ...footers];
      }).flat();
      const selected = candidates.at(choose(candidates.length));
      return selected ? { type: DOCUMENT_OP_TYPES.REMOVE_HEADER_FOOTER, ...selected } : undefined;
    }
    case "addNote": {
      const kind = choose(2) === 0 ? "footnote" : "endnote";
      const collection =
        kind === "footnote" ? document.package.footnotes : document.package.endnotes;
      let id = 1;
      while (collection?.some((note) => note.id === id)) id += 1;
      return {
        type: DOCUMENT_OP_TYPES.ADD_NOTE,
        at,
        note: { type: kind, id, content: [freshParagraph] },
      };
    }
    case "removeNote": {
      const references = paragraphs.flatMap(({ paragraph: referenceParagraph }) =>
        leafSpans(referenceParagraph.content).flatMap(({ node, before }) => {
          if (node.type !== "footnoteRef" && node.type !== "endnoteRef") return [];
          if (!referenceParagraph.paraId) return [];
          const kind = node.type === "footnoteRef" ? "footnote" : "endnote";
          const notes =
            kind === "footnote" ? document.package.footnotes : document.package.endnotes;
          if (!notes?.some(({ id }) => id === node.id)) return [];
          return [
            {
              at: {
                story: OP_STORIES.MAIN,
                blockId: referenceParagraph.paraId,
                offset: before.offset,
              },
              story: { kind, id: node.id } satisfies NoteStory,
            },
          ];
        }),
      );
      const selected = references.at(choose(references.length));
      return selected ? { type: DOCUMENT_OP_TYPES.REMOVE_NOTE, ...selected } : undefined;
    }
    case "setSectionProps": {
      const sectionIndex = choose(sectionCount);
      const properties = sectionPropertiesAt(document, sectionIndex);
      if (!properties) return undefined;
      return {
        type: DOCUMENT_OP_TYPES.SET_SECTION_PROPS,
        sectionIndex,
        patch: {
          marginTop: properties.marginTop === 720 ? 1440 : 720,
          evenAndOddHeaders: ([true, false, null] as const).at(choose(3)) ?? null,
        },
      };
    }
    case "insertText":
      return {
        type: DOCUMENT_OP_TYPES.INSERT_TEXT,
        at,
        text: ["x", "é", "🧭"][choose(3)] ?? "x",
        runProps: INHERIT_RUN_PROPS,
        newIds,
        ...review,
      };
    case "insertContent":
      return {
        type: DOCUMENT_OP_TYPES.INSERT_CONTENT,
        at,
        slice: { content: freshParagraph.content, openStart: 0, openEnd: 0 },
        newIds,
        ...review,
      };
    case "deleteRange":
      return { type: DOCUMENT_OP_TYPES.DELETE_RANGE, from: at, to: end, newIds, ...review };
    case "setRunProps":
      return {
        type: DOCUMENT_OP_TYPES.SET_RUN_PROPS,
        from: at,
        to: end,
        patch: { bold: choose(2) === 0 },
        newIds,
        ...review,
      };
    case "setParagraphProps":
      return {
        type: DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS,
        story: OP_STORIES.MAIN,
        blockId,
        patch: { alignment: paragraph.formatting?.alignment === "center" ? "left" : "center" },
        ...review,
      };
    case "splitBlock":
      return { type: DOCUMENT_OP_TYPES.SPLIT_BLOCK, at, newBlockId, newIds, ...review };
    case "joinBlocks": {
      const next = blockListAt(document.package.document.content, location.list).at(
        location.index + 1,
      );
      if (next?.type !== "paragraph" || next.paraId === undefined) return undefined;
      return {
        type: DOCUMENT_OP_TYPES.JOIN_BLOCKS,
        story: OP_STORIES.MAIN,
        blockId,
        nextBlockId: next.paraId,
        newIds,
        ...review,
      };
    }
    case "insertBlocks":
      return {
        type: DOCUMENT_OP_TYPES.INSERT_BLOCKS,
        story: OP_STORIES.MAIN,
        at: { type: "before", blockId },
        blocks: [freshParagraph],
        newIds,
        ...review,
      };
    case "deleteBlocks":
      return {
        type: DOCUMENT_OP_TYPES.DELETE_BLOCKS,
        story: OP_STORIES.MAIN,
        blockIds: [blockId],
        newIds,
        ...review,
      };
    case "insertTable":
      return {
        type: DOCUMENT_OP_TYPES.INSERT_TABLE,
        story: OP_STORIES.MAIN,
        at: { type: "before", blockId },
        table: {
          type: "table",
          rows: [{ type: "tableRow", cells: [{ type: "tableCell", content: [freshParagraph] }] }],
        },
      };
    case "deleteTable":
      if (!location.list.some((step) => step.kind === "tableCell")) return undefined;
      return { type: DOCUMENT_OP_TYPES.DELETE_TABLE, story: OP_STORIES.MAIN, blockId };
    case "setContainerBlocks": {
      const expected = blockListAt(document.package.document.content, location.list);
      const blocks = [...expected];
      blocks.splice(location.index, 0, {
        type: "table",
        rows: [{ type: "tableRow", cells: [{ type: "tableCell", content: [freshParagraph] }] }],
      });
      return {
        type: DOCUMENT_OP_TYPES.SET_CONTAINER_BLOCKS,
        story: OP_STORIES.MAIN,
        blockId,
        expected,
        blocks,
      };
    }
    case "insertRow": {
      if (!location.list.some((step) => step.kind === "tableCell")) return undefined;
      return {
        type: DOCUMENT_OP_TYPES.INSERT_ROW,
        story: OP_STORIES.MAIN,
        blockId,
        at: 0,
        row: { type: "tableRow", cells: [{ type: "tableCell", content: [freshParagraph] }] },
        newIds,
        ...review,
      };
    }
    case "deleteRow":
      if (!location.list.some((step) => step.kind === "tableCell")) return undefined;
      return {
        type: DOCUMENT_OP_TYPES.DELETE_ROW,
        story: OP_STORIES.MAIN,
        blockId,
        newIds,
        ...review,
      };
    case "resolveRevision":
      if (revisions.length === 0) return undefined;
      return {
        type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
        story: OP_STORIES.MAIN,
        revisionIds: revisions,
        decision: choose(2) === 0 ? REVISION_DECISIONS.ACCEPT : REVISION_DECISIONS.REJECT,
      };
    default: {
      const unreachable: never = family;
      return unreachable;
    }
  }
};

export type OpSequenceStep = { before: Document; op: DocumentOp; edit: AppliedDocumentOp };
export type OpSequence = {
  original: Document;
  originalModel: unknown;
  originalXml: string;
  document: Document;
  steps: OpSequenceStep[];
  inverse: DocumentOp[];
  mutations: string[];
  refusals: string[];
};

/** A schedule covers every family; the file seed varies positions, modes and order. */
export const generateOpSequence = (document: Document, seed: number): OpSequence => {
  const original = cloneDocumentWithParagraphPropertySources(document);
  const originalModel = exactOpModel(original);
  const originalXml = serializeOpDocument(original);
  const choose = randomFor(seed);
  const shift = choose(OP_SEQUENCE_FAMILIES.length);
  let current = original;
  const steps: OpSequenceStep[] = [];
  const revisions: number[] = [];
  const mutations: string[] = [];
  const refusals: string[] = [];
  const inverseGroups: (readonly DocumentOp[])[] = [];
  for (let index = 0; index < OP_SEQUENCE_LENGTH; index += 1) {
    const family = OP_SEQUENCE_FAMILIES.at((index + shift) % OP_SEQUENCE_FAMILIES.length);
    if (family === undefined) continue;
    const op = candidate({
      document: current,
      choose,
      family,
      revisions,
      tracked: choose(2) === 0,
    });
    if (op === undefined) continue;
    const before = current;
    const snapshot = exactOpModel(before);
    const applied = applyDocumentOp(before, op);
    if (!isDeepStrictEqual(snapshot, exactOpModel(before))) mutations.push(op.type);
    if (applied.isErr()) {
      refusals.push(`${op.type}:${applied.error.reason}`);
      continue;
    }
    steps.push({ before, op, edit: applied.value });
    inverseGroups.unshift(applied.value.inverse);
    revisions.push(...applied.value.revisions);
    current = applied.value.document;
  }
  return {
    original,
    originalModel,
    originalXml,
    document: current,
    steps,
    inverse: inverseGroups.flat(),
    mutations,
    refusals,
  };
};
