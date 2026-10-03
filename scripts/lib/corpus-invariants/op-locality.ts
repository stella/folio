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
} from "../../../packages/docx-core/src/model/document";
import { DEFAULT_TAB_STOP_TWIPS } from "../../../packages/docx-core/src/model/document";
import { Result } from "better-result";
import {
  documentStories,
  storyBody,
  type DocumentOp,
  DOCUMENT_OP_TYPES,
  OP_STORIES,
  sameStory,
  type OpStory,
} from "../../../packages/docx-core/src/ops/documentOps";
import type { StoryParts } from "../../../packages/docx-core/src/ops/types";
import { storyParagraphs } from "../../../packages/docx-core/src/ops/blocks";
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

type WithoutOwnedRecordsOptions = { document: Document; original: Document; op: DocumentOp };

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
  } satisfies Record<keyof Omit<StoryParts, "body" | "sections">, () => void>;
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

const withoutOwnedRecords = ({ document, original, op }: WithoutOwnedRecordsOptions): Document => {
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

const ownsStoryContent = (op: DocumentOp, story: OpStory): boolean => {
  const addressed = addressedStory(op);
  if (addressed !== undefined && sameStory(addressed, story)) return true;
  const lifecycle = lifecycleOwnership(op);
  if (lifecycle !== undefined) {
    if (lifecycle.story !== undefined && sameStory(lifecycle.story, story)) return true;
    if (lifecycle.sectionIndex !== undefined && story === OP_STORIES.MAIN) return true;
  }
  if (op.type !== DOCUMENT_OP_TYPES.RESTORE_STORY_PARTS) return false;
  if (story === OP_STORIES.MAIN)
    return op.parts.body?.content !== undefined || op.parts.sections !== undefined;
  switch (story.kind) {
    case "header":
      return op.parts.headers !== undefined;
    case "footer":
      return op.parts.footers !== undefined;
    case "footnote":
      return op.parts.footnotes !== undefined;
    case "endnote":
      return op.parts.endnotes !== undefined;
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
  for (const document of [before, edit.document]) {
    for (const story of documentStories(document)) {
      if (ownsStoryContent(op, story)) continue;
      if (
        storyParagraphs(storyBody(document, story)).some(
          ({ paragraph }) => paragraph.paraId !== undefined && touched.has(idKey(paragraph.paraId)),
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
  const scopedBefore = withoutOwnedRecords({ document: before, original: before, op });
  const scopedAfter = withoutOwnedRecords({ document: edit.document, original: before, op });
  // Section-field ownership does not grant ownership of the paragraph carrying it.
  const ownsMainContent =
    addressedStory(op) === OP_STORIES.MAIN ||
    (op.type === DOCUMENT_OP_TYPES.RESTORE_STORY_PARTS && op.parts.body?.content !== undefined);
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
  for (const { op } of sequence.steps) {
    if (op.type === DOCUMENT_OP_TYPES.RESTORE_STORY_PARTS) {
      if (op.parts.body !== undefined || op.parts.sections !== undefined)
        ownedPaths.add(documentPart);
      const restoredRelationships = {
        headers: RELATIONSHIP_TYPES.header,
        footers: RELATIONSHIP_TYPES.footer,
        footnotes: RELATIONSHIP_TYPES.footnotes,
        endnotes: RELATIONSHIP_TYPES.endnotes,
        settings: RELATIONSHIP_TYPES.settings,
      } satisfies Record<keyof Omit<StoryParts, "body" | "sections">, string>;
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
