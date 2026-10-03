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
} from "../../../packages/docx-core/src/model/document";
import { DEFAULT_TAB_STOP_TWIPS } from "../../../packages/docx-core/src/model/document";
import { Result } from "better-result";
import {
  documentStories,
  storyBody,
  type DocumentOp,
  DOCUMENT_OP_TYPES,
} from "../../../packages/docx-core/src/ops/documentOps";
import { storyParagraphs } from "../../../packages/docx-core/src/ops/blocks";
import { idKey } from "../../../packages/docx-core/src/ops/ids";
import { failureFromAssertion, failureFromError } from "../corpus-signature";
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
import { generalizePartPath } from "./save-idempotence";

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
const withoutOwnedRecords = ({ document, original, op }: WithoutOwnedRecordsOptions): Document => {
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

const unrelatedModel = (document: Document): unknown => ({
  ...document,
  package: {
    ...document.package,
    document: { ...document.package.document, content: undefined, sections: undefined },
  },
});

const projected = (document: Document, touched: ReadonlySet<string>): Document => {
  const body = {
    ...document.package.document,
    content: untouchedBlocks(document.package.document.content, touched),
  };
  delete body.sections;
  return { ...document, package: { ...document.package, document: body } };
};

/** Check the producer's declared touched set, never infer it from observed differences. */
export const localityStepFailures = ({ before, op, edit }: OpSequenceStep): string[] => {
  const failures: string[] = [];
  const modified = new Set(edit.touched.modified.map(idKey));
  const inserted = new Set(edit.touched.inserted.map(idKey));
  const removed = new Set(edit.touched.removed.map(idKey));
  const touched = new Set([...modified, ...inserted, ...removed]);
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
  if (!sameOpModel(unrelatedModel(scopedBefore), unrelatedModel(scopedAfter)))
    failures.push(`${op.type} changed records outside its declared story and section fields`);
  const originalUntouched = projected(scopedBefore, touched);
  const editedUntouched = projected(scopedAfter, touched);
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
  sequence: OpSequence;
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
  const ownedPaths = new Set([documentPart]);
  const ownedRelationshipIds = new Set<string>();
  const ownedRelationshipTypes = new Set<string>();
  let lifecycle = false;
  for (const { op } of sequence.steps) {
    const ownership = lifecycleOwnership(op);
    if (!ownership) continue;
    lifecycle = true;
    const story = ownership.story;
    if (story?.kind === "header" || story?.kind === "footer") ownedRelationshipIds.add(story.rId);
    if (story?.kind === "footnote") ownedRelationshipTypes.add(RELATIONSHIP_TYPES.footnotes);
    if (story?.kind === "endnote") ownedRelationshipTypes.add(RELATIONSHIP_TYPES.endnotes);
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
  if (lifecycle) {
    for (const relationship of [
      ...controlRelationships.values(),
      ...editedRelationships.values(),
    ]) {
      if (ownsRelationship(relationship) && relationship.targetMode !== "External")
        ownedPaths.add(
          path.posix.resolve("/", path.posix.dirname(documentPart), relationship.target).slice(1),
        );
    }
    const remaining = (relationships: typeof controlRelationships) =>
      new Map([...relationships].filter(([, relationship]) => !ownsRelationship(relationship)));
    if (!sameOpModel(remaining(controlRelationships), remaining(editedRelationships)))
      failures.push("sequence changed unrelated package relationships");
    const contentTypesPath = "[Content_Types].xml";
    const unownedContentTypes = (parts: Map<string, Uint8Array>) => {
      const root = parseXmlDocument(decode(parts, contentTypesPath));
      if (
        !root ||
        getLocalName(root.name) !== "Types" ||
        getNamespaceUri(root) !== "http://schemas.openxmlformats.org/package/2006/content-types"
      ) {
        failures.push("sequence has invalid package content types");
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
      failures.push("sequence changed unrelated package content types");
    ownedPaths.add(relationshipsPath);
    ownedPaths.add(contentTypesPath);
  }
  for (const part of new Set([...control.keys(), ...edited.keys()])) {
    if (ownedPaths.has(part)) continue;
    if (!sameOpModel(control.get(part), edited.get(part)))
      failures.push(`sequence changed unrelated serialized part: ${generalizePartPath(part)}`);
  }
  return failures;
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
          for (const step of sequence.steps) failures.push(...localityStepFailures(step));
          // oxlint-disable-next-line no-await-in-loop -- each sequence is checked against one shared control save
          const edited = await serializedOpParts(sequence.document);
          failures.push(
            ...serializedLocalityFailures({
              sequence,
              control,
              edited,
              documentPart: input.documentPart,
            }),
          );
        }
        return [...new Set(failures)];
      },
      catch: (cause: unknown) => cause,
    }),
  );
  return {
    timings,
    failures: outcome.isErr()
      ? [failureFromError(INVARIANT, outcome.error)]
      : outcome.value.map((message) => failureFromAssertion(INVARIANT, message)),
  };
};
