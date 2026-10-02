import { Result } from "better-result";

import {
  MAX_REVISION_ID,
  type BlockContent,
  type Document,
  type Section,
  type SectionProperties,
} from "../model/document";
import { structurallyEqual } from "./equality";
import { DOCUMENT_OP_REFUSAL_REASONS, DocumentOpRefusal } from "./refusal";
import {
  DOCUMENT_OP_TYPES,
  type CreateNumberingInstanceOp,
  type DeleteNumberingInstanceOp,
  type DocumentOp,
  type NumberingPartState,
  type SetSectionEndpointOp,
  type SectionViewEntry,
} from "./types";
import type { DocumentEdit } from "./edits";

const failed = (
  op: DocumentOp,
  reason: (typeof DOCUMENT_OP_REFUSAL_REASONS)[keyof typeof DOCUMENT_OP_REFUSAL_REASONS],
  message: string,
) => Result.err(new DocumentOpRefusal({ message, reason, opType: op.type }));

const unchanged = (document: Document): Result<DocumentEdit, DocumentOpRefusal> =>
  Result.ok({ document, inverse: [], touched: { modified: [], inserted: [], removed: [] } });

const emptyEdit = (document: Document, inverse: DocumentOp): DocumentEdit => ({
  document,
  inverse: [inverse],
  touched: { modified: [], inserted: [], removed: [] },
});

const numberingStateOf = (document: Document): NumberingPartState => {
  if (!Object.hasOwn(document.package, "numbering")) return { type: "omitted" };
  const numbering = document.package.numbering;
  return numbering === undefined
    ? { type: "undefined" }
    : { type: "definitions", value: numbering };
};

const withNumberingState = (document: Document, state: NumberingPartState): Document => {
  const packageWithoutNumbering = { ...document.package };
  delete packageWithoutNumbering.numbering;
  switch (state.type) {
    case "omitted":
      return { ...document, package: packageWithoutNumbering };
    case "undefined":
      return { ...document, package: { ...packageWithoutNumbering, numbering: undefined } };
    case "definitions":
      return { ...document, package: { ...packageWithoutNumbering, numbering: state.value } };
    default: {
      const unreachable: never = state;
      return unreachable;
    }
  }
};

const packageReferencesNumbering = (value: unknown, numId: number): boolean => {
  if (typeof value !== "object" || value === null) return false;
  if (Array.isArray(value)) return value.some((item) => packageReferencesNumbering(item, numId));
  if (value instanceof Map)
    return [...value.values()].some((item) => packageReferencesNumbering(item, numId));
  if (Reflect.get(value, "kind") === "reference" && Reflect.get(value, "numId") === numId)
    return true;
  return Object.values(value).some((item) => packageReferencesNumbering(item, numId));
};

const applyCreateNumberingInstance = (document: Document, op: CreateNumberingInstanceOp) => {
  const priorState = numberingStateOf(document);
  if (op.expected !== undefined && !structurallyEqual(priorState, op.expected)) {
    return failed(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.STALE,
      "The numbering part differs from the expected state.",
    );
  }
  const prior = document.package.numbering;
  const numbering = prior ?? { abstractNums: [], nums: [] };
  if (!Number.isSafeInteger(op.num.numId) || op.num.numId <= 0 || op.num.numId > MAX_REVISION_ID) {
    return failed(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
      "A numbering id must be a positive safe integer.",
    );
  }
  if (numbering.nums.some(({ numId }) => numId === op.num.numId)) {
    return failed(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.ID_COLLISION,
      `Numbering id ${op.num.numId} already exists.`,
    );
  }
  if (op.abstractNum !== undefined) {
    if (
      !Number.isSafeInteger(op.abstractNum.abstractNumId) ||
      op.abstractNum.abstractNumId < 0 ||
      op.abstractNum.abstractNumId > MAX_REVISION_ID
    ) {
      return failed(
        op,
        DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
        "An abstract numbering id must be a non-negative safe integer.",
      );
    }
    if (op.abstractNum.abstractNumId !== op.num.abstractNumId) {
      return failed(
        op,
        DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
        "The instance and abstract numbering ids do not match.",
      );
    }
    if (
      numbering.abstractNums.some(
        ({ abstractNumId }) => abstractNumId === op.abstractNum?.abstractNumId,
      )
    ) {
      return failed(
        op,
        DOCUMENT_OP_REFUSAL_REASONS.ID_COLLISION,
        `Abstract numbering id ${op.abstractNum.abstractNumId} already exists.`,
      );
    }
  } else if (
    !numbering.abstractNums.some(({ abstractNumId }) => abstractNumId === op.num.abstractNumId)
  ) {
    return failed(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
      `Abstract numbering id ${op.num.abstractNumId} does not exist.`,
    );
  }
  const nextNumbering = {
    ...numbering,
    abstractNums:
      op.abstractNum === undefined
        ? numbering.abstractNums
        : [...numbering.abstractNums, op.abstractNum],
    nums: [...numbering.nums, op.num],
  };
  const next = { ...document, package: { ...document.package, numbering: nextNumbering } };
  const nextState = numberingStateOf(next);
  if (op.restore !== undefined && !structurallyEqual(nextState, op.restore)) {
    return failed(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.STALE,
      "The created numbering data does not match the requested restoration.",
    );
  }
  return Result.ok(
    emptyEdit(next, {
      type: DOCUMENT_OP_TYPES.DELETE_NUMBERING_INSTANCE,
      num: op.num,
      ...(op.abstractNum === undefined ? {} : { abstractNum: op.abstractNum }),
      expected: nextState,
      restore: priorState,
    }),
  );
};

const applyDeleteNumberingInstance = (document: Document, op: DeleteNumberingInstanceOp) => {
  const priorState = numberingStateOf(document);
  if (op.expected !== undefined && !structurallyEqual(priorState, op.expected)) {
    return failed(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.STALE,
      "The numbering part differs from the expected state.",
    );
  }
  const prior = document.package.numbering;
  const instance = prior?.nums.find(({ numId }) => numId === op.num.numId);
  if (prior === undefined || instance === undefined || !structurallyEqual(instance, op.num)) {
    return failed(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.STALE,
      "The numbering instance differs from the expected value.",
    );
  }
  if (packageReferencesNumbering(document.package, op.num.numId)) {
    return failed(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.STALE,
      "The numbering instance is still referenced by a paragraph.",
    );
  }
  const abstract =
    op.abstractNum === undefined
      ? undefined
      : prior.abstractNums.find(
          ({ abstractNumId }) => abstractNumId === op.abstractNum?.abstractNumId,
        );
  if (
    op.abstractNum !== undefined &&
    (abstract === undefined || !structurallyEqual(abstract, op.abstractNum))
  ) {
    return failed(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.STALE,
      "The abstract numbering definition differs from the expected value.",
    );
  }
  if (
    op.abstractNum !== undefined &&
    prior.nums.some(
      ({ numId, abstractNumId }) =>
        numId !== op.num.numId && abstractNumId === op.abstractNum?.abstractNumId,
    )
  ) {
    return failed(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.STALE,
      "The abstract numbering definition is still referenced by another instance.",
    );
  }
  const numbering = {
    ...prior,
    nums: prior.nums.filter(({ numId }) => numId !== op.num.numId),
    abstractNums:
      op.abstractNum === undefined
        ? prior.abstractNums
        : prior.abstractNums.filter(
            ({ abstractNumId }) => abstractNumId !== op.abstractNum?.abstractNumId,
          ),
  };
  const intermediate = { ...document, package: { ...document.package, numbering } };
  const removedState = numberingStateOf(intermediate);
  if (op.restore?.type === "definitions" && !structurallyEqual(removedState, op.restore)) {
    return failed(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.STALE,
      "The restored numbering definitions differ from the exact inverse.",
    );
  }
  if (
    (op.restore?.type === "omitted" || op.restore?.type === "undefined") &&
    (numbering.nums.length !== 0 ||
      numbering.abstractNums.length !== 0 ||
      prior.preserved !== undefined ||
      prior.preservedAttributes !== undefined)
  ) {
    return failed(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.STALE,
      "The numbering part gained data after the instance was created.",
    );
  }
  const next =
    op.restore === undefined ? intermediate : withNumberingState(intermediate, op.restore);
  return Result.ok(
    emptyEdit(next, {
      type: DOCUMENT_OP_TYPES.CREATE_NUMBERING_INSTANCE,
      num: op.num,
      ...(op.abstractNum === undefined ? {} : { abstractNum: op.abstractNum }),
      expected: numberingStateOf(next),
      restore: priorState,
    }),
  );
};

const groupsFor = (blocks: readonly BlockContent[]): BlockContent[][] => {
  const groups: BlockContent[][] = [];
  let current: BlockContent[] = [];
  for (const block of blocks) {
    current.push(block);
    if (block.type === "paragraph" && block.sectionProperties !== undefined) {
      groups.push(current);
      current = [];
    }
  }
  if (current.length > 0 || groups.length === 0) groups.push(current);
  return groups;
};

const sectionViewOf = (sections: readonly Section[] | undefined): SectionViewEntry[] | undefined =>
  sections?.map(({ properties, headers, footers }) => ({
    properties,
    ...(headers === undefined ? {} : { headers: [...headers] }),
    ...(footers === undefined ? {} : { footers: [...footers] }),
  }));

const sectionFromView = (
  { properties, headers, footers }: SectionViewEntry,
  content: BlockContent[],
): Section => ({
  properties,
  content,
  ...(headers === undefined ? {} : { headers: new Map(headers) }),
  ...(footers === undefined ? {} : { footers: new Map(footers) }),
});

const nextSectionView = (
  document: Document,
  content: BlockContent[],
  finalProperties: SectionProperties | undefined,
): Section[] | undefined => {
  const previous = document.package.document.sections;
  if (previous === undefined) return undefined;
  const groups = groupsFor(content);
  const oldGroups = groupsFor(document.package.document.content);
  return groups.map((group, index) => {
    const end = group.at(-1);
    const previousIndex = oldGroups.findIndex(
      (oldGroup) => end !== undefined && oldGroup.includes(end),
    );
    const prior =
      previous[previousIndex < 0 ? Math.min(index, previous.length - 1) : previousIndex];
    const properties =
      end?.type === "paragraph" && end.sectionProperties !== undefined
        ? end.sectionProperties
        : (finalProperties ?? prior?.properties ?? {});
    return Object.assign({}, prior, { properties, content: group });
  });
};

const applySectionEndpoint = (document: Document, op: SetSectionEndpointOp) => {
  const body = document.package.document;
  const { endpoint } = op;
  const oldView = sectionViewOf(body.sections);
  if (
    op.expectedSectionMetadata !== undefined &&
    !structurallyEqual(oldView, op.expectedSectionMetadata ?? undefined)
  ) {
    return failed(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.STALE,
      "The section view differs from the expected value.",
    );
  }
  const target =
    endpoint.type === "paragraph"
      ? body.content.find(
          (block) => block.type === "paragraph" && block.paraId === endpoint.blockId,
        )
      : undefined;
  let previous = body.finalSectionProperties;
  if (endpoint.type === "paragraph")
    previous = target?.type === "paragraph" ? target.sectionProperties : undefined;
  const found = endpoint.type === "final" || target !== undefined;
  if (!found)
    return failed(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.BLOCK_NOT_FOUND,
      "The section endpoint paragraph does not exist.",
    );
  if (
    op.expected !== undefined &&
    !structurallyEqual(previous, op.expected.type === "absent" ? undefined : op.expected.value)
  )
    return failed(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.STALE,
      "The section endpoint differs from the expected value.",
    );
  if (structurallyEqual(previous, op.properties)) return unchanged(document);
  const content =
    endpoint.type === "final"
      ? body.content
      : body.content.map((block) => {
          if (block.type !== "paragraph" || block.paraId !== endpoint.blockId) return block;
          const paragraph = Object.assign({}, block);
          if (op.properties === undefined) delete paragraph.sectionProperties;
          else paragraph.sectionProperties = op.properties;
          return paragraph;
        });
  const finalProperties = endpoint.type === "final" ? op.properties : body.finalSectionProperties;
  const sections =
    op.sectionMetadata === undefined
      ? nextSectionView(document, content, finalProperties)
      : groupsFor(content).map((group, index) =>
          sectionFromView(op.sectionMetadata?.[index] ?? { properties: {} }, group),
        );
  const bodyWithoutFinalProperties = { ...body };
  delete bodyWithoutFinalProperties.finalSectionProperties;
  const nextBody = {
    ...(endpoint.type === "final" ? bodyWithoutFinalProperties : body),
    content,
    ...(endpoint.type === "final" && op.properties !== undefined
      ? { finalSectionProperties: op.properties }
      : {}),
    ...(sections === undefined ? {} : { sections }),
  };
  const next = { ...document, package: { ...document.package, document: nextBody } };
  const inverse: SetSectionEndpointOp = {
    type: DOCUMENT_OP_TYPES.SET_SECTION_ENDPOINT,
    endpoint,
    expected:
      op.properties === undefined ? { type: "absent" } : { type: "present", value: op.properties },
    ...(previous === undefined ? {} : { properties: previous }),
    expectedSectionMetadata: sectionViewOf(sections) ?? null,
    ...(oldView === undefined ? {} : { sectionMetadata: oldView }),
  };
  return Result.ok({
    document: next,
    inverse: [inverse],
    touched: {
      modified: endpoint.type === "paragraph" ? [endpoint.blockId] : [],
      inserted: [],
      removed: [],
    },
  });
};

export const applyNumberingSectionOp = (
  document: Document,
  op: CreateNumberingInstanceOp | DeleteNumberingInstanceOp | SetSectionEndpointOp,
) => {
  switch (op.type) {
    case DOCUMENT_OP_TYPES.CREATE_NUMBERING_INSTANCE:
      return applyCreateNumberingInstance(document, op);
    case DOCUMENT_OP_TYPES.DELETE_NUMBERING_INSTANCE:
      return applyDeleteNumberingInstance(document, op);
    case DOCUMENT_OP_TYPES.SET_SECTION_ENDPOINT:
      return applySectionEndpoint(document, op);
    default: {
      const unreachable: never = op;
      return unreachable;
    }
  }
};
