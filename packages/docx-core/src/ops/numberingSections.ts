import { Result } from "better-result";
import { isNumberingReference, NO_NUMBERING_NUM_ID } from "../model/paragraphNumbering";

import {
  MAX_REVISION_ID,
  type Document,
  type DocumentBody,
  type Paragraph,
} from "../model/document";
import {
  captureSectionViewState,
  rebuildSections,
  restoreSectionViewState,
  sectionsInStep,
} from "./blocks";
import { structurallyEqual } from "./equality";
import { idKey } from "./ids";
import { DOCUMENT_OP_REFUSAL_REASONS, DocumentOpRefusal } from "./refusal";
import {
  DOCUMENT_OP_TYPES,
  type CreateNumberingInstanceOp,
  type DeleteNumberingInstanceOp,
  type DocumentOp,
  type NumberingPartState,
  type SetSectionEndpointOp,
  type SectionPropertiesState,
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
  if (
    !Number.isSafeInteger(op.num.numId) ||
    !isNumberingReference(op.num.numId) ||
    op.num.numId < NO_NUMBERING_NUM_ID ||
    op.num.numId > MAX_REVISION_ID
  ) {
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

type SectionPropertiesRecord = Pick<Paragraph, "sectionProperties"> &
  Pick<DocumentBody, "finalSectionProperties">;

type CaptureSectionPropertiesOptions = {
  record: SectionPropertiesRecord;
  key: keyof SectionPropertiesRecord;
};

const captureSectionProperties = ({
  record,
  key,
}: CaptureSectionPropertiesOptions): SectionPropertiesState => {
  const value = record[key];
  if (value !== undefined) return { type: "present", value };
  return Object.hasOwn(record, key) ? { type: "undefined" } : { type: "omitted" };
};

type RestoreSectionPropertiesOptions = CaptureSectionPropertiesOptions & {
  state: SectionPropertiesState;
};

const restoreSectionProperties = ({
  record,
  key,
  state,
}: RestoreSectionPropertiesOptions): void => {
  switch (state.type) {
    case "omitted":
      switch (key) {
        case "sectionProperties":
          delete record.sectionProperties;
          return;
        case "finalSectionProperties":
          delete record.finalSectionProperties;
          return;
        default: {
          const unreachable: never = key;
          return unreachable;
        }
      }
    case "undefined":
      record[key] = undefined;
      return;
    case "present":
      record[key] = state.value;
      return;
    default: {
      const unreachable: never = state;
      return unreachable;
    }
  }
};

const applySectionEndpoint = (document: Document, op: SetSectionEndpointOp) => {
  const body = document.package.document;
  const { endpoint } = op;
  const oldView = captureSectionViewState(body);
  if (
    op.expectedSectionMetadata !== undefined &&
    !structurallyEqual(oldView, op.expectedSectionMetadata)
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
          (block) =>
            block.type === "paragraph" &&
            block.paraId !== undefined &&
            idKey(block.paraId) === idKey(endpoint.blockId),
        )
      : undefined;
  const endpointParagraph = target?.type === "paragraph" ? target : undefined;
  const record = endpoint.type === "final" ? body : endpointParagraph;
  if (record === undefined)
    return failed(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.BLOCK_NOT_FOUND,
      "The section endpoint paragraph does not exist.",
    );
  const key = endpoint.type === "final" ? "finalSectionProperties" : "sectionProperties";
  const previous = captureSectionProperties({ record, key });
  if (op.expected !== undefined && !structurallyEqual(previous, op.expected))
    return failed(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.STALE,
      "The section endpoint differs from the expected value.",
    );
  const content =
    endpoint.type === "final"
      ? body.content
      : body.content.map((block) => {
          if (
            block.type !== "paragraph" ||
            block.paraId === undefined ||
            idKey(block.paraId) !== idKey(endpoint.blockId)
          )
            return block;
          const paragraph = Object.assign({}, block);
          restoreSectionProperties({
            record: paragraph,
            key: "sectionProperties",
            state: op.properties,
          });
          return paragraph;
        });
  let nextBody = { ...body, content };
  if (endpoint.type === "final")
    restoreSectionProperties({
      record: nextBody,
      key: "finalSectionProperties",
      state: op.properties,
    });
  if (op.sectionMetadata !== undefined) {
    const restored = restoreSectionViewState(nextBody, op.sectionMetadata);
    if (restored.isErr()) return failed(op, restored.error.reason, restored.error.message);
    nextBody = restored.value;
  } else if (body.sections !== undefined) {
    const source = {
      ...document,
      package: {
        ...document.package,
        document: {
          ...body,
          finalSectionProperties: nextBody.finalSectionProperties,
        },
      },
    };
    const rebuilt = rebuildSections({ document: source, content, previous: body.sections });
    if (rebuilt.isErr()) return failed(op, rebuilt.error.reason, rebuilt.error.message);
    nextBody.sections = rebuilt.value;
  }
  if (!sectionsInStep(nextBody))
    return failed(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.SECTION_BOUNDARY,
      "The section metadata conflicts with its canonical boundaries.",
    );
  const nextView = captureSectionViewState(nextBody);
  if (structurallyEqual(previous, op.properties) && structurallyEqual(oldView, nextView))
    return unchanged(document);
  const next = { ...document, package: { ...document.package, document: nextBody } };
  const inverse: SetSectionEndpointOp = {
    type: DOCUMENT_OP_TYPES.SET_SECTION_ENDPOINT,
    endpoint,
    expected: op.properties,
    properties: previous,
    expectedSectionMetadata: nextView,
    sectionMetadata: oldView,
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
