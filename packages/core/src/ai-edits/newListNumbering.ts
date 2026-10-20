import type { FolioContentStatedNumbering } from "../compare/content-types";
/**
 * Operations that start a new list (`numbering: { kind: "newList", format }`)
 * resolved to the numbering instance each one defines.
 *
 * The instance is minted against the document's numbering (the package's
 * definitions plus every instance its paragraphs already define), so its id
 * collides with nothing, and against the instances earlier operations in the
 * batch minted, so two operations start two lists. The paragraphs the
 * operation numbers then carry the instance's rendering, which is what
 * defines it in the saved package (see `docx/listNumberingInstances.ts`).
 */

import { mintListInstance, type ListKind } from "../docx/listNumberingInstances";
import { paragraphNumberingReference } from "../docx/numberingReference";
import { createNumberingMap, type NumberingMap } from "../docx/numberingParser";
import type { NumberingDefinitions } from "../types/document";
import { panic } from "better-result";
import type {
  FolioAIBlockParagraphProperties,
  FolioAIEditOperation,
  FolioAIResolvedEditOperation,
  FolioAIListNumbering,
  FolioAINewListReference,
} from "./types";

/** Whether `numbering` asks for a new list rather than naming an instance. */
export const isFolioAINewListReference = (
  numbering: FolioAIListNumbering,
): numbering is FolioAINewListReference => numbering.kind === "newList";

type ResolvedNewLists = {
  operations: FolioAIResolvedEditOperation[];
  numbering: NumberingMap | null;
};

type DeferredNewListReference = {
  kind: "deferredNewList";
  allocation: { format: ListKind };
  level: number;
};

export type FolioAIStagedEditOperation = FolioAIEditOperation<
  FolioContentStatedNumbering | DeferredNewListReference
>;

const operationNumbering = (
  operation: FolioAIStagedEditOperation,
): readonly (FolioContentStatedNumbering | DeferredNewListReference | undefined)[] => {
  switch (operation.type) {
    case "insertAfterBlock":
    case "insertBeforeBlock":
      return [operation.numbering];
    case "setBlockParagraphProperties":
      return [operation.properties.numbering];
    case "splitBlock":
      return [
        operation.firstParagraphProperties?.numbering,
        operation.secondParagraphProperties?.numbering,
      ];
    case "mergeBlockWithNext":
      return [operation.mergedParagraphProperties?.numbering];
    default:
      return [];
  }
};

export function assertNewListsAllocated(
  operation: FolioAIStagedEditOperation,
): asserts operation is FolioAIResolvedEditOperation {
  if (operationNumbering(operation).some((numbering) => numbering?.kind === "deferredNewList")) {
    panic("A deferred new-list request reached execution", { operationId: operation.id });
  }
}

const deferNewListRequests = (
  operations: readonly FolioAIEditOperation[],
): FolioAIStagedEditOperation[] => {
  const deferOperation = (operation: FolioAIEditOperation): FolioAIStagedEditOperation => {
    const allocations = new Map<ListKind, DeferredNewListReference["allocation"]>();
    const resolve = (
      value: FolioAIListNumbering | undefined,
    ): FolioContentStatedNumbering | DeferredNewListReference | undefined => {
      if (value === undefined || !isFolioAINewListReference(value)) return value;
      let allocation = allocations.get(value.format);
      if (allocation === undefined) {
        allocation = { format: value.format };
        allocations.set(value.format, allocation);
      }
      return { kind: "deferredNewList", allocation, level: value.level ?? 0 };
    };
    const resolveProperties = (
      properties: FolioAIBlockParagraphProperties,
    ): FolioAIBlockParagraphProperties<FolioContentStatedNumbering | DeferredNewListReference> => {
      const { numbering: request, ...rest } = properties;
      const resolved = resolve(request);
      return { ...rest, ...(resolved !== undefined && { numbering: resolved }) };
    };

    switch (operation.type) {
      case "insertAfterBlock":
      case "insertBeforeBlock": {
        const { numbering: request, ...rest } = operation;
        const resolved = resolve(request);
        return { ...rest, ...(resolved !== undefined && { numbering: resolved }) };
      }
      case "setBlockParagraphProperties":
        return { ...operation, properties: resolveProperties(operation.properties) };
      case "splitBlock": {
        const {
          firstParagraphProperties: firstRequest,
          secondParagraphProperties: secondRequest,
          ...rest
        } = operation;
        return {
          ...rest,
          ...(firstRequest !== undefined && {
            firstParagraphProperties: resolveProperties(firstRequest),
          }),
          ...(secondRequest !== undefined && {
            secondParagraphProperties: resolveProperties(secondRequest),
          }),
        };
      }
      case "mergeBlockWithNext": {
        const { mergedParagraphProperties: request, ...rest } = operation;
        return {
          ...rest,
          ...(request !== undefined && { mergedParagraphProperties: resolveProperties(request) }),
        };
      }
      default:
        return operation;
    }
  };
  return operations.map(deferOperation);
};

/** Make requests structurally plannable without allocating document resources. */
export const stageNewListOperations = (
  operations: readonly FolioAIEditOperation[],
): FolioAIStagedEditOperation[] => deferNewListRequests(operations);

/**
 * Replace every new-list request with a reference to an instance minted for
 * its operation: one instance per operation and kind. `numbering` is the map
 * the batch applies against; the result's map adds the minted instances.
 */
export const resolveNewListOperations = (
  operations: readonly FolioAIStagedEditOperation[],
  numbering: NumberingMap | null,
): ResolvedNewLists => {
  let definitions: NumberingDefinitions | undefined = numbering?.definitions;
  let minted = false;
  const resolved = operations.map((operation): FolioAIResolvedEditOperation => {
    const instances = new Map<DeferredNewListReference["allocation"], number>();
    const resolve = (
      value: FolioContentStatedNumbering | DeferredNewListReference | undefined,
    ): FolioContentStatedNumbering | undefined => {
      if (value === undefined) return value;
      if (value.kind !== "deferredNewList") return value;
      let numId = instances.get(value.allocation);
      if (numId === undefined) {
        const instance = mintListInstance(definitions, { kind: value.allocation.format });
        definitions = instance.definitions;
        minted = true;
        numId = instance.numId;
        instances.set(value.allocation, numId);
      }
      return paragraphNumberingReference({ numId, ilvl: value.level });
    };
    const resolveProperties = (
      properties: FolioAIBlockParagraphProperties<
        FolioContentStatedNumbering | DeferredNewListReference
      >,
    ): FolioAIBlockParagraphProperties<FolioContentStatedNumbering> => {
      const { numbering: request, ...rest } = properties;
      const resolvedNumbering = resolve(request);
      return { ...rest, ...(resolvedNumbering !== undefined && { numbering: resolvedNumbering }) };
    };
    switch (operation.type) {
      case "insertAfterBlock":
      case "insertBeforeBlock": {
        const { numbering: request, ...rest } = operation;
        const resolvedNumbering = resolve(request);
        return {
          ...rest,
          ...(resolvedNumbering !== undefined && { numbering: resolvedNumbering }),
        };
      }
      case "setBlockParagraphProperties":
        return { ...operation, properties: resolveProperties(operation.properties) };
      case "splitBlock": {
        const { firstParagraphProperties, secondParagraphProperties, ...rest } = operation;
        return {
          ...rest,
          ...(firstParagraphProperties !== undefined && {
            firstParagraphProperties: resolveProperties(firstParagraphProperties),
          }),
          ...(secondParagraphProperties !== undefined && {
            secondParagraphProperties: resolveProperties(secondParagraphProperties),
          }),
        };
      }
      case "mergeBlockWithNext": {
        const { mergedParagraphProperties, ...rest } = operation;
        return {
          ...rest,
          ...(mergedParagraphProperties !== undefined && {
            mergedParagraphProperties: resolveProperties(mergedParagraphProperties),
          }),
        };
      }
      default:
        return operation;
    }
  });
  return {
    operations: resolved,
    numbering: minted && definitions ? createNumberingMap(definitions) : numbering,
  };
};
