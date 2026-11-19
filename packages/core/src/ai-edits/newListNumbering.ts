import { panic } from "better-result";
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

import type { NewListInstance } from "@stll/docx-core/ops";

import type { FolioContentStatedNumbering } from "../compare/content-types";
import {
  mintListInstance,
  type NewListDefinition,
} from "../docx/listNumberingInstances";
import {
  NO_NUMBERING_IDS_IN_USE,
  type NumberingIdsInUse,
} from "../docx/numberingIds";
import { createNumberingMap, type NumberingMap } from "../docx/numberingParser";
import { paragraphNumberingReference } from "../docx/numberingReference";
import type { Document, NumberingDefinitions } from "../types/document";
import { canonicalJson } from "../utils/canonicalJson";
import type {
  FolioAIBlockParagraphProperties,
  FolioAIEditOperation,
  FolioAIListNumbering,
  FolioAINewListReference,
  FolioAIResolvedEditOperation,
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
  allocation: NewListDefinition;
  level: number;
};

/** Start a list discovered while compiling one insertion. */
export const deferInsertedList = (
  kind: NewListDefinition["kind"],
): DeferredNewListReference => ({
  kind: "deferredNewList",
  allocation: { kind },
  level: 0,
});

export type FolioAIStagedEditOperation = FolioAIEditOperation<
  FolioContentStatedNumbering | DeferredNewListReference
>;

const operationNumbering = (
  operation: FolioAIStagedEditOperation,
): readonly (
  FolioContentStatedNumbering | DeferredNewListReference | undefined
)[] => {
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
  if (
    operationNumbering(operation).some(
      (numbering) => numbering?.kind === "deferredNewList",
    )
  ) {
    panic("A deferred new-list request reached execution", {
      operationId: operation.id,
    });
  }
}

const newListDefinition = ({
  format: kind,
  level,
  numFmt,
  lvlText,
  start,
}: FolioAINewListReference): NewListDefinition => ({
  kind,
  ...(numFmt === undefined && lvlText === undefined
    ? {}
    : { format: { numFmt, lvlText } }),
  ...(start === undefined ? {} : { start }),
  ...(numFmt === undefined && lvlText === undefined && start === undefined
    ? {}
    : { level: level ?? 0 }),
});

const deferNewListRequests = (
  operations: readonly FolioAIEditOperation[],
): FolioAIStagedEditOperation[] => {
  const deferOperation = (
    operation: FolioAIEditOperation,
  ): FolioAIStagedEditOperation => {
    const allocations = new Map<string, NewListDefinition>();
    const resolve = (
      value: FolioAIListNumbering | undefined,
    ): FolioContentStatedNumbering | DeferredNewListReference | undefined => {
      if (value === undefined || !isFolioAINewListReference(value))
        return value;
      const definition = newListDefinition(value);
      const key = canonicalJson(definition);
      const allocation = allocations.get(key) ?? definition;
      allocations.set(key, allocation);
      return { kind: "deferredNewList", allocation, level: value.level ?? 0 };
    };
    const resolveProperties = (
      properties: FolioAIBlockParagraphProperties,
    ): FolioAIBlockParagraphProperties<
      FolioContentStatedNumbering | DeferredNewListReference
    > => {
      const { numbering: request, ...rest } = properties;
      const resolved = resolve(request);
      return {
        ...rest,
        ...(resolved !== undefined && { numbering: resolved }),
      };
    };

    switch (operation.type) {
      case "insertAfterBlock":
      case "insertBeforeBlock": {
        const { numbering: request, ...rest } = operation;
        const resolved = resolve(request);
        return {
          ...rest,
          ...(resolved !== undefined && { numbering: resolved }),
        };
      }
      case "setBlockParagraphProperties":
        return {
          ...operation,
          properties: resolveProperties(operation.properties),
        };
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
          ...(request !== undefined && {
            mergedParagraphProperties: resolveProperties(request),
          }),
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
 * its operation, shared by paragraphs requesting the same definition. `numbering` is the map
 * the batch applies against; the result's map adds the minted instances.
 * `elsewhere` holds the ids the package's other stories define, which no
 * minted instance takes.
 */
export const resolveNewListOperations = (
  operations: readonly FolioAIStagedEditOperation[],
  numbering: NumberingMap | null,
  elsewhere: NumberingIdsInUse,
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
        const instance = mintListInstance(
          definitions,
          value.allocation,
          elsewhere,
        );
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
      return {
        ...rest,
        ...(resolvedNumbering !== undefined && {
          numbering: resolvedNumbering,
        }),
      };
    };
    switch (operation.type) {
      case "insertAfterBlock":
      case "insertBeforeBlock": {
        const { numbering: request, ...rest } = operation;
        const resolvedNumbering = resolve(request);
        return {
          ...rest,
          ...(resolvedNumbering !== undefined && {
            numbering: resolvedNumbering,
          }),
        };
      }
      case "setBlockParagraphProperties":
        return {
          ...operation,
          properties: resolveProperties(operation.properties),
        };
      case "splitBlock": {
        const { firstParagraphProperties, secondParagraphProperties, ...rest } =
          operation;
        return {
          ...rest,
          ...(firstParagraphProperties !== undefined && {
            firstParagraphProperties: resolveProperties(
              firstParagraphProperties,
            ),
          }),
          ...(secondParagraphProperties !== undefined && {
            secondParagraphProperties: resolveProperties(
              secondParagraphProperties,
            ),
          }),
        };
      }
      case "mergeBlockWithNext": {
        const { mergedParagraphProperties, ...rest } = operation;
        return {
          ...rest,
          ...(mergedParagraphProperties !== undefined && {
            mergedParagraphProperties: resolveProperties(
              mergedParagraphProperties,
            ),
          }),
        };
      }
      default:
        return operation;
    }
  });
  return {
    operations: resolved,
    numbering:
      minted && definitions ? createNumberingMap(definitions) : numbering,
  };
};

/** A list one operation starts: its instance id and the definitions it adds. */
export type MintedList = { numId: number; newList: NewListInstance };

/**
 * Mints new lists for one batch of a canonical session, each after the ones
 * before it. The document is read at the first mint, so a batch that starts
 * no list never reads it.
 */
export const createBatchListMinter = (document: () => Document) => {
  let definitions: Document["package"]["numbering"] | null = null;
  return (request: FolioAINewListReference): MintedList => {
    definitions ??= document().package.numbering;
    const minted = mintListInstance(
      definitions,
      newListDefinition(request),
      NO_NUMBERING_IDS_IN_USE,
    );
    definitions = minted.definitions;
    const num =
      minted.definitions.nums.find(({ numId }) => numId === minted.numId) ??
      panic("A minted list is missing its instance.");
    const abstractNum = minted.definitions.abstractNums.find(
      ({ abstractNumId }) => abstractNumId === num.abstractNumId,
    );
    return {
      numId: minted.numId,
      newList: { num, ...(abstractNum === undefined ? {} : { abstractNum }) },
    };
  };
};
