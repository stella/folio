/**
 * Operations that start a new list (`numbering: { start: "new", kind }`)
 * resolved to the numbering instance each one defines.
 *
 * The instance is minted against the document's numbering (the package's
 * definitions plus every instance its paragraphs already define), so its id
 * collides with nothing, and against the instances earlier operations in the
 * batch minted, so two operations start two lists. The paragraphs the
 * operation numbers then carry the instance's rendering, which is what
 * defines it in the saved package (see `docx/listNumberingInstances.ts`).
 */

import { panic } from "better-result";

import { mintListInstance, type ListKind } from "../docx/listNumberingInstances";
import { createNumberingMap, type NumberingMap } from "../docx/numberingParser";
import type { NumberingDefinitions } from "../types/document";
import type {
  FolioAIBlockParagraphProperties,
  FolioAIEditOperation,
  FolioAIListNumbering,
  FolioAIListReference,
  FolioAINewListReference,
} from "./types";

/** Whether `numbering` asks for a new list rather than naming an instance. */
export const isFolioAINewListReference = (
  numbering: FolioAIListNumbering,
): numbering is FolioAINewListReference => "start" in numbering;

/**
 * The concrete instance an operation's numbering names once new lists are
 * resolved. A new-list request reaching this point skipped resolution, which
 * is a bug in the apply path rather than a malformed operation.
 */
export const concreteListReference = (numbering: FolioAIListNumbering): FolioAIListReference => {
  if (isFolioAINewListReference(numbering)) {
    return panic("A new-list numbering request reached apply unresolved");
  }
  return numbering;
};

type ResolvedNewLists = {
  operations: FolioAIEditOperation[];
  numbering: NumberingMap | null;
};

/**
 * Replace every new-list request with a reference to an instance minted for
 * its operation: one instance per operation and kind. `numbering` is the map
 * the batch applies against; the result's map adds the minted instances.
 */
export const resolveNewListOperations = (
  operations: readonly FolioAIEditOperation[],
  numbering: NumberingMap | null,
): ResolvedNewLists => {
  let definitions: NumberingDefinitions | undefined = numbering?.definitions;
  let minted = false;

  const resolveOperation = (operation: FolioAIEditOperation): FolioAIEditOperation => {
    const instances = new Map<ListKind, number>();
    const resolve = <T extends FolioAIListNumbering | null | undefined>(
      value: T,
    ): T | FolioAIListReference => {
      if (value === null || value === undefined || !isFolioAINewListReference(value)) {
        return value;
      }
      let numId = instances.get(value.kind);
      if (numId === undefined) {
        const instance = mintListInstance(definitions, { kind: value.kind });
        definitions = instance.definitions;
        numId = instance.numId;
        instances.set(value.kind, numId);
        minted = true;
      }
      return { numId, level: value.level ?? 0 };
    };
    const resolveProperties = (
      properties: FolioAIBlockParagraphProperties | undefined,
    ): FolioAIBlockParagraphProperties | undefined =>
      properties?.numbering
        ? { ...properties, numbering: resolve(properties.numbering) }
        : properties;

    switch (operation.type) {
      case "insertAfterBlock":
      case "insertBeforeBlock":
        return operation.numbering
          ? { ...operation, numbering: resolve(operation.numbering) }
          : operation;
      case "setBlockParagraphProperties":
        return operation.properties.numbering
          ? {
              ...operation,
              properties: {
                ...operation.properties,
                numbering: resolve(operation.properties.numbering),
              },
            }
          : operation;
      case "splitBlock": {
        const firstParagraphProperties = resolveProperties(operation.firstParagraphProperties);
        const secondParagraphProperties = resolveProperties(operation.secondParagraphProperties);
        return {
          ...operation,
          ...(firstParagraphProperties && { firstParagraphProperties }),
          ...(secondParagraphProperties && { secondParagraphProperties }),
        };
      }
      case "mergeBlockWithNext": {
        const mergedParagraphProperties = resolveProperties(operation.mergedParagraphProperties);
        return { ...operation, ...(mergedParagraphProperties && { mergedParagraphProperties }) };
      }
      default:
        return operation;
    }
  };

  const resolved = operations.map(resolveOperation);
  return {
    operations: resolved,
    numbering: minted && definitions ? createNumberingMap(definitions) : numbering,
  };
};
