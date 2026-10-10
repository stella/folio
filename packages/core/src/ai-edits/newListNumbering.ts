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

  const resolveOperation = (operation: FolioAIEditOperation): FolioAIResolvedEditOperation => {
    const instances = new Map<ListKind, number>();
    const resolve = (
      value: FolioAIListNumbering | undefined,
    ): FolioContentStatedNumbering | undefined => {
      if (value === undefined || !isFolioAINewListReference(value)) {
        return value;
      }
      let numId = instances.get(value.format);
      if (numId === undefined) {
        const instance = mintListInstance(definitions, { kind: value.format });
        definitions = instance.definitions;
        numId = instance.numId;
        instances.set(value.format, numId);
        minted = true;
      }
      return paragraphNumberingReference({ numId, ilvl: value.level ?? 0 });
    };
    const resolveProperties = (
      properties: FolioAIBlockParagraphProperties,
    ): FolioAIBlockParagraphProperties<FolioContentStatedNumbering> => {
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
      case "setBlockParagraphProperties": {
        return { ...operation, properties: resolveProperties(operation.properties) };
      }
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

  const resolved = operations.map(resolveOperation);
  return {
    operations: resolved,
    numbering: minted && definitions ? createNumberingMap(definitions) : numbering,
  };
};
