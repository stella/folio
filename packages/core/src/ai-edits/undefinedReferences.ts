/**
 * The one check both editing engines run before an explicit editing request
 * applies: every paragraph style and numbering instance an operation names
 * must exist in the document. A request naming one that does not is refused
 * with `missingStyle` or `missingNumbering`, whichever engine holds the
 * document; source-preserving imports (`undefinedReferences: "keep"`) keep
 * their references verbatim.
 */

import { panic } from "better-result";

import { isNumberingReference } from "../docx/numberingReference";
import type {
  ReadonlyStyle,
  StyleResolver,
} from "../prosemirror/styles/styleResolver";
import type { FolioAIStagedEditOperation } from "./newListNumbering";
import { classifyResourceReference } from "./referenceClassification";
import type { FolioAIEditOperation } from "./types";

/**
 * What an editing request does with a style or numbering reference the
 * document does not define: `refuse` the operation, or `keep` the reference
 * verbatim (source-preserving imports).
 */
export type FolioUndefinedReferencePolicy = "refuse" | "keep";

/** An operation as either engine holds it: staged (new lists deferred) or as parsed. */
type StagedOperation = FolioAIStagedEditOperation | FolioAIEditOperation;
type StagedOperationNumbering = NonNullable<
  Extract<
    StagedOperation,
    { type: "insertAfterBlock" | "insertBeforeBlock" }
  >["numbering"]
>;

type UndefinedParagraphStyleReference = {
  path: string;
  styleId: string;
  /** The type of the definition the id names, when it names one of another type. */
  definedAs: ReadonlyStyle["type"] | undefined;
};

/**
 * The styles a document defines, as the reference check reads them; `null`
 * when the caller cannot say (a state without the styles plugin), which checks
 * nothing.
 */
export type StyleDefinitionLookup = {
  get: (styleId: string) => ReadonlyStyle | undefined;
  paragraphStyles: () => ReadonlyStyle[];
} | null;

/** The lookup over a style resolver, for callers that hold a document model. */
export const styleDefinitionLookup = (
  resolver: StyleResolver,
): NonNullable<StyleDefinitionLookup> => ({
  get: (styleId) => resolver.getStyle(styleId),
  paragraphStyles: () => resolver.getParagraphStyles(),
});

/**
 * Every paragraph style an operation names, with the field that names it.
 * `null` clears the style and names none; an omitted one keeps the style the
 * paragraph or its anchor already carries, which the document holds by
 * construction.
 */
const operationParagraphStyleReferences = (
  operation: StagedOperation,
): { path: string; styleId: string | null | undefined }[] => {
  switch (operation.type) {
    case "insertAfterBlock":
    case "insertBeforeBlock":
    case "replaceBlock":
      return [{ path: "styleId", styleId: operation.styleId }];
    case "setBlockParagraphProperties":
      return [
        { path: "properties.styleId", styleId: operation.properties.styleId },
      ];
    case "splitBlock":
      return [
        {
          path: "firstParagraphProperties.styleId",
          styleId: operation.firstParagraphProperties?.styleId,
        },
        {
          path: "secondParagraphProperties.styleId",
          styleId: operation.secondParagraphProperties?.styleId,
        },
      ];
    case "mergeBlockWithNext":
      return [
        {
          path: "mergedParagraphProperties.styleId",
          styleId: operation.mergedParagraphProperties?.styleId,
        },
      ];
    default:
      return [];
  }
};

type FindUndefinedParagraphStyleReferenceOptions = {
  operation: StagedOperation;
  definitions: StyleDefinitionLookup;
  undefinedReferences: FolioUndefinedReferencePolicy;
};

/**
 * Explicit editing requests refuse unavailable paragraph styles before apply.
 * Source-preserving imports keep their references verbatim; closure import
 * validates every defined resource separately. A state without the styles
 * plugin cannot classify a reference and checks nothing.
 */
const findUndefinedParagraphStyleReference = ({
  operation,
  definitions,
  undefinedReferences,
}: FindUndefinedParagraphStyleReferenceOptions):
  UndefinedParagraphStyleReference | undefined => {
  if (definitions === null) {
    return undefined;
  }
  for (const { path, styleId } of operationParagraphStyleReferences(
    operation,
  )) {
    if (styleId === null || styleId === undefined) {
      continue;
    }
    const reference = classifyResourceReference({
      kind: "style",
      id: styleId,
      sourceDefines: (id) => definitions.get(id) !== undefined,
      destinationDefines: (id) => definitions.get(id) !== undefined,
    });
    if (undefinedReferences === "keep") continue;
    if (reference !== "sourceDefined") {
      return { path, styleId, definedAs: undefined };
    }
    const definition =
      definitions.get(styleId) ??
      panic("A classified style lost its definition");
    if (definition.type !== "paragraph") {
      return { path, styleId, definedAs: definition.type };
    }
  }
  return undefined;
};

const MAX_LISTED_PARAGRAPH_STYLES = 12;

const describeUndefinedParagraphStyle = (
  { path, styleId, definedAs }: UndefinedParagraphStyleReference,
  definitions: StyleDefinitionLookup,
): string => {
  const defined = (definitions?.paragraphStyles() ?? []).map(
    (style) => style.styleId,
  );
  let available = "it defines no visible paragraph styles";
  if (defined.length > 0) {
    const listed = defined.slice(0, MAX_LISTED_PARAGRAPH_STYLES).join(", ");
    available = `its paragraph styles include ${listed}${defined.length > MAX_LISTED_PARAGRAPH_STYLES ? ", …" : ""}`;
  }
  const names =
    definedAs === undefined
      ? "names no style in this document"
      : `names a ${definedAs} style, not a paragraph style`;
  return `${path} "${styleId}" ${names} (${available}).`;
};

type UndefinedNumberingReference = { path: string; numId: number };

/** The `w:num` id a stated numbering names, when it names one. */
const statedNumberingInstance = (
  numbering: StagedOperationNumbering | undefined,
): number | undefined =>
  numbering?.kind === "reference" && isNumberingReference(numbering.numId)
    ? numbering.numId
    : undefined;

/**
 * Every numbering instance an operation names, with the field that names it.
 * Only reference variants name an instance; inheritance, cancellation and level-only do not.
 */
const operationNumberingReferences = (
  operation: StagedOperation,
): { path: string; numbering: StagedOperationNumbering | undefined }[] => {
  switch (operation.type) {
    case "insertAfterBlock":
    case "insertBeforeBlock":
      return [{ path: "numbering", numbering: operation.numbering }];
    case "setBlockParagraphProperties":
      return [
        {
          path: "properties.numbering",
          numbering: operation.properties.numbering,
        },
      ];
    case "splitBlock":
      return [
        {
          path: "firstParagraphProperties.numbering",
          numbering: operation.firstParagraphProperties?.numbering,
        },
        {
          path: "secondParagraphProperties.numbering",
          numbering: operation.secondParagraphProperties?.numbering,
        },
      ];
    case "mergeBlockWithNext":
      return [
        {
          path: "mergedParagraphProperties.numbering",
          numbering: operation.mergedParagraphProperties?.numbering,
        },
      ];
    default:
      return [];
  }
};

/**
 * The first numbering instance the operation names that the document does not
 * define, subject to the operation's undefined-reference policy. `null` instance
 * ids (a state without the numbering plugin) cannot say what is defined, and
 * check nothing.
 */
type FindUndefinedNumberingReferenceOptions = {
  operation: StagedOperation;
  instanceIds: ReadonlySet<number> | null;
  undefinedReferences: FolioUndefinedReferencePolicy;
};

const findUndefinedNumberingReference = ({
  operation,
  instanceIds,
  undefinedReferences,
}: FindUndefinedNumberingReferenceOptions):
  UndefinedNumberingReference | undefined => {
  if (instanceIds === null) {
    return undefined;
  }
  for (const { path, numbering } of operationNumberingReferences(operation)) {
    const numId = statedNumberingInstance(numbering);
    if (numId === undefined) continue;
    const reference = classifyResourceReference({
      kind: "numbering",
      id: numId,
      sourceDefines: (id) => instanceIds.has(id),
      destinationDefines: (id) => instanceIds.has(id),
    });
    if (undefinedReferences !== "keep" && reference !== "sourceDefined") {
      return { path, numId };
    }
  }
  return undefined;
};

const MAX_LISTED_NUMBERING_INSTANCES = 12;

const describeUndefinedNumbering = (
  { path, numId }: UndefinedNumberingReference,
  instanceIds: ReadonlySet<number> | null,
): string => {
  const defined = [...(instanceIds ?? [])].sort((left, right) => left - right);
  let available = "it defines none";
  if (defined.length > 0) {
    const listed = defined.slice(0, MAX_LISTED_NUMBERING_INSTANCES).join(", ");
    available = `it defines ${listed}${defined.length > MAX_LISTED_NUMBERING_INSTANCES ? ", …" : ""}`;
  }
  return `${path}.numId ${String(numId)} names no numbering instance in this document (${available}).`;
};

/** Why an operation names a reference the document does not define. */
export type UndefinedReferenceRefusal = {
  reason: "missingStyle" | "missingNumbering";
  message: string;
};

type FindUndefinedOperationReferenceOptions = {
  operation: StagedOperation;
  styles: StyleDefinitionLookup;
  numberingInstanceIds: ReadonlySet<number> | null;
  undefinedReferences: FolioUndefinedReferencePolicy;
};

/** The first style or numbering reference the operation names that the document does not define. */
export const findUndefinedOperationReference = ({
  operation,
  styles,
  numberingInstanceIds,
  undefinedReferences,
}: FindUndefinedOperationReferenceOptions):
  UndefinedReferenceRefusal | undefined => {
  const style = findUndefinedParagraphStyleReference({
    operation,
    definitions: styles,
    undefinedReferences,
  });
  if (style !== undefined)
    return {
      reason: "missingStyle",
      message: describeUndefinedParagraphStyle(style, styles),
    };
  const numbering = findUndefinedNumberingReference({
    operation,
    instanceIds: numberingInstanceIds,
    undefinedReferences,
  });
  if (numbering !== undefined)
    return {
      reason: "missingNumbering",
      message: describeUndefinedNumbering(numbering, numberingInstanceIds),
    };
  return undefined;
};
