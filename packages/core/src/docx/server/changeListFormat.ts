import { Result } from "better-result";

import {
  isNumberingLevel,
  MAX_REVISION_ID,
  NUMBER_FORMATS,
  type NumberFormat,
} from "@stll/docx-core/model";
import {
  type AppliedDocumentOp,
  applyDocumentOps,
  compileEditorIntent,
  DOCUMENT_OP_REFUSAL_REASONS,
  DOCUMENT_OP_TYPES,
  DocumentOpRefusal,
  type DocumentOpRefusalReason,
  documentStories,
  type EditorIntentMode,
  findStoryBody,
  idKey,
  OP_STORIES,
  type OpStory,
  storyParagraphs,
} from "@stll/docx-core/ops";

import { effectiveParagraphNumberingReference } from "../../prosemirror/numberingAttr";
import { createStyleResolver } from "../../prosemirror/styles/styleResolver";
import type { Document, ParagraphFormatting } from "../../types/document";
import { cloneListInstanceWithFormat } from "../listNumberingInstances";
import { NO_NUMBERING_IDS_IN_USE } from "../numberingIds";

export type ChangeListFormatOptions = {
  sourceNumId: number;
  /** Level whose format changes. Each paragraph retains its own level. */
  level: number;
  numFmt: Exclude<NumberFormat, "custom">;
  lvlText?: string;
  /** Omit to retain the source definition's start and start override. */
  start?: number;
  /** Defaults to the main story. */
  story?: OpStory;
  /** Omit to change every member of the instance in the story. */
  paragraphIds?: readonly string[];
  mode: EditorIntentMode;
};

const numberFormats: ReadonlySet<string> = new Set(
  NUMBER_FORMATS.filter((format) => format !== "custom"),
);

const refuse = (reason: DocumentOpRefusalReason, message: string) =>
  Result.err(
    new DocumentOpRefusal({
      reason,
      message,
      opType: DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS,
    }),
  );

// Imported definitions already record dangling references. Reserve references
// from caller-built models as well, so this edit cannot activate another list.
const referencedNumberingIds = (document: Document): ReadonlySet<number> => {
  const ids = new Set<number>();
  const add = (
    formatting:
      Pick<ParagraphFormatting, "numPr" | "numPrFromStyle"> | undefined,
  ) => {
    for (const reference of [formatting?.numPr, formatting?.numPrFromStyle]) {
      if (reference?.kind === "reference") ids.add(reference.numId);
    }
  };
  for (const story of documentStories(document)) {
    const body = findStoryBody(document, story);
    if (body === undefined) continue;
    for (const { paragraph } of storyParagraphs(body)) {
      add(paragraph.formatting);
      for (const change of paragraph.propertyChanges ?? [])
        add(change.previousFormatting);
    }
  }
  for (const comment of document.package.document.comments ?? [])
    for (const paragraph of comment.content) add(paragraph.formatting);
  const styles = document.package.styles;
  add(styles?.docDefaults?.pPr);
  for (const style of styles?.styles ?? []) {
    add(style.pPr);
    for (const part of style.tblStylePr ?? []) add(part.pPr);
  }
  return ids;
};

/** Clone one list instance and rebind the selected paragraphs as one edit. */
export const changeListFormat = (
  document: Document,
  {
    sourceNumId,
    level,
    numFmt,
    lvlText,
    start,
    story = OP_STORIES.MAIN,
    paragraphIds,
    mode,
  }: ChangeListFormatOptions,
): Result<AppliedDocumentOp, DocumentOpRefusal> => {
  if (
    !Number.isSafeInteger(sourceNumId) ||
    sourceNumId < 1 ||
    sourceNumId > MAX_REVISION_ID
  )
    return refuse(
      DOCUMENT_OP_REFUSAL_REASONS.INVALID_OPERATION,
      "The source list id is invalid.",
    );
  if (!isNumberingLevel(level) || !numberFormats.has(numFmt))
    return refuse(
      DOCUMENT_OP_REFUSAL_REASONS.INVALID_OPERATION,
      "The list format or level is invalid.",
    );
  if (start !== undefined && (!Number.isSafeInteger(start) || start < 0))
    return refuse(
      DOCUMENT_OP_REFUSAL_REASONS.INVALID_OPERATION,
      "The list start must be a non-negative safe integer.",
    );
  const body = findStoryBody(document, story);
  if (body === undefined)
    return refuse(
      DOCUMENT_OP_REFUSAL_REASONS.BLOCK_NOT_FOUND,
      "The list story does not exist.",
    );
  const requested =
    paragraphIds === undefined ? undefined : new Set(paragraphIds.map(idKey));
  const styles = createStyleResolver(document.package.styles);
  const levels = new Map<string, number>();
  for (const { paragraph } of storyParagraphs(body)) {
    const { paraId, formatting } = paragraph;
    if (
      requested !== undefined &&
      (paraId === undefined || !requested.has(idKey(paraId)))
    ) {
      continue;
    }
    const reference = effectiveParagraphNumberingReference({
      numPr: formatting?.numPr,
      numPrFromStyle:
        formatting?.styleId === undefined
          ? formatting?.numPrFromStyle
          : styles.resolveParagraphStyle(formatting.styleId).paragraphFormatting
              ?.numPr,
    });
    if (reference === undefined || reference.numId !== sourceNumId) continue;
    if (paraId === undefined)
      return refuse(
        DOCUMENT_OP_REFUSAL_REASONS.INVALID_BLOCK_ID,
        "A paragraph in the list has no paragraph id.",
      );
    levels.set(idKey(paraId), reference.ilvl ?? 0);
  }
  if (
    levels.size === 0 ||
    (requested !== undefined && requested.size !== levels.size)
  ) {
    return refuse(
      DOCUMENT_OP_REFUSAL_REASONS.BLOCK_NOT_FOUND,
      "Every selected paragraph must belong to the source list in the specified story.",
    );
  }
  const cloned = cloneListInstanceWithFormat(
    document.package.numbering,
    { sourceNumId, level, numFmt, lvlText, start },
    { ...NO_NUMBERING_IDS_IN_USE, numIds: referencedNumberingIds(document) },
  );
  const num =
    cloned === undefined
      ? undefined
      : cloned.definitions.nums.find(({ numId }) => numId === cloned.numId);
  if (num === undefined)
    return refuse(
      DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
      "The source list instance or its requested level is undefined.",
    );
  const compiled = compileEditorIntent(document, {
    intent: {
      type: "setList",
      items: [...levels].map(([blockId, ilvl]) => ({
        at: { story, blockId, offset: 0 },
        ilvl,
      })),
      target: { type: "new", num },
    },
    mode,
  });
  if (compiled.isErr()) return Result.err(compiled.error);
  // A pending paragraph join can redirect a logical edit to another mark.
  const rebound = new Set<string>();
  for (const op of compiled.value.ops) {
    if (op.type !== DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS) continue;
    const blockId = idKey(op.blockId);
    const reference = op.patch.numPr;
    if (
      reference?.kind !== "reference" ||
      reference.numId !== num.numId ||
      levels.get(blockId) !== (reference.ilvl ?? 0)
    ) {
      return refuse(
        DOCUMENT_OP_REFUSAL_REASONS.REVISION_CONFLICT,
        "A pending paragraph join changes the selected list's property owner.",
      );
    }
    rebound.add(blockId);
  }
  if (rebound.size !== levels.size)
    return refuse(
      DOCUMENT_OP_REFUSAL_REASONS.REVISION_CONFLICT,
      "A pending paragraph join combines selected list members.",
    );
  return applyDocumentOps(document, compiled.value.ops);
};
