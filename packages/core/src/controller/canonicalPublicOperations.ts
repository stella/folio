/** Public batches compile against the read projection; only the canonical journal publishes. */
import { panic, Result, TaggedError } from "better-result";
import {
  applyDocumentOps,
  freshCommentId,
  findStoryBody,
  paragraphLength,
  type TextPosition,
  compileEditorIntent,
  createEditorIntentIdAllocator,
  documentStories,
  type DocumentOp,
  type EditorIntent,
  type OpStory,
  type RunPropsPatch,
} from "@stll/docx-core/ops";
import type { Document } from "../types/document";
import { stripInlineEmphasisMarkers } from "../ai-edits/inline-emphasis";
import { wordDiffSessionFromOptions } from "../ai-edits/word-diff";
import { Mark } from "prosemirror-model";
import type { EditorState } from "prosemirror-state";
import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  getFolioDocumentOperationIssues,
  getFolioDocumentOperationReceiptsForStory,
  parseFolioDocumentOperationBatch,
  type ApplyFolioDocumentOperationsOptions,
  type FolioDocumentOperation,
  type FolioDocumentOperationResult,
  type FolioDocumentOperationStory,
  type FolioDocumentOperationBatch,
  type FolioDocumentOperationUndoHandle,
  type FolioDocumentOperationUndoResult,
} from "../document-operations";
import { buildCleanBlockText, resolveCleanTextRange } from "../ai-edits/clean-text";
import { collectNoteReferenceLabels, hashFolioAIBlockText } from "../ai-edits/snapshot";
import type {
  FolioAIEditSkippedOperation,
  FolioAIEditSkipReason,
  FolioAIEditNormalization,
} from "../ai-edits/types";
import { maxAnnotationIdInDoc } from "../prosemirror/plugins/revisionIds";
import { splitsGraphemeCluster, splitsSurrogatePair } from "../ai-edits/character-boundaries";
import { type CanonicalCommit, type CanonicalSession } from "./canonicalSession";

import { canonicalCommentBody, compileCanonicalComments } from "./canonicalComments";
import { CANONICAL_GAP, type CanonicalGap } from "../types/canonicalCapabilities";

type CanonicalPublicOperationRefusal = NonNullable<FolioAIEditSkippedOperation["canonicalRefusal"]>;

/** A new public kind cannot bypass a canonical compiler/refusal decision. */
export const CANONICAL_PUBLIC_OPERATION_DISPOSITIONS = {
  replaceInBlock: "compile",
  replaceRange: "compile",
  replaceBlock: "compile",
  splitBlock: "compile",
  formatRange: "compile",
  mergeBlockWithNext: "compile",
  setBlockParagraphProperties: CANONICAL_GAP.publicUnsupportedInline,
  insertAfterBlock: CANONICAL_GAP.publicUnsupportedInline,
  insertBeforeBlock: CANONICAL_GAP.publicUnsupportedInline,
  deleteBlock: CANONICAL_GAP.publicUnsupportedInline,
  commentOnBlock: "compile",
  commentOnRange: "compile",
  insertTable: CANONICAL_GAP.publicTableProjection,
  insertSignatureTable: CANONICAL_GAP.publicTableProjection,
  deleteTable: CANONICAL_GAP.publicTableProjection,
  insertTableRow: CANONICAL_GAP.publicTableProjection,
  deleteTableRow: CANONICAL_GAP.publicTableProjection,
  insertTableColumn: CANONICAL_GAP.publicTableProjection,
  deleteTableColumn: CANONICAL_GAP.publicTableProjection,
  mergeTableCells: CANONICAL_GAP.publicTableProjection,
  splitTableCell: CANONICAL_GAP.publicTableProjection,
} as const satisfies Record<FolioDocumentOperation["type"], "compile" | CanonicalGap>;

export type CanonicalPublicOperationOptions = Omit<
  ApplyFolioDocumentOperationsOptions,
  "view" | "createUndoHandle" | "createCommentId"
>;

type ResolvedOperation = {
  operation: FolioDocumentOperation;
  intents: readonly EditorIntent[];
  comment?: { text: string; from: TextPosition; to: TextPosition };
  normalizations: readonly FolioAIEditNormalization[];
  from: number;
  to: number;
};

type CanonicalPublicOperationsOptions = {
  session: CanonicalSession;
  getState: (story: OpStory) => EditorState | null;
  publish: (commit: CanonicalCommit) => boolean;
};

type CompilePublicIntentsOptions = {
  document: Document;
  intents: readonly EditorIntent[];
  mode: "direct" | "tracked-changes";
  author: string;
  date: string;
  nextRevisionId: number;
  allocate: ReturnType<typeof createEditorIntentIdAllocator>;
};

const compilePublicIntents = ({
  document,
  intents,
  mode,
  author,
  date,
  nextRevisionId,
  allocate,
}: CompilePublicIntentsOptions) => {
  const ops: DocumentOp[] = [];
  const revisions: number[] = [];
  for (const sourceIntent of intents) {
    const ids = allocate(document, sourceIntent);
    const intent =
      sourceIntent.type === "splitParagraph"
        ? ({
            type: "splitParagraph",
            at: sourceIntent.at,
            ...(sourceIntent.to === undefined ? {} : { to: sourceIntent.to }),
            newBlockId: ids.newBlockId,
          } as const)
        : sourceIntent;
    const firstRevisionId = nextRevisionId;
    const compiled = compileEditorIntent(document, {
      intent,
      mode:
        mode === "direct"
          ? {
              type: "editing",
              newIds: ids.newIds,
            }
          : {
              type: "suggesting",
              revision: { id: firstRevisionId, author, date },
              newIds: {
                control: ids.newIds.control,
                revision: ids.newIds.revision.map((_, index) => firstRevisionId + index + 1),
              },
            },
    });
    if (compiled.isErr()) return Result.err(compiled.error);
    const applied = applyDocumentOps(document, compiled.value.ops);
    if (applied.isErr()) return Result.err(applied.error);
    document = applied.value.document;
    ops.push(...compiled.value.ops);
    revisions.push(...applied.value.revisions);
    for (const id of applied.value.revisions) nextRevisionId = Math.max(nextRevisionId, id + 1);
  }
  return Result.ok({ document, ops, revisions, nextRevisionId });
};

class CanonicalPublicOperationResolutionError extends TaggedError(
  "CanonicalPublicOperationResolutionError",
)<{
  message: string;
  reason: FolioAIEditSkipReason;
}> {}

type ResolveCanonicalPublicOperationOptions = {
  operation: FolioDocumentOperation;
  options: CanonicalPublicOperationOptions;
  state: EditorState;
  story: OpStory;
  diff: ReturnType<typeof wordDiffSessionFromOptions>["diff"];
};

const canonicalPublicStory = (story: FolioDocumentOperationStory): OpStory => {
  if (story === "main") return story;
  switch (story.type) {
    case "header":
    case "footer":
      return { kind: story.type, rId: story.relationshipId };
    case "footnote":
    case "endnote":
      return { kind: story.type, id: story.noteId };
    default: {
      const unreachable: never = story;
      return panic(`Unknown public operation story ${unreachable}`);
    }
  }
};

const batchStatus = (batch: FolioDocumentOperationBatch, skippedCount: number) => {
  if (batch.dryRun) return "previewed" as const;
  if (batch.atomic && skippedCount > 0) return "rejected" as const;
  return "committed" as const;
};

/** Handles address isolated journal groups; no second document or PM history is retained. */
export class CanonicalPublicOperations {
  private readonly options: CanonicalPublicOperationsOptions;
  private readonly handles: {
    handle: FolioDocumentOperationUndoHandle;
    version: number;
    story: OpStory;
  }[] = [];
  private nextHandle = 1;
  private readonly handleNamespace = crypto.randomUUID();

  constructor(options: CanonicalPublicOperationsOptions) {
    this.options = options;
  }

  apply(options: CanonicalPublicOperationOptions): FolioDocumentOperationResult {
    const batch = parseFolioDocumentOperationBatch(options.batch);
    const { session } = this.options;
    const publicStory = options.story ?? "main";
    const story = canonicalPublicStory(publicStory);
    const state = this.options.getState(story);
    const projected = session.projectStory(story);
    const skipped: FolioAIEditSkippedOperation[] = [];
    const normalizations: FolioAIEditNormalization[] = [];
    const refusals = new Map<string, CanonicalPublicOperationRefusal>();
    const resolved: ResolvedOperation[] = [];
    const { diff } = wordDiffSessionFromOptions(options.wordDiff);
    let nextRevisionId = options.revisionStamp?.idSeed ?? 1;
    if (!session.isComposing && options.revisionStamp === undefined) {
      for (const currentStory of documentStories(session.document)) {
        const currentProjection = session.projectStory(currentStory);
        if (currentProjection.isOk())
          nextRevisionId = Math.max(
            nextRevisionId,
            maxAnnotationIdInDoc(currentProjection.value.doc) + 1,
          );
      }
    }
    const skip = (id: string, reason: FolioAIEditSkipReason, message?: string) => {
      skipped.push({ id, reason, ...(message === undefined ? {} : { message }) });
    };
    const result = (
      applied: FolioDocumentOperationResult["applied"],
      undoHandle: FolioDocumentOperationUndoHandle | null,
    ): FolioDocumentOperationResult => {
      const failures = new Map(skipped.map((item) => [item.id, item]));
      const ordered = batch.operations.flatMap(({ id }) => {
        const failure = failures.get(id);
        if (!failure) return [];
        const canonicalRefusal = refusals.get(id);
        return [canonicalRefusal ? { ...failure, canonicalRefusal } : failure];
      });
      return {
        version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
        status: batchStatus(batch, ordered.length),
        applied,
        skipped: ordered,
        issues: getFolioDocumentOperationIssues(batch.operations, ordered),
        receipts: getFolioDocumentOperationReceiptsForStory({
          operations: batch.operations,
          applied,
          story: publicStory,
        }),
        undoHandle,
        nextRevisionId,
        ...(normalizations.length === 0 ? {} : { normalizations }),
      };
    };
    for (const operation of batch.operations) {
      if (!state || projected.isErr() || session.isComposing) {
        skip(operation.id, "documentNotEditable");
        continue;
      }
      if (batch.mode === "suggested") {
        refusals.set(operation.id, {
          gap: CANONICAL_GAP.publicSuggestedMode,
        });
        skip(operation.id, "unsupportedMode", "Canonical pending suggestions are unavailable.");
        continue;
      }
      if (options.tableTemplates !== undefined || options.undefinedStyles === "keep") {
        refusals.set(operation.id, {
          gap: CANONICAL_GAP.publicUnsupportedInline,
        });
        skip(
          operation.id,
          "unsupportedBlock",
          "This canonical compiler does not support the supplied execution options.",
        );
        continue;
      }
      const disposition = CANONICAL_PUBLIC_OPERATION_DISPOSITIONS[operation.type];
      if (disposition !== "compile") {
        refusals.set(operation.id, {
          gap: disposition,
        });
        skip(operation.id, "unsupportedBlock");
        continue;
      }
      const compiled = this.resolve({ operation, options, state, story, diff });
      if (compiled.isErr()) {
        if (compiled.error.reason === "unsupportedBlock")
          refusals.set(operation.id, {
            gap: CANONICAL_GAP.publicUnsupportedInline,
          });
        skip(operation.id, compiled.error.reason, compiled.error.message);
        continue;
      }
      if (
        resolved.some(
          ({ from, to }) =>
            (compiled.value.from < to && compiled.value.to > from) ||
            (from === to && compiled.value.from <= from && compiled.value.to >= from) ||
            (compiled.value.from === compiled.value.to &&
              compiled.value.from >= from &&
              compiled.value.from <= to),
        )
      ) {
        skip(operation.id, "overlappingOperation");
        continue;
      }
      resolved.push(compiled.value);
    }
    if (batch.atomic && skipped.length > 0) {
      for (const { operation } of resolved) skip(operation.id, "atomicBatchRejected");
      return result([], null);
    }
    if (!state || projected.isErr() || session.isComposing) return result([], null);
    // Descending read positions keep independent snapshot coordinates valid.
    resolved.sort((left, right) => right.from - left.from);
    let document = session.document;
    const ops: DocumentOp[] = [];
    const appliedById = new Map<string, FolioDocumentOperationResult["applied"][number]>();
    const allocate = createEditorIntentIdAllocator();
    const date = options.revisionStamp?.date ?? new Date().toISOString();
    for (const { operation, intents, comment, normalizations: adjustments } of resolved) {
      const compiled = compilePublicIntents({
        document,
        intents,
        allocate,
        nextRevisionId,
        mode: batch.mode === "direct" ? "direct" : "tracked-changes",
        author: options.author ?? "AI",
        date,
      });
      if (compiled.isErr()) {
        refusals.set(operation.id, {
          gap: CANONICAL_GAP.publicUnsupportedInline,
        });
        skip(operation.id, "unsupportedBlock", compiled.error.message);
        continue;
      }
      let nextDocument = compiled.value.document;
      const operationOps = [...compiled.value.ops];
      let commentId: number | undefined;
      if (comment) {
        const id = freshCommentId(nextDocument);
        const before = findStoryBody(document, story)?.content.find(
          (block) => block.type === "paragraph" && block.paraId === comment.from.blockId,
        );
        const after = findStoryBody(nextDocument, story)?.content.find(
          (block) => block.type === "paragraph" && block.paraId === comment.from.blockId,
        );
        if (id.isErr() || before?.type !== "paragraph" || after?.type !== "paragraph") {
          refusals.set(operation.id, { gap: CANONICAL_GAP.publicComments });
          skip(
            operation.id,
            "unsupportedBlock",
            id.isErr() ? id.error.message : "The comment anchor no longer exists.",
          );
          continue;
        }
        const to = {
          ...comment.to,
          offset: comment.to.offset + paragraphLength(after) - paragraphLength(before),
        };
        const commented = compileCanonicalComments({
          document: nextDocument,
          command: {
            type: "create",
            comment: {
              id: id.value,
              author: options.author ?? "AI",
              date,
              done: false,
              content: canonicalCommentBody(nextDocument, comment.text),
            },
            anchor:
              to.offset === comment.from.offset
                ? { kind: "point", at: comment.from }
                : { kind: "range", from: comment.from, to },
          },
        });
        if (commented.isErr()) {
          refusals.set(operation.id, { gap: CANONICAL_GAP.publicComments });
          skip(operation.id, "unsupportedBlock", commented.error.message);
          continue;
        }
        nextDocument = commented.value.document;
        operationOps.push(...commented.value.ops);
        commentId = id.value;
      }
      if (operationOps.length === 0) {
        skip(operation.id, "noopOperation");
        continue;
      }
      normalizations.push(...adjustments);
      document = nextDocument;
      ops.push(...operationOps);
      nextRevisionId = Math.max(compiled.value.nextRevisionId, (commentId ?? -1) + 1);
      const revisionIds = compiled.value.revisions;
      const revisionId = revisionIds.at(0);
      appliedById.set(operation.id, {
        id: operation.id,
        ...(commentId === undefined ? {} : { commentId }),
        ...(revisionId === undefined ? {} : { revisionId, revisionIds }),
      });
    }
    if (batch.atomic && skipped.length > 0) {
      for (const { id } of appliedById.values()) skip(id, "atomicBatchRejected");
      return result([], null);
    }
    if (ops.length === 0) return result([], null);
    const selection = projected.value.selectionAt(state);
    const prepared = selection.isErr()
      ? selection
      : session.prepareOps(state, ops, selection.value);
    if (prepared.isErr()) {
      for (const { id } of appliedById.values()) {
        refusals.set(id, {
          gap: CANONICAL_GAP.publicUnsupportedInline,
        });
        skip(
          id,
          prepared.error.reason === "noChange" ? "noopOperation" : "unsupportedBlock",
          prepared.error.message,
        );
      }
      return result([], null);
    }
    const applied = batch.operations.flatMap(({ id }) => {
      const entry = appliedById.get(id);
      return entry ? [batch.dryRun ? { id } : entry] : [];
    });
    if (batch.dryRun) return result(applied, null);
    if (!this.options.publish(prepared.value)) {
      for (const { id } of applied) skip(id, "documentNotEditable");
      return result([], null);
    }
    const handle = {
      type: "documentOperationUndo",
      id: `canonical-${this.handleNamespace}-${String(this.nextHandle++)}`,
    } as const;
    this.handles.push({ handle, version: session.version, story });
    return result(applied, handle);
  }

  undo(handle: FolioDocumentOperationUndoHandle): FolioDocumentOperationUndoResult {
    const index = this.handles.findIndex(
      ({ handle: existing }) => existing.id === handle.id && existing.type === handle.type,
    );
    if (index < 0) return { status: "rejected", undoHandle: handle, reason: "unknownHandle" };
    if (index !== this.handles.length - 1)
      return { status: "rejected", undoHandle: handle, reason: "notLatest" };
    const entry = this.handles.at(-1);
    const { session } = this.options;
    if (!entry || entry.version !== session.version)
      return { status: "rejected", undoHandle: handle, reason: "documentChanged" };
    const state = this.options.getState(entry.story);
    if (!state) return { status: "rejected", undoHandle: handle, reason: "documentChanged" };
    const prepared = session.prepareUndo(state, entry.story);
    if (prepared.isErr() || !this.options.publish(prepared.value))
      return { status: "rejected", undoHandle: handle, reason: "documentChanged" };
    this.handles.pop();
    // Undo advances the version while exposing the preceding unchanged journal group.
    const previous = this.handles.at(-1);
    if (previous && previous.version === entry.version - 1) previous.version = session.version;
    return { status: "undone", undoHandle: handle };
  }

  private resolve({
    operation,
    options,
    state,
    story,
    diff,
  }: ResolveCanonicalPublicOperationOptions) {
    const refusal = (reason: FolioAIEditSkipReason, message?: string) =>
      Result.err(
        new CanonicalPublicOperationResolutionError({
          reason,
          message: message ?? `The operation was refused: ${reason}.`,
        }),
      );
    switch (operation.type) {
      case "replaceInBlock":
      case "replaceRange":
      case "replaceBlock":
      case "splitBlock":
      case "formatRange":
      case "mergeBlockWithNext":
      case "commentOnRange":
      case "commentOnBlock":
        break;
      default:
        return refusal(
          "unsupportedBlock",
          `Canonical public operation ${operation.type} is unavailable.`,
        );
    }
    const blockId =
      operation.type === "replaceRange" ||
      operation.type === "formatRange" ||
      operation.type === "commentOnRange"
        ? operation.range.blockId
        : operation.blockId;
    const anchor = Object.hasOwn(options.snapshot.anchors, blockId)
      ? options.snapshot.anchors[blockId]
      : undefined;
    const projection = this.options.session.projectStory(story);
    if (!anchor || projection.isErr()) return refusal("missingBlock");
    const paragraph = projection.value.paragraph(blockId);
    if (!paragraph) return refusal("missingBlock");
    if (paragraph.source.pPrMark?.kind === "del" || paragraph.source.pPrMark?.kind === "moveFrom")
      return refusal("pendingDeletion");
    const clean = buildCleanBlockText(paragraph.node, paragraph.start - 1, {
      fieldResults: "text",
      noteReferences: collectNoteReferenceLabels(state.doc),
    });
    const hash = hashFolioAIBlockText(clean.text);
    if (operation.precondition && hash !== operation.precondition.blockTextHash)
      return refusal("preconditionFailed");
    if (hash !== anchor.textHash) return refusal("changedBlock");
    let start = 0;
    let end = clean.text.length;
    let text = "";
    switch (operation.type) {
      case "replaceInBlock": {
        if (operation.find === "") return refusal("emptyOperation");
        start = clean.text.indexOf(operation.find);
        if (start < 0) return refusal("missingFind");
        if (clean.text.indexOf(operation.find, start + 1) >= 0) return refusal("ambiguousFind");
        end = start + operation.find.length;
        text = operation.replace;
        break;
      }
      case "commentOnBlock":
        if (operation.quote !== undefined) {
          start = clean.text.indexOf(operation.quote);
          if (start < 0) return refusal("missingFind");
          if (clean.text.indexOf(operation.quote, start + 1) >= 0) return refusal("ambiguousFind");
          end = start + operation.quote.length;
        }
        break;
      case "replaceRange":
      case "formatRange":
      case "commentOnRange":
        start = operation.range.startOffset;
        end = operation.range.endOffset;
        if (hashFolioAIBlockText(clean.text.slice(start, end)) !== operation.range.selectedTextHash)
          return refusal("staleRange");
        text = operation.type === "replaceRange" ? operation.replace : "";
        break;
      case "replaceBlock":
        if (operation.styleId !== undefined || operation.preserveFormatting === false)
          return refusal(
            "unsupportedBlock",
            "Canonical replacement paragraph properties are unavailable.",
          );
        text = operation.text;
        break;
      case "splitBlock":
        if (operation.firstParagraphProperties || operation.secondParagraphProperties)
          return refusal(
            "unsupportedBlock",
            "Canonical split paragraph properties are unavailable.",
          );
        start = operation.offset;
        end = start + (operation.separator?.length ?? 0);
        if (
          start <= 0 ||
          start >= clean.text.length ||
          clean.text.slice(start, end) !== (operation.separator ?? "")
        )
          return refusal("staleRange");
        break;
      case "mergeBlockWithNext":
        if (operation.mergedParagraphProperties || operation.separator)
          return refusal(
            "unsupportedBlock",
            "Canonical public join properties and separators are unavailable.",
          );
        start = clean.text.length;
        end = start;
        break;
      default: {
        const unreachable: never = operation;
        return unreachable;
      }
    }
    if (
      [start, end].some(
        (offset) =>
          splitsSurrogatePair(clean.text, offset) || splitsGraphemeCluster(clean.text, offset),
      )
    )
      return refusal("splitsCharacter");
    if (
      clean.structuralBoundaries.some(
        (boundary) =>
          boundary.type === "noteReference" &&
          start < boundary.offset + boundary.length &&
          end > boundary.offset,
      )
    )
      return refusal("protectedReference");
    if (stripInlineEmphasisMarkers(text) !== text)
      return refusal("unsupportedBlock", "Canonical public emphasis replacements are unavailable.");
    if (/\[\^(?:e)?\d+\]/u.test(text) || /[\t\r\n]/u.test(text))
      return refusal("unsupportedBlock", "Canonical public replacements require plain text.");
    const range = resolveCleanTextRange({ cleanBlock: clean, startOffset: start, endOffset: end });
    if (!range) return refusal("staleRange");
    const from = projection.value.addressAt(range.from);
    const to = projection.value.addressAt(range.to);
    if (from.isErr() || to.isErr()) return refusal("unsupportedBlock");
    const comment =
      "comment" in operation && operation.comment !== undefined
        ? { text: operation.comment.text, from: from.value, to: to.value }
        : undefined;
    if (operation.type === "commentOnBlock" || operation.type === "commentOnRange")
      return Result.ok({
        operation,
        intents: [],
        normalizations: [],
        ...(comment === undefined ? {} : { comment }),
        from: range.from,
        to: range.to,
      });
    if (
      comment === undefined &&
      operation.type !== "splitBlock" &&
      operation.type !== "formatRange" &&
      operation.type !== "mergeBlockWithNext" &&
      clean.text.slice(start, end) === text
    )
      return refusal("noopOperation");
    let intent: EditorIntent;
    switch (operation.type) {
      case "splitBlock":
        intent = {
          type: "splitParagraph",
          at: from.value,
          to: to.value,
          newBlockId: createEditorIntentIdAllocator()(this.options.session.document, {
            type: "splitParagraph",
            at: from.value,
            to: to.value,
          }).newBlockId,
        };
        break;
      case "mergeBlockWithNext": {
        const body = this.options.session.document.package.document.content;
        if (story !== "main")
          return refusal(
            "unsupportedBlock",
            "Canonical public joins currently require the main story.",
          );
        const next = body.at(
          body.findIndex((block) => block.type === "paragraph" && block.paraId === blockId) + 1,
        );
        if (next?.type !== "paragraph" || !next.paraId) return refusal("missingBlock");
        intent = { type: "joinParagraphs", story, blockId, nextBlockId: next.paraId };
        break;
      }
      case "formatRange": {
        const formatting = operation.formatting;
        const patch = {
          ...(formatting.bold === undefined ? {} : { bold: formatting.bold }),
          ...(formatting.italic === undefined ? {} : { italic: formatting.italic }),
          ...(formatting.underline === undefined
            ? {}
            : {
                underline:
                  formatting.underline === null
                    ? null
                    : { style: formatting.underline ? "single" : "none" },
              }),
          ...(formatting.strike === undefined
            ? {}
            : { strike: formatting.strike, doubleStrike: null }),
          ...(formatting.fontSizePt === undefined
            ? {}
            : { fontSize: formatting.fontSizePt === null ? null : formatting.fontSizePt * 2 }),
          ...(formatting.fontFamily === undefined
            ? {}
            : {
                fontFamily:
                  formatting.fontFamily === null
                    ? null
                    : { ascii: formatting.fontFamily, hAnsi: formatting.fontFamily },
              }),
          ...(formatting.color === undefined
            ? {}
            : {
                color:
                  formatting.color === null
                    ? null
                    : { rgb: formatting.color.replace(/^#/u, "").toUpperCase() },
              }),
        } as const satisfies RunPropsPatch;
        intent = { type: "formatRun", from: from.value, to: to.value, patch };
        break;
      }
      case "replaceInBlock":
      case "replaceRange":
      case "replaceBlock":
        intent = {
          type: "replaceText",
          from: from.value,
          to: to.value,
          text,
          ...(options.replacementBackground === "keep"
            ? {}
            : { runPropsPatch: { highlight: null, shading: null } }),
        };
        break;
      default: {
        const unreachable: never = operation;
        return panic(`Unknown canonical public operation ${unreachable}`);
      }
    }
    const intents: EditorIntent[] = [];
    const normalizations: FolioAIEditNormalization[] = [];
    if (intent.type === "replaceText") {
      let offset = start;
      let pending: { start: number; end: number; text: string } | null = null;
      const changes: { start: number; end: number; text: string }[] = [];
      for (const segment of diff(clean.text.slice(start, end), text)) {
        switch (segment.type) {
          case "equal":
            if (pending) changes.push(pending);
            pending = null;
            offset += segment.text.length;
            break;
          case "del":
            pending ??= { start: offset, end: offset, text: "" };
            offset += segment.text.length;
            pending.end = offset;
            break;
          case "ins":
            pending ??= { start: offset, end: offset, text: "" };
            pending.text += segment.text;
            break;
          default: {
            const unreachable: never = segment.type;
            return panic(`Unknown replacement diff ${unreachable}`);
          }
        }
      }
      if (pending) changes.push(pending);
      for (const change of changes.toReversed()) {
        const gap = resolveCleanTextRange({
          cleanBlock: clean,
          startOffset: change.start,
          endOffset: change.end,
        });
        if (!gap) return refusal("unsupportedBlock");
        if (change.text.length > 0 && gap.from < gap.to) {
          let first: readonly Mark[] | null = null;
          let uniform = true;
          paragraph.node.nodesBetween(
            gap.from - paragraph.start,
            gap.to - paragraph.start,
            (node) => {
              if (!node.isText || node.marks.some((mark) => mark.type.name === "deletion")) return;
              const marks = node.marks.filter((mark) => mark.type.name !== "insertion");
              if (first === null) first = marks;
              else if (!Mark.sameSet(first, marks)) uniform = false;
            },
          );
          if (!uniform && normalizations.length === 0)
            normalizations.push({ id: operation.id, code: "uniformReplacementFormatting" });
        }
        const changeFrom = projection.value.addressAt(gap.from);
        const changeTo = projection.value.addressAt(gap.to);
        if (changeFrom.isErr() || changeTo.isErr()) return refusal("unsupportedBlock");
        intents.push({
          type: "replaceText",
          from: changeFrom.value,
          to: changeTo.value,
          text: change.text,
          ...(intent.runPropsPatch === undefined ? {} : { runPropsPatch: intent.runPropsPatch }),
        });
      }
    } else intents.push(intent);
    const structural = operation.type === "splitBlock" || operation.type === "mergeBlockWithNext";
    const nextParagraph =
      operation.type === "mergeBlockWithNext" && intent.type === "joinParagraphs"
        ? projection.value.paragraph(intent.nextBlockId)
        : undefined;
    let claimedTo = range.to;
    if (structural) claimedTo = paragraph.start + paragraph.node.content.size;
    if (nextParagraph) claimedTo = nextParagraph.start + nextParagraph.node.content.size;
    return Result.ok({
      operation,
      intents,
      ...(comment === undefined ? {} : { comment }),
      normalizations,
      from: structural ? paragraph.start : range.from,
      to: claimedTo,
    });
  }
}
