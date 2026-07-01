/**
 * Headless `.docx` review path: buffer -> apply AI edits -> buffer, with
 * no `EditorView` and no DOM. A queue worker or agent can read a document,
 * apply `FolioAIEditOperation`s (tracked-changes or direct), and write a
 * reviewed `.docx` back out — the server-side counterpart to the React
 * editor's live apply flow.
 *
 * The operation applier is shared, not forked: {@link applyFolioAIEditOperations}
 * only needs a {@link FolioAIEditView} seam (`{ state, dispatch }`), which a
 * headless `EditorState` satisfies via `state.apply(tr)`. The React editor and
 * this reviewer therefore run byte-for-byte the same anchor resolution,
 * word-diff redlines, and tracked-change bookkeeping.
 *
 * Save mirrors the editor's own path: a selective patch of only the changed
 * paragraphs in `document.xml` (leaving every untouched part byte-exact), with
 * a full repack as the fallback for structural edits.
 *
 * Scope: BODY blocks. Footnote / endnote note-body apply is intentionally out
 * of scope here and can build on the note serialization work separately.
 */

import { EditorState } from "prosemirror-state";
import type { Plugin } from "prosemirror-state";
import type { Transaction } from "prosemirror-state";

import { attemptSelectiveSave } from "../docx/selectiveSave";
import { parseDocx } from "../docx/parser";
import { repackDocx } from "../docx/rezip";
import { updateDocumentContent } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import { ensureParaIdsInState } from "../prosemirror/extensions/features/ParaIdAllocatorExtension";
import {
  getChangedParagraphIds,
  hasStructuralChanges,
  hasUntrackedChanges,
} from "../prosemirror/extensions/features/ParagraphChangeTrackerExtension";
import { schema, singletonManager } from "../prosemirror/schema";
import type { Comment } from "../types/content";
import type { Document } from "../types/document";
import { applyFolioAIEditOperations } from "./apply";
import { createFolioAIEditSnapshot } from "./snapshot";
import type {
  FolioAIEditApplyMode,
  FolioAIEditApplyResult,
  FolioAIEditOperation,
  FolioAIEditSnapshot,
} from "./types";

/**
 * Standalone comment-id allocator for the reviewer. The React editor uses
 * `Date.now()`-seeded ids from `commentsHelpers`; that module is DOM-bound
 * (`EditorView`, `findBodyPmAnchors`) and unimportable here, so the reviewer
 * mints its own. Seeded once per realm and incremented so ids stay unique
 * across reviewers created within the same millisecond.
 */
let commentIdCursor = Date.now();

/**
 * Build the note-free comment thread the apply layer references by id.
 * Mirrors `commentsHelpers.createComment` (a pure object literal there) so a
 * headless `commentOnBlock` / `comment` op serialises the same `comments.xml`
 * shape the editor produces.
 */
const createReviewerComment = (text: string, author: string): Comment => ({
  id: commentIdCursor++,
  author,
  date: new Date().toISOString(),
  content: [
    {
      type: "paragraph",
      formatting: {},
      content: [{ type: "run", formatting: {}, content: [{ type: "text", text }] }],
    },
  ],
});

/** Options for {@link FolioDocxReviewer.fromBuffer}. */
export type FolioDocxReviewerOptions = {
  /** Default author for tracked changes and comments. (default: `"AI"`) */
  author?: string;
};

/** Options for {@link FolioDocxReviewer.applyOperations}. */
export type FolioApplyOperationsOptions = {
  /** `"tracked-changes"` (default) produces ins/del redlines; `"direct"` edits in place. */
  mode?: FolioAIEditApplyMode;
  /**
   * The snapshot the `operations`' block ids were built against. Omit to
   * snapshot the reviewer's current state — correct for the common flow of
   * `snapshot()` -> build ops -> `applyOperations()` on the same reviewer.
   */
  snapshot?: FolioAIEditSnapshot;
};

/**
 * Headless `.docx` reviewer. Parse a buffer, read blocks, apply
 * `FolioAIEditOperation`s against the document model, and write the reviewed
 * `.docx` back out — no editor instance, no DOM.
 *
 * @example
 * ```ts
 * const reviewer = await FolioDocxReviewer.fromBuffer(buffer, { author: "AI" });
 * const { blocks } = reviewer.snapshot();
 * reviewer.applyOperations([
 *   { id: "1", type: "replaceInBlock", blockId: blocks[0].id, find: "$50k", replace: "$500k" },
 * ]);
 * const reviewed = await reviewer.toBuffer();
 * ```
 */
export class FolioDocxReviewer {
  /** Default author for tracked changes and comments. */
  readonly author: string;
  private readonly baseDocument: Document;
  private readonly originalBuffer: ArrayBuffer;
  private state: EditorState;
  private readonly createdComments: Comment[] = [];

  private constructor(args: {
    baseDocument: Document;
    originalBuffer: ArrayBuffer;
    state: EditorState;
    author: string;
  }) {
    this.baseDocument = args.baseDocument;
    this.originalBuffer = args.originalBuffer;
    this.state = args.state;
    this.author = args.author;
  }

  /** Parse a `.docx` buffer into a reviewer. */
  static async fromBuffer(
    buffer: ArrayBuffer,
    options: FolioDocxReviewerOptions = {},
  ): Promise<FolioDocxReviewer> {
    const baseDocument = await parseDocx(buffer, {
      detectVariables: false,
      preloadFonts: false,
    });
    // Same plugin set the live editor mounts: the change tracker feeds the
    // selective-save key set, and the paraId allocator hands freshly inserted
    // paragraphs a stable `w14:paraId`. No plugin `view()` runs headlessly, so
    // the DOM-facing halves stay dormant.
    const plugins: Plugin[] = singletonManager.getPlugins();
    // Allocate paraIds up front (the editor does this on load) so every block
    // anchors on a stable id and the selective-save path can key changed
    // paragraphs by paraId. `ensureParaIdsInState` marks its transaction
    // ignore-tracked, so the change baseline stays clean.
    const state = ensureParaIdsInState(
      EditorState.create({ schema, doc: toProseDoc(baseDocument), plugins }),
    );
    return new FolioDocxReviewer({
      baseDocument,
      originalBuffer: buffer,
      state,
      author: options.author ?? "AI",
    });
  }

  /**
   * Snapshot the current document into AI-facing blocks (id, kind, text) plus
   * the anchor map the apply layer resolves operations against.
   */
  snapshot(): FolioAIEditSnapshot {
    return createFolioAIEditSnapshot(this.state.doc);
  }

  /**
   * Apply operations against the current state. Reuses the live-editor applier
   * verbatim via a headless `{ state, dispatch }` seam; the resulting state
   * (including any comments) is retained for {@link toBuffer}.
   */
  applyOperations(
    operations: FolioAIEditOperation[],
    options: FolioApplyOperationsOptions = {},
  ): FolioAIEditApplyResult {
    const snapshot = options.snapshot ?? this.snapshot();
    const view = {
      state: this.state,
      dispatch: (transaction: Transaction) => {
        view.state = view.state.apply(transaction);
      },
    };

    const result = applyFolioAIEditOperations({
      view,
      snapshot,
      operations,
      mode: options.mode ?? "tracked-changes",
      author: this.author,
      createCommentId: (text) => {
        const comment = createReviewerComment(text, this.author);
        this.createdComments.push(comment);
        return comment.id;
      },
    });

    this.state = view.state;
    return result;
  }

  /** The current document model with edits merged back in. */
  toDocument(): Document {
    const document = updateDocumentContent(this.baseDocument, this.state.doc);
    if (this.createdComments.length > 0) {
      document.package.document.comments = [
        ...(document.package.document.comments ?? []),
        ...this.createdComments,
      ];
    }
    return document;
  }

  /**
   * Serialise the reviewed document to a new `.docx` buffer. Tries a selective
   * patch first (only changed paragraphs are rewritten; every other byte of the
   * original package is preserved), falling back to a full repack for
   * structural edits — the same two-tier path the editor's save uses.
   */
  async toBuffer(): Promise<ArrayBuffer> {
    const document = this.toDocument();
    const selective = await attemptSelectiveSave(document, this.originalBuffer, {
      changedParaIds: getChangedParagraphIds(this.state),
      structuralChange: hasStructuralChanges(this.state),
      hasUntrackedChanges: hasUntrackedChanges(this.state),
    });
    if (selective) {
      return selective;
    }
    return repackDocx({ ...document, originalBuffer: this.originalBuffer });
  }
}

/** Options for {@link applyFolioAIEditsToBuffer}. */
export type ApplyFolioAIEditsToBufferOptions = {
  /** Default author for tracked changes and comments. (default: `"AI"`) */
  author?: string;
  /** `"tracked-changes"` (default) produces ins/del redlines; `"direct"` edits in place. */
  mode?: FolioAIEditApplyMode;
  /**
   * The snapshot the `operations`' block ids were built against. Omit to
   * snapshot the freshly parsed buffer. Pass it when the ids were derived in a
   * separate process (e.g. server-side `@stll/folio-core/server` block ids) so
   * anchor resolution matches exactly.
   */
  snapshot?: FolioAIEditSnapshot;
};

/** Result of {@link applyFolioAIEditsToBuffer}. */
export type ApplyFolioAIEditsToBufferResult = FolioAIEditApplyResult & {
  /** The reviewed `.docx` as a new buffer. */
  buffer: ArrayBuffer;
};

/**
 * One-shot headless review: parse `buffer`, apply `operations`, and return the
 * reviewed `.docx` buffer alongside the applied / skipped breakdown. Convenience
 * wrapper over {@link FolioDocxReviewer} for callers that already hold the
 * operations to run.
 */
export const applyFolioAIEditsToBuffer = async (
  buffer: ArrayBuffer,
  operations: FolioAIEditOperation[],
  options: ApplyFolioAIEditsToBufferOptions = {},
): Promise<ApplyFolioAIEditsToBufferResult> => {
  const reviewer = await FolioDocxReviewer.fromBuffer(buffer, {
    ...(options.author !== undefined && { author: options.author }),
  });
  const { applied, skipped } = reviewer.applyOperations(operations, {
    ...(options.mode !== undefined && { mode: options.mode }),
    ...(options.snapshot !== undefined && { snapshot: options.snapshot }),
  });
  const reviewed = await reviewer.toBuffer();
  return { buffer: reviewed, applied, skipped };
};
