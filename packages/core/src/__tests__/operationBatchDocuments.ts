/**
 * Small packages and projections for the document-operation batch tests: build
 * paragraphs from runs, open them in a reviewer, save and reopen, accept, and
 * read back what a reader sees — each paragraph's text, its alignment, and
 * every character's formatting, link and comments.
 */

import type { Mark, Node as PMNode } from "prosemirror-model";
import type { EditorState, Transaction } from "prosemirror-state";

import { FolioDocxReviewer } from "../ai-edits/headless";
import { createFolioAIEditSnapshotWithStyleResolver } from "../ai-edits/snapshot";
import type { FolioAIEditSnapshot } from "../ai-edits/types";
import {
  applyFolioDocumentOperations,
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperation,
  type FolioDocumentOperationResult,
} from "../document-operations";
import { createDocx } from "../docx/rezip";
import { resolveAllChangesInHeadlessState } from "../prosemirror/commands/comments";
import { marksToTextFormatting } from "../prosemirror/conversion/fromProseDoc";
import { getDocumentStyleResolver } from "../prosemirror/plugins/documentStyleState";
import { isZeroWidthAnchor } from "../prosemirror/zeroWidthAnchors";
import type { Comment, Paragraph, ParagraphContent, TextFormatting } from "../types/document";

export const textRun = (text: string, formatting?: TextFormatting): ParagraphContent => ({
  type: "run",
  ...(formatting !== undefined && { formatting }),
  content: [{ type: "text", text }],
});

export const reviewComment = (id: number, text = `Note ${String(id)}`): Comment => ({
  id,
  author: "Dana Lindqvist",
  initials: "DL",
  date: "2024-01-01T00:00:00Z",
  content: [{ type: "paragraph", content: [textRun(text)] }],
});

/** A package whose paragraphs carry stable paraIds, so block ids survive edits. */
export const paragraphsDocx = (
  paragraphs: readonly (readonly ParagraphContent[])[],
  comments: readonly Comment[] = [],
): Promise<ArrayBuffer> => {
  const content: Paragraph[] = paragraphs.map((runs, index) => ({
    type: "paragraph",
    paraId: (0x1000_0000 + index).toString(16).toUpperCase(),
    content: [...runs],
  }));
  return createDocx({
    package: { document: { content, ...(comments.length > 0 && { comments: [...comments] }) } },
  });
};

export const openReviewer = (buffer: ArrayBuffer): Promise<FolioDocxReviewer> =>
  FolioDocxReviewer.fromBuffer(buffer, { author: "Reviewer" });

/** The reviewer's document, saved and opened again. */
export const reopened = async (reviewer: FolioDocxReviewer): Promise<FolioDocxReviewer> =>
  FolioDocxReviewer.fromBuffer(await reviewer.toBuffer(), { author: "Reviewer" });

/** Saved, reopened, every revision accepted, and saved and reopened once more. */
export const reopenedAccepted = async (reviewer: FolioDocxReviewer): Promise<FolioDocxReviewer> => {
  const opened = await reopened(reviewer);
  opened.acceptAll();
  return reopened(opened);
};

export const blockTexts = (reviewer: FolioDocxReviewer): string[] =>
  reviewer.snapshot().blocks.map((block) => block.text);

const canonical = (value: unknown): unknown => {
  if (typeof value !== "object" || value === null) {
    return value;
  }
  if (Array.isArray(value)) {
    return value.map(canonical);
  }
  return Object.fromEntries(
    Object.entries(value)
      .filter(([, entry]) => entry !== undefined)
      .toSorted(([left], [right]) => left.localeCompare(right))
      .map(([key, entry]) => [key, canonical(entry)]),
  );
};

const describeMarks = (marks: readonly Mark[], commentText: (id: number) => string): string => {
  const formatting = JSON.stringify(canonical(marksToTextFormatting(marks)));
  const link = marks.find((mark) => mark.type.name === "hyperlink")?.attrs["href"];
  const comments = marks
    .filter((mark) => mark.type.name === "comment")
    .map((mark) => {
      const id = mark.attrs["commentId"];
      return typeof id === "number" ? commentText(id) : "?";
    })
    .toSorted();
  const revisions = marks
    .filter((mark) => mark.type.name === "insertion" || mark.type.name === "deletion")
    .map((mark) => mark.type.name)
    .toSorted();
  return `${formatting}|${typeof link === "string" ? link : ""}|${comments.join(",")}|${revisions.join(",")}`;
};

export type ParagraphPresentation = {
  text: string;
  alignment: string | null;
  /** The kind of the paragraph's own tracked mark (`ins`, `del`), if it has one. */
  mark: string | null;
  /**
   * Runs of characters that present alike, as `text` and a description, and
   * every inline node that is not text (its type in the description).
   */
  spans: { text: string; presentation: string }[];
};

/**
 * What a reader sees of every paragraph of `doc`. Comments are named by their
 * text, not their id: two documents that allocated ids differently still read
 * the same.
 */
export const presentationOf = (
  doc: PMNode,
  commentText: (id: number) => string = (id) => String(id),
  { withAnchors = true }: { withAnchors?: boolean } = {},
): ParagraphPresentation[] => {
  const paragraphs: ParagraphPresentation[] = [];
  doc.descendants((node) => {
    if (!node.isTextblock) {
      return true;
    }
    const spans: ParagraphPresentation["spans"] = [];
    node.descendants((child) => {
      if (child.isInline && !child.isText && child.isLeaf) {
        if (!withAnchors && isZeroWidthAnchor(child)) {
          return false;
        }
        // A field, a tab, a break, an anchor: what it is and what it carries.
        spans.push({
          text: "",
          presentation: `<${child.type.name}>${describeMarks(child.marks, commentText)}`,
        });
        return false;
      }
      if (!child.isText) {
        return true;
      }
      const presentation = describeMarks(child.marks, commentText);
      const last = spans.at(-1);
      if (last?.presentation === presentation) {
        last.text += child.text ?? "";
      } else {
        spans.push({ text: child.text ?? "", presentation });
      }
      return false;
    });
    const alignment = node.attrs["alignment"];
    const paragraphMark: unknown = node.attrs["pPrMark"];
    const markKind =
      typeof paragraphMark === "object" && paragraphMark !== null
        ? Reflect.get(paragraphMark, "kind")
        : null;
    paragraphs.push({
      text: node.textContent,
      alignment: typeof alignment === "string" ? alignment : null,
      mark: typeof markKind === "string" ? markKind : null,
      spans,
    });
    return false;
  });
  return paragraphs;
};

/** {@link presentationOf} the reviewer's document, comments named by their text. */
export const reviewerPresentation = (reviewer: FolioDocxReviewer): ParagraphPresentation[] => {
  const texts = new Map(reviewer.getComments().map((comment) => [comment.id, comment.text]));
  return presentationOf(reviewer.state.doc, (id) => texts.get(id) ?? `#${String(id)}`);
};

/**
 * A document held in memory and edited through the same applier the reviewer
 * uses, without a package round trip per batch: what a test running hundreds
 * of batches needs. Comments are remembered by the text they were created
 * with.
 */
export class OperationSession {
  state: EditorState;
  private readonly comments = new Map<number, string>();
  private nextCommentId = 1;

  constructor(state: EditorState) {
    this.state = state;
  }

  snapshot(): FolioAIEditSnapshot {
    return createFolioAIEditSnapshotWithStyleResolver(
      this.state.doc,
      getDocumentStyleResolver(this.state),
    );
  }

  apply(
    mode: "direct" | "tracked-changes",
    operations: FolioDocumentOperation[],
  ): FolioDocumentOperationResult {
    const view = {
      state: this.state,
      dispatch: (transaction: Transaction) => {
        view.state = view.state.apply(transaction);
      },
    };
    const result = applyFolioDocumentOperations({
      view,
      snapshot: this.snapshot(),
      batch: { version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION, mode, operations },
      author: "Reviewer",
      createCommentId: (text) => {
        const id = this.nextCommentId++;
        this.comments.set(id, text);
        return id;
      },
    });
    this.state = view.state;
    return result;
  }

  acceptAll(): void {
    this.state = resolveAllChangesInHeadlessState(this.state, "accept");
  }

  /** See {@link presentationOf}; `withAnchors: false` leaves zero-width anchors out. */
  presentation(options: { withAnchors?: boolean } = {}): ParagraphPresentation[] {
    return presentationOf(
      this.state.doc,
      (id) => this.comments.get(id) ?? `#${String(id)}`,
      options,
    );
  }
}
