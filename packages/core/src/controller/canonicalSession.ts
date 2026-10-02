import { panic, Result, TaggedError } from "better-result";
import type { Node as PMNode } from "prosemirror-model";
import { TextSelection, type EditorState, type Transaction } from "prosemirror-state";
import {
  applyDocumentOps,
  DOCUMENT_OP_TYPES,
  normalizeForOps,
  OP_STORIES,
  validateOpsDocument,
  type DocumentOp,
  type TextPosition,
  type TouchedBlocks,
} from "@stll/docx-core/ops";
import { hasIllegalXmlCharacters } from "@stll/docx-core";

import { splitsGraphemeCluster, splitsSurrogatePair } from "../ai-edits/character-boundaries";
import {
  cloneDocumentWithParagraphPropertySources,
  copyDocumentParagraphPropertySourceContract,
  copyParagraphPropertySource,
} from "../docx/paragraphPropertySource";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import type { Document, Paragraph, StyleDefinitions, TextFormatting } from "../types/document";

export class CanonicalSessionError extends TaggedError("CanonicalSessionError")<{
  message: string;
}> {}

const refuse = (message: string) => Result.err(new CanonicalSessionError({ message }));

export type CanonicalSelection = { anchor: TextPosition; head: TextPosition };
export type CanonicalOrigin = "input" | "undo" | "redo";
export const CANONICAL_PROJECTION_META = "folioCanonicalProjection";

type ParagraphAddress = {
  blockId: string;
  start: number;
  text: string;
  node: PMNode;
  source: Paragraph;
};

/** Only plain paragraphs have a one-to-one UTF-16 address map. */
class CanonicalProjection {
  readonly doc: PMNode;
  private readonly paragraphs: readonly ParagraphAddress[];

  constructor(doc: PMNode, paragraphs: readonly ParagraphAddress[]) {
    this.doc = doc;
    this.paragraphs = paragraphs;
  }

  addressAt(position: number): Result<TextPosition, CanonicalSessionError> {
    if (!Number.isInteger(position)) return refuse("The input position is not an integer.");
    const paragraph = this.paragraphs.find(
      ({ start, text }) => position >= start && position <= start + text.length,
    );
    if (paragraph === undefined) return refuse("The input is outside a plain paragraph.");
    const offset = position - paragraph.start;
    if (splitsSurrogatePair(paragraph.text, offset)) {
      return refuse("The input would split a surrogate pair.");
    }
    return Result.ok({ story: OP_STORIES.MAIN, blockId: paragraph.blockId, offset });
  }

  positionAt(address: TextPosition): Result<number, CanonicalSessionError> {
    const paragraph = this.paragraphs.find(({ blockId }) => blockId === address.blockId);
    if (
      address.story !== OP_STORIES.MAIN ||
      (address.zeroWidthBefore !== undefined && address.zeroWidthBefore !== 0) ||
      paragraph === undefined ||
      !Number.isInteger(address.offset) ||
      address.offset < 0 ||
      address.offset > paragraph.text.length ||
      splitsSurrogatePair(paragraph.text, address.offset)
    ) {
      return refuse("The canonical selection is outside a plain paragraph.");
    }
    return Result.ok(paragraph.start + address.offset);
  }

  selectionAt(state: EditorState): Result<CanonicalSelection, CanonicalSessionError> {
    if (!(state.selection instanceof TextSelection)) {
      return refuse("Canonical input requires a text selection.");
    }
    const anchor = this.addressAt(state.selection.anchor);
    if (anchor.isErr()) return anchor;
    const head = this.addressAt(state.selection.head);
    if (head.isErr()) return head;
    return Result.ok({ anchor: anchor.value, head: head.value });
  }

  paragraph(blockId: string): ParagraphAddress | undefined {
    return this.paragraphs.find((paragraph) => paragraph.blockId === blockId);
  }
}

/** Text edits preserve the same paragraphs; private source captures follow their identities. */
const preservePropertySources = (target: Document, source: Document): void => {
  const sources = source.package.document.content;
  const targets = target.package.document.content;
  if (sources.length !== targets.length)
    panic("A canonical text edit changed paragraph structure.");
  for (const [index, original] of sources.entries()) {
    const derived = targets.at(index);
    if (
      original.type !== "paragraph" ||
      derived?.type !== "paragraph" ||
      derived.paraId !== original.paraId
    ) {
      panic("A canonical text edit changed paragraph ownership.");
    }
    if (derived !== original) copyParagraphPropertySource(derived, original);
  }
  copyDocumentParagraphPropertySourceContract(target, source);
};

type AuthoredInputFormattingOptions = {
  paragraph: Paragraph;
  offset: number;
  affinity: "before" | "after";
};

/** PM caret input inherits left; range replacement inherits the first replaced unit. */
const authoredInputFormatting = ({
  paragraph,
  offset,
  affinity,
}: AuthoredInputFormattingOptions): TextFormatting => {
  const unit = affinity === "before" && offset > 0 ? offset - 1 : offset;
  let end = 0;
  for (const run of paragraph.content) {
    if (run.type !== "run") panic("Canonical input encountered a non-text run.");
    for (const child of run.content) {
      if (child.type !== "text") panic("Canonical input encountered a non-text leaf.");
      end += child.text.length;
    }
    if (unit < end) return run.formatting ?? {};
  }
  if (end !== 0) panic("Canonical input could not find the authored insertion formatting.");
  return {};
};

const supportsSeed = (document: Document): boolean => {
  const pkg = document.package;
  const body = pkg.document;
  if (
    (pkg.headers?.size ?? 0) > 0 ||
    (pkg.footers?.size ?? 0) > 0 ||
    (pkg.footnotes?.length ?? 0) > 0 ||
    (pkg.endnotes?.length ?? 0) > 0 ||
    (body.comments?.length ?? 0) > 0 ||
    body.content.length === 0
  ) {
    return false;
  }
  return body.content.every(
    (paragraph) =>
      paragraph.type === "paragraph" &&
      paragraph.pPrMark === undefined &&
      paragraph.reviewCarrier === undefined &&
      (paragraph.propertyChanges?.length ?? 0) === 0 &&
      paragraph.content.every(
        (run) =>
          run.type === "run" &&
          (run.propertyChanges?.length ?? 0) === 0 &&
          run.content.every(
            (child) =>
              child.type === "text" &&
              !hasIllegalXmlCharacters(child.text) &&
              !/[\t\r\n]/u.test(child.text),
          ),
      ),
  );
};

const project = (
  document: Document,
  styles: StyleDefinitions | null | undefined,
): Result<CanonicalProjection, CanonicalSessionError> => {
  const converted = Result.try({
    try: () => toProseDoc(document, styles == null ? undefined : { styles }),
    catch: (cause) =>
      new CanonicalSessionError({
        message: `Canonical projection failed: ${cause instanceof Error ? cause.message : String(cause)}`,
      }),
  });
  if (converted.isErr()) return converted;
  const paragraphs: ParagraphAddress[] = [];
  let failure: CanonicalSessionError | undefined;
  converted.value.forEach((node, offset, index) => {
    const source = document.package.document.content.at(index);
    if (source?.type !== "paragraph" || source.paraId === undefined) {
      failure = new CanonicalSessionError({
        message: "The projection changed paragraph structure.",
      });
      return;
    }
    const text = source.content
      .flatMap((run) => (run.type === "run" ? run.content : []))
      .map((child) => (child.type === "text" ? child.text : ""))
      .join("");
    if (
      node.type.name !== "paragraph" ||
      node.attrs["paraId"] !== source.paraId ||
      node.textContent !== text ||
      node.content.size !== text.length
    ) {
      failure = new CanonicalSessionError({
        message: "The paragraph cannot be projected as plain text.",
      });
      return;
    }
    paragraphs.push({ blockId: source.paraId, start: offset + 1, text, node, source });
  });
  if (failure !== undefined) return Result.err(failure);
  if (paragraphs.length !== document.package.document.content.length) {
    return refuse("The projection changed paragraph structure.");
  }
  return Result.ok(new CanonicalProjection(converted.value, paragraphs));
};

/** Input intent, never projection-step adjacency, determines the journal partition. */
export type CanonicalInputSemantic =
  | "typing"
  | "deleteBackward"
  | "deleteForward"
  | "paste"
  | "composition"
  | "replacement"
  | "structure";

const UNDO_GROUP_WINDOW_MS = 500;

const samePosition = (left: TextPosition, right: TextPosition): boolean =>
  left.story === right.story &&
  left.blockId === right.blockId &&
  left.offset === right.offset &&
  (left.zeroWidthBefore ?? 0) === (right.zeroWidthBefore ?? 0);

const sameSelection = (left: CanonicalSelection, right: CanonicalSelection): boolean =>
  samePosition(left.anchor, right.anchor) && samePosition(left.head, right.head);

type AppliedJournalEntry = {
  type: "applied";
  ops: readonly DocumentOp[];
  inverse: readonly DocumentOp[];
  preSelection: CanonicalSelection;
  postSelection: CanonicalSelection;
  version: number;
  origin: "input";
  semantic: CanonicalInputSemantic;
  grouping: "run" | "isolated";
  time: number;
  boundary: number;
};

type AppliedJournalGroup = {
  entries: AppliedJournalEntry[];
  undo: { type: "entries" } | { type: "replayed"; inverse: readonly DocumentOp[] };
  preSelection: CanonicalSelection;
  postSelection: CanonicalSelection;
};

const continuesGroup = (previous: AppliedJournalEntry, next: AppliedJournalEntry): boolean => {
  switch (next.semantic) {
    case "typing":
    case "deleteBackward":
    case "deleteForward":
      return (
        previous.grouping === "run" &&
        next.grouping === "run" &&
        previous.semantic === next.semantic &&
        previous.boundary === next.boundary &&
        next.time >= previous.time &&
        next.time - previous.time <= UNDO_GROUP_WINDOW_MS &&
        sameSelection(previous.postSelection, next.preSelection) &&
        samePosition(next.preSelection.anchor, next.preSelection.head)
      );
    case "paste":
    case "composition":
    case "replacement":
    case "structure":
      return false;
    default: {
      const unreachable: never = next.semantic;
      return panic(`Unknown canonical input semantic ${unreachable}`);
    }
  }
};

type UndoneJournalEntry = {
  type: "undone";
  group: AppliedJournalGroup;
  redoOps: readonly DocumentOp[];
};

export type CanonicalCommit = {
  transaction: Transaction;
  document: Document;
  projection: CanonicalProjection;
  touched: TouchedBlocks;
  version: number;
  origin: CanonicalOrigin;
  /** Publish only after every projection/plugin transaction has been accepted. */
  publish: () => Result<void, CanonicalSessionError>;
};

type ProjectionAuthorization = { session: CanonicalSession; version: number; doc: PMNode };
const authorizedProjections = new WeakMap<Transaction, ProjectionAuthorization>();

/** Metadata is descriptive; transaction identity and session version authorize mutations. */
export const isCanonicalProjectionTransaction = (
  transaction: Transaction,
  session: CanonicalSession,
): boolean => {
  const authorization = authorizedProjections.get(transaction);
  return (
    authorization?.session === session &&
    authorization.version === session.version &&
    transaction.doc.eq(authorization.doc)
  );
};

type PublishCanonicalProjectionOptions = {
  state: EditorState;
  commit: CanonicalCommit;
  session: CanonicalSession;
};

/** The model publishes only after PM and every plugin accept its exact projection. */
export const publishCanonicalProjection = ({
  state,
  commit,
  session,
}: PublishCanonicalProjectionOptions) => {
  if (!isCanonicalProjectionTransaction(commit.transaction, session))
    return refuse("The staged canonical projection is unauthorized.");
  const applied = Result.try({
    try: () => state.applyTransaction(commit.transaction),
    catch: (cause) =>
      new CanonicalSessionError({
        message: `Canonical plugin staging failed: ${cause instanceof Error ? cause.message : String(cause)}`,
      }),
  });
  if (applied.isErr()) return applied;
  const staged = applied.value;
  if (
    !staged.transactions.includes(commit.transaction) ||
    !staged.state.doc.eq(commit.projection.doc) ||
    staged.transactions.some(
      (transaction) =>
        transaction.docChanged && !isCanonicalProjectionTransaction(transaction, session),
    )
  )
    return refuse("A plugin refused or changed the canonical projection.");
  const published = commit.publish();
  if (published.isErr()) return published;
  return Result.ok(staged);
};

type StageOptions = {
  state: EditorState;
  ops: readonly DocumentOp[];
  selection: CanonicalSelection;
  origin: CanonicalOrigin;
  onPublish: (inverse: readonly DocumentOp[], version: number) => void;
};

type CanonicalSessionSeedOptions = {
  document: Document;
  projection: CanonicalProjection;
  styles: StyleDefinitions | null | undefined;
};

type CanonicalReplaceTextInput = {
  from: number;
  to: number;
  text: string;
  semantic?: CanonicalInputSemantic;
  time?: number;
};

/** Immutable model authority with a journal staged independently of the PM view. */
class CanonicalSession {
  private currentDocument: Document;
  private currentProjection: CanonicalProjection;
  private currentVersion = 0;
  private readonly applied: AppliedJournalGroup[] = [];
  private groupingBoundary = 0;
  private lifecycle: { type: "committed" } | { type: "composing" } = { type: "committed" };
  private readonly undone: UndoneJournalEntry[] = [];
  private readonly styles: StyleDefinitions | null | undefined;

  constructor({ document, projection, styles }: CanonicalSessionSeedOptions) {
    this.currentDocument = document;
    this.currentProjection = projection;
    this.styles = styles == null ? styles : structuredClone(styles);
  }

  get document(): Document {
    if (this.isComposing)
      throw new CanonicalSessionError({
        message: "Composition must finish before taking a snapshot.",
      });
    return this.currentDocument;
  }
  get projection(): CanonicalProjection {
    return this.currentProjection;
  }
  get version(): number {
    return this.currentVersion;
  }
  get canUndo(): boolean {
    return this.applied.length > 0;
  }
  get canRedo(): boolean {
    return this.undone.length > 0;
  }

  get isComposing(): boolean {
    return this.lifecycle.type === "composing";
  }

  breakUndoGroup(): void {
    this.groupingBoundary += 1;
  }

  beginComposition(): Result<void, CanonicalSessionError> {
    if (this.isComposing) return refuse("A canonical composition is already pending.");
    this.lifecycle = { type: "composing" };
    this.breakUndoGroup();
    return Result.ok(undefined);
  }

  endComposition(): void {
    this.lifecycle = { type: "committed" };
    this.breakUndoGroup();
  }

  private checkState(state: EditorState): Result<void, CanonicalSessionError> {
    if (this.isComposing) return refuse("Composition must finish before committing another edit.");
    if (!state.doc.eq(this.projection.doc)) {
      return refuse("The input uses a stale canonical projection.");
    }
    return Result.ok(undefined);
  }

  prepareReplace(
    state: EditorState,
    { from, to, text, semantic = "replacement", time = Date.now() }: CanonicalReplaceTextInput,
  ): Result<CanonicalCommit, CanonicalSessionError> {
    const checked = this.checkState(state);
    if (checked.isErr()) return checked;
    if (!Number.isFinite(time)) return refuse("The input time is invalid.");
    if (state.storedMarks !== null && state.storedMarks.length > 0) {
      return refuse("Stored formatting is not supported in the canonical session.");
    }
    if (from > to || hasIllegalXmlCharacters(text) || /[\t\r\n]/u.test(text)) {
      return refuse("Canonical input accepts a well-formed same-paragraph text replacement.");
    }
    if (from === to && text.length === 0) return refuse("The input makes no text change.");
    const start = this.projection.addressAt(from);
    if (start.isErr()) return start;
    const end = this.projection.addressAt(to);
    if (end.isErr()) return end;
    if (start.value.blockId !== end.value.blockId) {
      return refuse("Cross-paragraph replacement is not supported in the canonical session.");
    }
    const preSelection = this.projection.selectionAt(state);
    if (preSelection.isErr()) return preSelection;
    const ops: DocumentOp[] = [];
    if (from !== to) {
      ops.push({ type: DOCUMENT_OP_TYPES.DELETE_RANGE, from: start.value, to: end.value });
    }
    if (text.length > 0) {
      const paragraph = this.projection.paragraph(start.value.blockId);
      if (paragraph === undefined) panic("Canonical input lost its addressed paragraph.");
      const runProps = authoredInputFormatting({
        paragraph: paragraph.source,
        offset: start.value.offset,
        affinity: from === to ? "before" : "after",
      });
      ops.push({ type: DOCUMENT_OP_TYPES.INSERT_TEXT, at: start.value, text, runProps });
    }
    const caret = { ...start.value, offset: start.value.offset + text.length };
    const postSelection = { anchor: caret, head: caret };
    const isCaret = state.selection.empty;
    const run =
      isCaret &&
      ((semantic === "typing" && from === to && from === state.selection.head) ||
        (semantic === "deleteBackward" && text.length === 0 && to === state.selection.head) ||
        (semantic === "deleteForward" && text.length === 0 && from === state.selection.head));
    return this.stage({
      state,
      ops,
      selection: postSelection,
      origin: "input",
      onPublish: (inverse, version) => {
        const entry = {
          type: "applied",
          ops,
          inverse,
          preSelection: preSelection.value,
          postSelection,
          version,
          origin: "input",
          semantic,
          grouping: run ? "run" : "isolated",
          time,
          boundary: this.groupingBoundary,
        } as const satisfies AppliedJournalEntry;
        const group = this.applied.at(-1);
        const previous = group?.entries.at(-1);
        if (group !== undefined && previous !== undefined && continuesGroup(previous, entry)) {
          group.entries.push(entry);
          group.postSelection = postSelection;
        } else {
          this.applied.push({
            entries: [entry],
            undo: { type: "entries" },
            preSelection: preSelection.value,
            postSelection,
          });
        }
        this.undone.length = 0;
      },
    });
  }

  prepareUndo(state: EditorState): Result<CanonicalCommit, CanonicalSessionError> {
    const group = this.applied.at(-1);
    if (group === undefined) return refuse("There is no canonical edit to undo.");
    return this.stage({
      state,
      ops:
        group.undo.type === "replayed"
          ? group.undo.inverse
          : group.entries.toReversed().flatMap((entry) => entry.inverse),
      selection: group.preSelection,
      origin: "undo",
      onPublish: (redoOps) => {
        this.applied.pop();
        this.undone.push({ type: "undone", group, redoOps });
        this.breakUndoGroup();
      },
    });
  }

  prepareRedo(state: EditorState): Result<CanonicalCommit, CanonicalSessionError> {
    const undone = this.undone.at(-1);
    if (undone === undefined) return refuse("There is no canonical edit to redo.");
    return this.stage({
      state,
      ops: undone.redoOps,
      selection: undone.group.postSelection,
      origin: "redo",
      onPublish: (inverse) => {
        this.undone.pop();
        this.applied.push({
          entries: undone.group.entries,
          undo: { type: "replayed", inverse },
          preSelection: undone.group.preSelection,
          postSelection: undone.group.postSelection,
        });
        this.breakUndoGroup();
      },
    });
  }

  private stage({
    state,
    ops,
    selection,
    origin,
    onPublish,
  }: StageOptions): Result<CanonicalCommit, CanonicalSessionError> {
    const checked = this.checkState(state);
    if (checked.isErr()) return checked;
    const applied = applyDocumentOps(this.currentDocument, ops);
    if (applied.isErr()) return refuse(applied.error.message);
    preservePropertySources(applied.value.document, this.currentDocument);
    const projected = project(applied.value.document, this.styles);
    if (projected.isErr()) return projected;
    const anchor = projected.value.positionAt(selection.anchor);
    if (anchor.isErr()) return anchor;
    const head = projected.value.positionAt(selection.head);
    if (head.isErr()) return head;
    const transaction = state.tr;
    const changed = applied.value.touched.modified
      .map((blockId) => {
        const paragraph = this.projection.paragraph(blockId);
        if (paragraph === undefined)
          panic("Canonical input reported an unknown touched paragraph.");
        return paragraph;
      })
      .sort((left, right) => right.start - left.start);
    for (const oldParagraph of changed) {
      const nextParagraph = projected.value.paragraph(oldParagraph.blockId);
      if (nextParagraph === undefined) return refuse("The operation changed paragraph structure.");
      transaction.replaceWith(
        oldParagraph.start - 1,
        oldParagraph.start - 1 + oldParagraph.node.nodeSize,
        nextParagraph.node,
      );
    }
    if (!transaction.doc.eq(projected.value.doc))
      return refuse("The operation changed unsupported projection content.");
    transaction.setSelection(TextSelection.create(transaction.doc, anchor.value, head.value));
    transaction.setMeta(CANONICAL_PROJECTION_META, {
      type: "canonical",
      origin,
      version: this.version + 1,
    });
    transaction.setMeta("addToHistory", false);
    const baseVersion = this.version;
    const baseBoundary = this.groupingBoundary;
    authorizedProjections.set(transaction, {
      session: this,
      version: baseVersion,
      doc: projected.value.doc,
    });
    return Result.ok({
      transaction,
      document: applied.value.document,
      projection: projected.value,
      touched: applied.value.touched,
      version: baseVersion + 1,
      origin,
      publish: () => {
        if (this.isComposing || this.groupingBoundary !== baseBoundary)
          return refuse("The staged canonical commit crossed a composition or selection boundary.");
        if (this.version !== baseVersion) return refuse("The staged canonical commit is stale.");
        if (!isCanonicalProjectionTransaction(transaction, this)) {
          return refuse("The staged canonical projection was changed after preparation.");
        }
        this.currentDocument = applied.value.document;
        this.currentProjection = projected.value;
        this.currentVersion = baseVersion + 1;
        onPublish(applied.value.inverse, this.currentVersion);
        return Result.ok(undefined);
      },
    });
  }
}

export type { CanonicalProjection, CanonicalSession };

export const createCanonicalSession = (
  document: Document,
  styles?: StyleDefinitions | null,
): Result<CanonicalSession, CanonicalSessionError> => {
  if (!supportsSeed(document)) {
    return refuse(
      "Canonical sessions currently require main-story plain paragraphs without secondary stories or revisions.",
    );
  }
  const owned = cloneDocumentWithParagraphPropertySources(document);
  const normalized = normalizeForOps(owned);
  preservePropertySources(normalized, owned);
  const validated = validateOpsDocument(normalized);
  if (validated.isErr()) return refuse(validated.error.message);
  const projected = project(normalized, styles);
  if (projected.isErr()) return projected;
  return Result.ok(
    new CanonicalSession({ document: normalized, projection: projected.value, styles }),
  );
};

/** Native character deletion stays in one paragraph and consumes a whole grapheme. */
export const deletionRange = (
  state: EditorState,
  direction: "backward" | "forward",
): Result<{ from: number; to: number }, CanonicalSessionError> => {
  const { selection } = state;
  if (!(selection instanceof TextSelection) || !selection.$from.sameParent(selection.$to)) {
    return refuse("Canonical deletion requires a same-paragraph text selection.");
  }
  if (!selection.empty) return Result.ok({ from: selection.from, to: selection.to });
  const text = selection.$from.parent.textContent;
  const offset = selection.$from.parentOffset;
  if (splitsSurrogatePair(text, offset))
    return refuse("The deletion would split a surrogate pair.");
  if (splitsGraphemeCluster(text, offset)) {
    return refuse("Character deletion requires a caret at a grapheme boundary.");
  }
  if (
    (direction === "backward" && offset === 0) ||
    (direction === "forward" && offset === text.length)
  ) {
    return refuse("Paragraph joins are not supported in the canonical session.");
  }
  let boundary = offset + (direction === "backward" ? -1 : 1);
  while (splitsGraphemeCluster(text, boundary)) boundary += direction === "backward" ? -1 : 1;
  const position = selection.$from.start() + boundary;
  return Result.ok(
    direction === "backward"
      ? { from: position, to: selection.to }
      : { from: selection.from, to: position },
  );
};
