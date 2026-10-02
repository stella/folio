import { panic, Result, TaggedError } from "better-result";
import type { Node as PMNode } from "prosemirror-model";
import { TextSelection, type EditorState, type Transaction } from "prosemirror-state";
import {
  applyDocumentOps,
  DOCUMENT_OP_TYPES,
  normalizeForOps,
  OP_STORIES,
  validateOpsDocument,
  compileEditorIntent,
  createEditorIntentIdAllocator,
  paragraphLogicalText,
  editorParagraphGroups,
  type EditorIntent,
  type EditorIntentMode,
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
import type { Document, Paragraph, StyleDefinitions } from "../types/document";

export class CanonicalSessionError extends TaggedError("CanonicalSessionError")<{
  message: string;
}> {}

const refuse = (message: string) => Result.err(new CanonicalSessionError({ message }));

export type CanonicalSelection = { anchor: TextPosition; head: TextPosition };
export type CanonicalOrigin = "input" | "undo" | "redo";
export type CanonicalSessionMode = { type: "editing" } | { type: "suggesting"; author: string };
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

  inputAddressAt(position: number): Result<TextPosition, CanonicalSessionError> {
    const address = this.addressAt(position);
    if (address.isErr()) return address;
    const paragraph = this.paragraph(address.value.blockId);
    if (paragraph === undefined) panic("Input lost its paragraph.");
    let offset = address.value.offset;
    paragraph.node.forEach((node, start) => {
      if (
        node.marks.some((mark) => mark.type.name === "deletion") &&
        offset > start &&
        offset < start + node.nodeSize
      )
        offset = start + node.nodeSize;
    });
    return Result.ok({ ...address.value, offset });
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

/** Source captures follow paragraph identities through structural edits and history. */
type PreservePropertySourcesOptions = { target: Document; source: Document };
const preservePropertySources = ({ target, source }: PreservePropertySourcesOptions): void => {
  const sources = new Map(
    source.package.document.content.flatMap((block) =>
      block.type === "paragraph" ? [[block.paraId, block] as const] : [],
    ),
  );
  for (const derived of target.package.document.content) {
    if (derived.type !== "paragraph") continue;
    const original = sources.get(derived.paraId);
    if (original && derived !== original) copyParagraphPropertySource(derived, original);
  }
  copyDocumentParagraphPropertySourceContract(target, source);
};

const supportsTextContent = (content: Paragraph["content"]): boolean =>
  content.every((item) => {
    switch (item.type) {
      case "run":
        return item.content.every(
          (child) =>
            child.type === "text" &&
            !hasIllegalXmlCharacters(child.text) &&
            !/[\t\r\n]/u.test(child.text),
        );
      case "insertion":
      case "deletion":
        return supportsTextContent(item.content);
      default:
        return false;
    }
  });

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
      paragraph.reviewCarrier === undefined &&
      supportsTextContent(paragraph.content),
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
    const text = paragraphLogicalText(source);
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

type JournalInputOptions = Pick<StageOptions, "state" | "ops"> & {
  preSelection: CanonicalSelection;
  postSelection: CanonicalSelection;
  semantic?: CanonicalInputSemantic;
  time?: number;
  grouping?: "run" | "isolated";
};

type IntentInputOptions = Pick<JournalInputOptions, "semantic" | "time" | "grouping"> & {
  intent:
    | Exclude<EditorIntent, { type: "splitParagraph" }>
    | Omit<Extract<EditorIntent, { type: "splitParagraph" }>, "newBlockId">;
};

/** Immutable model authority with a journal staged independently of the PM view. */
class CanonicalSession {
  private currentDocument: Document;
  private currentProjection: CanonicalProjection;
  private currentVersion = 0;
  private readonly allocateIntentIds = createEditorIntentIdAllocator();
  private mode: CanonicalSessionMode = { type: "editing" };
  private readonly sourceOwners = new Map<string, Paragraph>();
  private readonly applied: AppliedJournalGroup[] = [];
  private groupingBoundary = 0;
  private lifecycle: { type: "committed" } | { type: "composing" } = { type: "committed" };
  private readonly undone: UndoneJournalEntry[] = [];
  private readonly styles: StyleDefinitions | null | undefined;

  constructor({ document, projection, styles }: CanonicalSessionSeedOptions) {
    this.currentDocument = document;
    this.currentProjection = projection;
    this.styles = styles == null ? styles : structuredClone(styles);
    this.rememberSources(document);
  }

  setMode(mode: CanonicalSessionMode): void {
    if (
      this.mode.type !== mode.type ||
      (this.mode.type === "suggesting" &&
        mode.type === "suggesting" &&
        this.mode.author !== mode.author)
    )
      this.breakUndoGroup();
    this.mode =
      mode.type === "editing" ? { type: "editing" } : { type: "suggesting", author: mode.author };
  }

  private rememberSources(document: Document): void {
    for (const block of document.package.document.content) {
      if (block.type === "paragraph" && block.paraId !== undefined)
        this.sourceOwners.set(block.paraId, block);
    }
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
      return refuse("Canonical input accepts a well-formed text replacement.");
    }
    if (from === to && text.length === 0) return refuse("The input makes no text change.");
    const start = this.projection.inputAddressAt(from);
    if (start.isErr()) return start;
    const end = this.projection.inputAddressAt(to);
    if (end.isErr()) return end;
    const run =
      state.selection.empty &&
      ((semantic === "typing" && from === to && from === state.selection.head) ||
        (semantic === "deleteBackward" && text.length === 0 && to === state.selection.head) ||
        (semantic === "deleteForward" && text.length === 0 && from === state.selection.head));
    return this.prepareIntent(state, {
      intent: {
        type: "replaceText",
        from: start.value,
        to: end.value,
        text,
      },
      semantic,
      time,
      grouping: run ? "run" : "isolated",
    });
  }

  prepareSplit(state: EditorState): Result<CanonicalCommit, CanonicalSessionError> {
    const checked = this.checkState(state);
    if (checked.isErr()) return checked;
    const at = this.projection.inputAddressAt(state.selection.from);
    if (at.isErr()) return at;
    const to = this.projection.inputAddressAt(state.selection.to);
    if (to.isErr()) return to;
    return this.prepareIntent(state, {
      intent: {
        type: "splitParagraph",
        at: at.value,
        to: to.value,
      },
    });
  }

  prepareJoin(
    state: EditorState,
    direction: "backward" | "forward",
  ): Result<CanonicalCommit, CanonicalSessionError> {
    const checked = this.checkState(state);
    if (checked.isErr()) return checked;
    if (!state.selection.empty) return refuse("Join requires a collapsed text selection.");
    const at = this.projection.inputAddressAt(state.selection.from);
    if (at.isErr()) return at;
    const groups = editorParagraphGroups(this.document, at.value.story);
    const index = groups.findIndex(({ paragraphs }) =>
      paragraphs.some((paragraph) => paragraph.paraId === at.value.blockId),
    );
    if (index < 0 || (direction === "backward" && index === 0))
      return refuse("There is no adjacent paragraph to join.");
    const firstGroup = groups.at(direction === "backward" ? index - 1 : index);
    const secondGroup = groups.at(direction === "backward" ? index : index + 1);
    const first = firstGroup?.paragraphs.at(-1);
    const second = secondGroup?.paragraphs.at(0);
    if (!first?.paraId || !second?.paraId) return refuse("There is no adjacent paragraph to join.");
    if (!isCanonicalJoinBoundary(state, direction))
      return refuse("Join requires a caret at a paragraph boundary.");
    return this.prepareIntent(state, {
      intent: {
        type: "joinParagraphs",
        story: at.value.story,
        blockId: first.paraId,
        nextBlockId: second.paraId,
      },
    });
  }

  private prepareIntent(
    state: EditorState,
    { intent, ...journal }: IntentInputOptions,
  ): Result<CanonicalCommit, CanonicalSessionError> {
    const checked = this.checkState(state);
    if (checked.isErr()) return checked;
    const preSelection = this.projection.selectionAt(state);
    if (preSelection.isErr()) return preSelection;
    const ids = this.allocateIntentIds(this.document, intent);
    const mode =
      this.mode.type === "editing"
        ? ({ type: "editing", newIds: ids.newIds } as const satisfies EditorIntentMode)
        : ({
            type: "suggesting",
            revision: {
              id: ids.revisionId,
              author: this.mode.author,
              date: new Date().toISOString(),
            },
            newIds: ids.newIds,
          } as const satisfies EditorIntentMode);
    const compiled = compileEditorIntent(this.document, {
      intent: intent.type === "splitParagraph" ? { ...intent, newBlockId: ids.newBlockId } : intent,
      mode,
    });
    if (compiled.isErr()) return refuse(compiled.error.message);
    return this.prepareJournalled({
      state,
      ops: compiled.value.ops,
      ...journal,
      preSelection: preSelection.value,
      postSelection: {
        anchor: compiled.value.selection,
        head: compiled.value.selection,
      },
    });
  }

  prepareResolve(
    state: EditorState,
    {
      revisionIds,
      resolution,
    }: { revisionIds: readonly number[]; resolution: "accept" | "reject" },
  ): Result<CanonicalCommit, CanonicalSessionError> {
    const checked = this.checkState(state);
    if (checked.isErr()) return checked;
    const preSelection = this.projection.selectionAt(state);
    if (preSelection.isErr()) return preSelection;
    const ops = [
      {
        type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
        story: OP_STORIES.MAIN,
        revisionIds,
        decision: resolution,
      },
    ] as const;
    // Resolution can retire the selected paragraph; choose a valid result anchor before staging.
    const applied = applyDocumentOps(this.document, ops);
    if (applied.isErr()) return refuse(applied.error.message);
    const projected = project(applied.value.document, this.styles);
    if (projected.isErr()) return projected;
    const surviving = projected.value.paragraph(preSelection.value.head.blockId);
    const first = applied.value.document.package.document.content.at(0);
    if (first?.type !== "paragraph" || first.paraId === undefined)
      return refuse("Resolution left no paragraph.");
    let offset = surviving ? Math.min(preSelection.value.head.offset, surviving.text.length) : 0;
    if (surviving && splitsSurrogatePair(surviving.text, offset)) offset -= 1;
    const caret = { story: OP_STORIES.MAIN, blockId: surviving?.blockId ?? first.paraId, offset };
    return this.prepareJournalled({
      state,
      ops,
      preSelection: preSelection.value,
      postSelection: { anchor: caret, head: caret },
    });
  }

  private prepareJournalled({
    state,
    ops,
    preSelection,
    postSelection,
    semantic = "structure",
    time = Date.now(),
    grouping = "isolated",
  }: JournalInputOptions): Result<CanonicalCommit, CanonicalSessionError> {
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
          preSelection,
          postSelection,
          version,
          origin: "input",
          semantic,
          grouping,
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
            preSelection,
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
    preservePropertySources({ target: applied.value.document, source: this.document });
    const stagedSources = new Map(this.sourceOwners);
    for (const op of ops) {
      if (op.type !== DOCUMENT_OP_TYPES.SPLIT_BLOCK) continue;
      const source = stagedSources.get(op.at.blockId);
      if (source) stagedSources.set(op.newBlockId, source);
    }
    for (const block of applied.value.document.package.document.content) {
      if (block.type !== "paragraph" || block.paraId === undefined) continue;
      const source = stagedSources.get(block.paraId);
      if (source && block !== source) copyParagraphPropertySource(block, source);
    }
    const projected = project(applied.value.document, this.styles);
    if (projected.isErr()) return projected;
    const anchor = projected.value.positionAt(selection.anchor);
    if (anchor.isErr()) return anchor;
    const head = projected.value.positionAt(selection.head);
    if (head.isErr()) return head;
    const transaction = state.tr;
    const structureChanged =
      applied.value.touched.inserted.length > 0 || applied.value.touched.removed.length > 0;
    if (structureChanged) {
      transaction.replaceWith(0, transaction.doc.content.size, projected.value.doc.content);
    } else {
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
        if (nextParagraph === undefined)
          return refuse("The operation changed paragraph structure.");
        transaction.replaceWith(
          oldParagraph.start - 1,
          oldParagraph.start - 1 + oldParagraph.node.nodeSize,
          nextParagraph.node,
        );
      }
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
        this.rememberSources(applied.value.document);
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
      "Canonical sessions currently require main-story text paragraphs without secondary stories or inline atoms.",
    );
  }
  const owned = cloneDocumentWithParagraphPropertySources(document);
  const normalized = normalizeForOps(owned);
  preservePropertySources({ target: normalized, source: owned });
  const validated = validateOpsDocument(normalized);
  if (validated.isErr()) return refuse(validated.error.message);
  const projected = project(normalized, styles);
  if (projected.isErr()) return projected;
  return Result.ok(
    new CanonicalSession({ document: normalized, projection: projected.value, styles }),
  );
};

const hasDeletedParagraphMark = (node: PMNode): boolean => {
  const mark: unknown = node.attrs["pPrMark"];
  return (
    typeof mark === "object" &&
    mark !== null &&
    "kind" in mark &&
    (mark.kind === "del" || mark.kind === "moveFrom")
  );
};

/** Visible graphemes concatenate across retained deletions and deleted paragraph marks. */
const visibleDeletionContext = (state: EditorState) => {
  const { selection } = state;
  const index = selection.$from.index(0);
  let first = index;
  while (first > 0 && hasDeletedParagraphMark(state.doc.child(first - 1))) first -= 1;
  let last = index;
  while (last + 1 < state.doc.childCount && hasDeletedParagraphMark(state.doc.child(last)))
    last += 1;
  let text = "";
  const physicalGaps: number[] = [];
  let visibleOffset = 0;
  state.doc.forEach((paragraph, paragraphOffset, paragraphIndex) => {
    if (paragraphIndex < first || paragraphIndex > last) return;
    paragraph.forEach((node, offset) => {
      if (node.marks.some((mark) => mark.type.name === "deletion")) return;
      if (!node.isText) panic("Canonical deletion encountered an inline atom.");
      const value = node.text ?? "";
      for (let unit = 0; unit < value.length; unit += 1) {
        const position = paragraphOffset + 1 + offset + unit;
        physicalGaps[text.length] = position;
        text += value.charAt(unit);
        physicalGaps.push(position + 1);
        if (position < selection.from) visibleOffset += 1;
      }
    });
  });
  return { text, physicalGaps, visibleOffset };
};

export const isCanonicalJoinBoundary = (
  state: EditorState,
  direction: "backward" | "forward",
): boolean => {
  if (!(state.selection instanceof TextSelection) || !state.selection.empty) return false;
  const { text, visibleOffset } = visibleDeletionContext(state);
  return visibleOffset === (direction === "backward" ? 0 : text.length);
};

/** Native character deletion consumes one whole visible grapheme. */
export const deletionRange = (
  state: EditorState,
  direction: "backward" | "forward",
): Result<{ from: number; to: number }, CanonicalSessionError> => {
  const { selection } = state;
  if (!(selection instanceof TextSelection))
    return refuse("Canonical deletion requires a text selection.");
  if (!selection.empty) return Result.ok({ from: selection.from, to: selection.to });
  const { text, physicalGaps, visibleOffset } = visibleDeletionContext(state);
  if (splitsSurrogatePair(text, visibleOffset))
    return refuse("The deletion would split a surrogate pair.");
  if (splitsGraphemeCluster(text, visibleOffset))
    return refuse("Character deletion requires a caret at a grapheme boundary.");
  if (
    (direction === "backward" && visibleOffset === 0) ||
    (direction === "forward" && visibleOffset === text.length)
  )
    return refuse("There is no visible character to delete in this paragraph.");
  let boundary = visibleOffset + (direction === "backward" ? -1 : 1);
  while (splitsGraphemeCluster(text, boundary)) boundary += direction === "backward" ? -1 : 1;
  const position = physicalGaps.at(boundary);
  if (position === undefined) panic("Canonical deletion lost a grapheme boundary.");
  return Result.ok(
    direction === "backward"
      ? { from: position, to: selection.to }
      : { from: selection.from, to: position },
  );
};
