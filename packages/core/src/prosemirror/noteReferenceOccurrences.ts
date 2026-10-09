import { Result, TaggedError, panic } from "better-result";
import { Fragment, Mark, Slice, type Node as PMNode } from "prosemirror-model";
import { Plugin, type Transaction } from "prosemirror-state";
import { isHistoryTransaction } from "prosemirror-history";
import { readFootnoteRefMarkAttrs } from "./attrs";

export class NoteReferenceEditRefusal extends TaggedError("NoteReferenceEditRefusal")<{
  message: string;
}> {}

export class NoteReferenceReplayDefect extends TaggedError("NoteReferenceReplayDefect")<{
  message: string;
}> {}

const NOTE_OCCURRENCE_REPLAY_META = "noteReferenceOccurrenceReplay";
export const markNoteReferenceReplay = (transaction: Transaction): void => {
  transaction.setMeta(NOTE_OCCURRENCE_REPLAY_META, true);
};

const NOTE_OCCURRENCE_REPORT_META = "noteReferenceOccurrenceReport";
export const noteReferenceTransactionIssue = (
  transaction: Transaction,
): NoteReferenceEditRefusal | NoteReferenceReplayDefect | undefined => {
  const issue: unknown = transaction.getMeta(NOTE_OCCURRENCE_REPORT_META);
  return issue instanceof NoteReferenceEditRefusal || issue instanceof NoteReferenceReplayDefect
    ? issue
    : undefined;
};

export const mintNoteReferenceOccurrenceId = (): string => crypto.randomUUID();

type Occurrence = { id: string; node: PMNode; text: string; from: number; to: number };

/** Identity establishes boundaries; text only validates the already bounded unit. */
const readOccurrences = (paragraph: PMNode) => {
  const referenceType = paragraph.type.schema.marks["footnoteRef"];
  if (!referenceType || !paragraph.rangeHasMark(0, paragraph.content.size, referenceType))
    return Result.ok([]);
  const occurrences: Occurrence[] = [];
  const seen = new Set<string>();
  let previous: Occurrence | undefined;
  let issue: string | undefined;
  paragraph.forEach((node, from) => {
    if (issue !== undefined) return;
    const reference = node.marks.find((mark) => mark.type.name === "footnoteRef");
    if (!reference) {
      previous = undefined;
      return;
    }
    const attrs = readFootnoteRefMarkAttrs(reference);
    if (!attrs.ok || attrs.value.occurrenceId.length === 0 || !node.isText) {
      issue = "Note references require attributed text occurrences.";
      return;
    }
    const id = attrs.value.occurrenceId;
    if (previous?.id === id) {
      if (!Mark.sameSet(previous.node.marks, node.marks)) {
        issue =
          "A note reference cannot have different formatting or revision owners within one occurrence.";
        return;
      }
      previous.text += node.text ?? "";
      previous.to = from + node.nodeSize;
      return;
    }
    if (seen.has(id)) {
      issue = "A note-reference occurrence must remain contiguous.";
      return;
    }
    seen.add(id);
    previous = { id, node, text: node.text ?? "", from, to: from + node.nodeSize };
    occurrences.push(previous);
  });
  for (const occurrence of occurrences) {
    const reference = occurrence.node.marks.find((mark) => mark.type.name === "footnoteRef");
    if (!reference) return panic("Missing note-reference mark in an occurrence.");
    const attrs = readFootnoteRefMarkAttrs(reference);
    if (!attrs.ok) return panic("Invalid attributed note-reference occurrence.");
    if (occurrence.text !== String(attrs.value.id))
      issue ??= "Edit the whole note reference; its rendered label is indivisible.";
  }
  return issue === undefined
    ? Result.ok(occurrences)
    : Result.err(new NoteReferenceEditRefusal({ message: issue }));
};

/** Saving asserts the edit-owner invariant and emits each occurrence once. */
export const coalesceNoteReferenceOccurrences = (paragraph: PMNode): PMNode => {
  if (saveOccurrenceValidator?.cachedIds(paragraph)?.length === 0) return paragraph;
  const referenceType = paragraph.type.schema.marks["footnoteRef"];
  if (!referenceType || !paragraph.rangeHasMark(0, paragraph.content.size, referenceType))
    return paragraph;
  const result = readOccurrences(paragraph);
  if (result.isErr()) return panic(result.error.message);
  if (result.value.every(({ node, from, to }) => node.nodeSize === to - from)) return paragraph;
  const byStart = new Map(result.value.map((occurrence) => [occurrence.from, occurrence]));
  const nodes: PMNode[] = [];
  let through = -1;
  paragraph.forEach((node, from) => {
    if (from < through) return;
    const occurrence = byStart.get(from);
    if (occurrence) {
      nodes.push(node.type.schema.text(occurrence.text, node.marks));
      through = occurrence.to;
    } else nodes.push(node);
  });
  return paragraph.copy(Fragment.fromArray(nodes));
};

const createOccurrenceValidator = () => {
  // Immutable nodes share their validation result across transactions.
  const cache = new WeakMap<PMNode, Result<readonly string[], NoteReferenceEditRefusal>>();
  const validate = (node: PMNode): Result<readonly string[], NoteReferenceEditRefusal> => {
    const cached = cache.get(node);
    if (cached) return cached;
    const ids = new Set<string>();
    const include = (id: string) => {
      if (ids.has(id))
        return new NoteReferenceEditRefusal({
          message: "Separate note references require distinct occurrence identities.",
        });
      ids.add(id);
      return undefined;
    };
    if (node.inlineContent) {
      const own = readOccurrences(node);
      if (own.isErr()) return Result.err(own.error);
      for (const occurrence of own.value) {
        const issue = include(occurrence.id);
        if (issue) return Result.err(issue);
      }
    }
    for (let index = 0; index < node.childCount; index += 1) {
      const child = node.child(index);
      if (child.isLeaf) continue;
      const nested = validate(child);
      if (nested.isErr()) return nested;
      for (const id of nested.value) {
        const issue = include(id);
        if (issue) return Result.err(issue);
      }
    }
    const result = Result.ok([...ids]);
    cache.set(node, result);
    return result;
  };
  return {
    validate,
    cachedIds: (node: PMNode) => {
      const facts = cache.get(node);
      return facts?.isOk() ? facts.value : undefined;
    },
    validateSave: (doc: PMNode) => {
      const facts = cache.get(doc);
      // An immutable reference-free root has no identities to compose. Changed roots miss this cache.
      if (facts?.isOk() && facts.value.length === 0) return facts;
      // Compose cached subtree facts anew on every referenced save, including cross-subtree uniqueness.
      cache.delete(doc);
      return validate(doc);
    },
  };
};

// Saving reuses immutable subtrees just as the edit guard does. Weak keys do not retain documents.
let saveOccurrenceValidator: ReturnType<typeof createOccurrenceValidator> | undefined;
export const assertNoteReferenceOccurrences = (doc: PMNode): void => {
  const validate = (saveOccurrenceValidator ??= createOccurrenceValidator());
  const result = validate.validateSave(doc);
  if (result.isErr()) panic(result.error.message);
};

/** Production edit boundary: refuse before committing an unsaveable occurrence. */
export const noteReferenceOccurrencePlugin = () => {
  const { validate } = createOccurrenceValidator();
  return new Plugin({
    filterTransaction(transaction) {
      if (!transaction.docChanged) return true;
      const result = validate(transaction.doc);
      if (result.isOk()) return true;
      // The collaboration owner identifies its actual plugin key; never guess it.
      // A broken replicated state is a defect, not a local edit refusal.
      const replay =
        isHistoryTransaction(transaction) ||
        transaction.getMeta(NOTE_OCCURRENCE_REPLAY_META) === true;
      transaction.setMeta(
        NOTE_OCCURRENCE_REPORT_META,
        replay ? new NoteReferenceReplayDefect({ message: result.error.message }) : result.error,
      );
      return replay;
    },
    props: {
      transformPasted(slice) {
        const identities = new Map<string, string>();
        const remint = (fragment: Fragment): Fragment => {
          const nodes: PMNode[] = [];
          fragment.forEach((node) => {
            const marks = node.marks.map((mark) => {
              if (mark.type.name !== "footnoteRef") return mark;
              const attrs = readFootnoteRefMarkAttrs(mark);
              // Keep malformed input intact for the paste or transaction refusal boundary.
              if (!attrs.ok) return mark;
              const id = attrs.value.occurrenceId;
              let replacement = identities.get(id);
              if (replacement === undefined) {
                replacement = mintNoteReferenceOccurrenceId();
                identities.set(id, replacement);
              }
              return mark.type.create({ ...mark.attrs, occurrenceId: replacement });
            });
            nodes.push(node.copy(remint(node.content)).mark(marks));
          });
          return Fragment.fromArray(nodes);
        };
        return new Slice(remint(slice.content), slice.openStart, slice.openEnd);
      },
    },
  });
};
