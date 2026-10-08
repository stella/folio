import { expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import { Result, panic } from "better-result";
import { Fragment, Slice, type Node as PMNode } from "prosemirror-model";
import { EditorState, TextSelection } from "prosemirror-state";
import { inlineLeafSpans } from "@stll/docx-core/ops";
import { assertProperty, propertyTestTimeout } from "../../../../../test/property-testing";
import { createEmptyDocument } from "../../utils/createDocument";
import { createDocx } from "../../docx/rezip";
import {
  createHarnessState,
  HARNESS_AUTHOR,
  HeadlessEditorView,
  parseShapeDocument,
  resolveAllChanges,
  saveHarnessState,
} from "../../__tests__/editorHarness";
import { toggleBold } from "../commands/formatting";
import {
  NoteReferenceEditRefusal,
  noteReferenceTransactionIssue,
} from "../noteReferenceOccurrences";
import { singletonManager } from "../schema";
import { fromProseDoc } from "./fromProseDoc";

setDefaultTimeout(propertyTestTimeout(60_000));

const KINDS = ["footnote", "endnote"] as const;
const OWNERS = ["live", "ownInsertion", "otherDeletion", "ownDeletion"] as const;
const FROM = 2;
const LABEL = "123";
type Kind = (typeof KINDS)[number];

const sourceDocument = async (kind: Kind) => {
  const source = createEmptyDocument({ initialText: "LR" });
  source.package.document.content = [
    {
      type: "paragraph",
      paraId: "12345678",
      content: [
        { type: "run", content: [{ type: "text", text: "L" }] },
        ...Array.from({ length: 3 }, () => ({
          type: "run" as const,
          formatting: { bold: true },
          content: [
            {
              type: kind === "footnote" ? ("footnoteRef" as const) : ("endnoteRef" as const),
              id: 123,
            },
          ],
        })),
        { type: "run", content: [{ type: "text", text: "R" }] },
      ],
    },
  ];
  const content = createEmptyDocument({ initialText: "Note" }).package.document.content;
  if (kind === "footnote") source.package.footnotes = [{ type: "footnote", id: 123, content }];
  else source.package.endnotes = [{ type: "endnote", id: 123, content }];
  return parseShapeDocument(new Uint8Array(await createDocx(source)));
};

const tokens = (doc: PMNode) => {
  const result: unknown[] = [];
  doc.descendants((node) => {
    if (node.type.name === "paragraph") result.push({ type: "paragraph" });
    if (!node.isText) return;
    const ref = node.marks.find((mark) => mark.type.name === "footnoteRef");
    const revision = node.marks.find(
      (mark) => mark.type.name === "insertion" || mark.type.name === "deletion",
    );
    for (const text of node.text ?? "")
      result.push({
        text,
        ref: ref ? { id: ref.attrs.id, kind: ref.attrs.noteType } : null,
        bold: node.marks.some((mark) => mark.type.name === "bold"),
        italic: node.marks.some((mark) => mark.type.name === "italic"),
        revision: revision
          ? {
              type: revision.type.name,
              author: revision.attrs.author,
              date: revision.attrs.date ? new Date(revision.attrs.date).toISOString() : null,
              moveKind: revision.attrs.moveKind,
            }
          : null,
      });
  });
  return result;
};

const occurrences = (doc: PMNode) => {
  const ids = new Set<string>();
  doc.descendants((node) => {
    const reference = node.marks.find((mark) => mark.type.name === "footnoteRef");
    if (reference) ids.add(reference.attrs.occurrenceId);
  });
  return ids;
};

const assertRoundtrip = async (
  state: EditorState,
  base: Awaited<ReturnType<typeof sourceDocument>>,
) => {
  const saved = await saveHarnessState(state, base);
  const parsed = await parseShapeDocument(saved.bytes);
  const reopened = createHarnessState(parsed, "editing");
  expect(reopened.doc.textContent).toBe(state.doc.textContent);
  expect(tokens(reopened.doc)).toEqual(tokens(state.doc));
  expect(occurrences(reopened.doc).size).toBe(occurrences(state.doc).size);
  let sourceCount = 0;
  for (const block of parsed.package.document.content) {
    if (block.type !== "paragraph") continue;
    sourceCount += inlineLeafSpans(block.content).filter(
      ({ node }) => node.type === "footnoteRef" || node.type === "endnoteRef",
    ).length;
  }
  expect(sourceCount).toBe(occurrences(state.doc).size);
};

const withOwnership = (state: EditorState, owners: readonly (typeof OWNERS)[number][]) => {
  const tr = state.tr;
  for (const [index, owner] of owners.entries()) {
    if (owner === "live") continue;
    tr.addMark(
      FROM + index * LABEL.length,
      FROM + (index + 1) * LABEL.length,
      state.schema.mark(owner === "ownInsertion" ? "insertion" : "deletion", {
        revisionId: 200 + index,
        author: owner === "otherDeletion" ? "Other" : HARNESS_AUTHOR,
        date: "2026-01-01T00:00:00Z",
      }),
    );
  }
  return state.apply(tr);
};

test("retracted adjacent note occurrences preserve multiplicity and ownership through save and reopen", async () => {
  const bases = {
    footnote: await sourceDocument("footnote"),
    endnote: await sourceDocument("endnote"),
  };
  const check = async (
    kind: Kind,
    direction: "forward" | "backward",
    owners: readonly (typeof OWNERS)[number][],
  ) => {
    let state = withOwnership(createHarnessState(bases[kind], "suggesting"), owners);
    state = state.apply(
      state.tr.setSelection(
        TextSelection.create(state.doc, FROM + LABEL.length, FROM + 2 * LABEL.length),
      ),
    );
    const view = new HeadlessEditorView(state);
    expect(view.pressKey(direction === "forward" ? "Delete" : "Backspace")).toBe(true);
    const rejected = resolveAllChanges(view.state, "reject");
    if (owners.join(",") === "live,ownInsertion,live")
      expect(rejected.doc.textContent).toBe("L123123R");
    await assertRoundtrip(rejected, bases[kind]);
    await assertRoundtrip(resolveAllChanges(view.state, "accept"), bases[kind]);
  };
  // Retraction joins equal note IDs; the save/reopen oracle must retain both occurrences.
  // Ablating structural identities loses one label (L123123R becomes L123R).
  await check("footnote", "forward", ["live", "ownInsertion", "live"]);
  await assertProperty(
    fc.asyncProperty(
      fc.constantFrom(...KINDS),
      fc.constantFrom("forward", "backward"),
      fc.tuple(fc.constantFrom(...OWNERS), fc.constantFrom(...OWNERS), fc.constantFrom(...OWNERS)),
      check,
    ),
    { numRuns: 20 },
  );
});

test("arbitrary digit splits preserve adjacent occurrences and whole-unit revision owners", async () => {
  const bases = {
    footnote: await sourceDocument("footnote"),
    endnote: await sourceDocument("endnote"),
  };
  await assertProperty(
    fc.asyncProperty(
      fc.constantFrom(...KINDS),
      fc.tuple(
        fc.integer({ min: 0, max: 3 }),
        fc.integer({ min: 0, max: 3 }),
        fc.integer({ min: 0, max: 3 }),
      ),
      fc.tuple(fc.constantFrom(...OWNERS), fc.constantFrom(...OWNERS), fc.constantFrom(...OWNERS)),
      async (kind, splits, owners) => {
        const state = withOwnership(createHarnessState(bases[kind], "suggesting"), owners);
        const paragraph = state.doc.child(0);
        const nodes: PMNode[] = [];
        let index = 0;
        paragraph.forEach((node) => {
          if (!node.marks.some((mark) => mark.type.name === "footnoteRef")) {
            nodes.push(node);
            return;
          }
          const split = splits[index++];
          if (split === undefined) return panic("Missing split shape.");
          let start = 0;
          for (let end = 1; end <= LABEL.length; end += 1) {
            if (end !== LABEL.length && (split & (1 << (end - 1))) === 0) continue;
            nodes.push(state.schema.text(LABEL.slice(start, end), node.marks));
            start = end;
          }
        });
        // Retain legal adjacent text leaves instead of PM's automatic equal-mark merge.
        const splitParagraph = paragraph.copy(new Fragment(nodes, paragraph.content.size));
        const doc = state.doc.copy(Fragment.from(splitParagraph));
        const splitState = EditorState.create({ doc, plugins: state.plugins });
        await assertRoundtrip(splitState, bases[kind]);
        expect(fromProseDoc(doc, bases[kind]).package.document.content).toEqual(
          fromProseDoc(state.doc, bases[kind]).package.document.content,
        );
      },
    ),
    { numRuns: 30 },
  );
});

test("random public edits keep occurrences saveable or visibly refuse before commit", async () => {
  const base = await sourceDocument("footnote");
  const reference = createHarnessState(base, "editing").doc.nodeAt(FROM);
  if (!reference) return panic("Missing copy source occurrence.");
  const referenceSlice = new Slice(Fragment.from(reference), 0, 0);
  const editAction = fc.record({
    kind: fc.constantFrom(
      "bold",
      "delete",
      "type",
      "paste",
      "pasteReference",
      "split",
      "join",
      "undo",
      "redo",
    ),
    left: fc.nat(15),
    right: fc.nat(15),
  });
  await assertProperty(
    fc.asyncProperty(
      fc.constantFrom("editing", "suggesting"),
      fc.array(editAction, { minLength: 1, maxLength: 16 }),
      async (mode, actions) => {
        const view = new HeadlessEditorView(createHarnessState(base, mode));
        for (const action of actions) {
          const positions: number[] = [];
          view.state.doc.descendants((node, position) => {
            if (!node.isTextblock) return true;
            for (let offset = 0; offset <= node.content.size; offset += 1)
              positions.push(position + 1 + offset);
            return false;
          });
          const left = positions.at(action.left % positions.length);
          const right = positions.at(action.right % positions.length);
          if (left === undefined || right === undefined) return panic("Missing edit endpoint.");
          view.state = view.state.apply(
            view.state.tr.setSelection(
              TextSelection.create(view.state.doc, Math.min(left, right), Math.max(left, right)),
            ),
          );
          const before = view.state;
          const transactionCount = view.transactions.length;
          const attempt = Result.try({
            try: () => {
              switch (action.kind) {
                case "bold":
                  return toggleBold(view.state, view.dispatch);
                case "delete":
                  return view.pressKey("Delete");
                case "type":
                  return view.typeText("x");
                case "paste":
                  return view.paste(new Slice(Fragment.from(view.state.schema.text("x")), 0, 0));
                case "pasteReference":
                  return view.paste(referenceSlice);
                case "split":
                  return view.pressKey("Enter");
                case "join":
                  return view.pressKey("Backspace");
                case "undo":
                  return singletonManager.requireCommand("undo")()(view.state, view.dispatch);
                case "redo":
                  return singletonManager.requireCommand("redo")()(view.state, view.dispatch);
                default:
                  return panic(String(action.kind satisfies never));
              }
            },
            catch: (error) => error,
          });
          expect(attempt.isOk()).toBe(true);
          const transaction =
            view.transactions.length > transactionCount ? view.transactions.at(-1) : undefined;
          const refusal = transaction && noteReferenceTransactionIssue(transaction);
          if (refusal) {
            expect(refusal).toBeInstanceOf(NoteReferenceEditRefusal);
            expect(view.state).toBe(before);
          }
          await assertRoundtrip(view.state, base);
        }
      },
    ),
    { numRuns: 15 },
  );
});

test("unattributed or mixed-owner serializer inputs panic instead of guessing", async () => {
  const base = await sourceDocument("footnote");
  const state = createHarnessState(base, "editing");
  const mixed = state.tr.addMark(FROM, FROM + 1, state.schema.mark("bold"));
  // The source already has bold; use an actual conflicting partial change.
  mixed.addMark(FROM, FROM + 1, state.schema.mark("italic"));
  expect(state.apply(mixed)).toBe(state);
  expect(noteReferenceTransactionIssue(mixed)).toBeInstanceOf(NoteReferenceEditRefusal);
  expect(() => fromProseDoc(mixed.doc)).toThrow();
  const mark = state.doc
    .nodeAt(FROM)
    ?.marks.find((candidate) => candidate.type.name === "footnoteRef");
  if (!mark) return panic("Missing parsed reference.");
  const unowned = state.tr.addMark(
    FROM,
    FROM + LABEL.length,
    mark.type.create({ ...mark.attrs, occurrenceId: "" }),
  );
  expect(state.apply(unowned)).toBe(state);
  expect(noteReferenceTransactionIssue(unowned)).toBeInstanceOf(NoteReferenceEditRefusal);
  expect(() => fromProseDoc(unowned.doc)).toThrow();
});

test("DOM and clipboard preserve unit attribution while pasted occurrences get fresh identities", async () => {
  const base = await sourceDocument("footnote");
  const state = createHarnessState(base, "editing");
  const reference = state.doc.nodeAt(FROM)?.marks.find((mark) => mark.type.name === "footnoteRef");
  if (!reference) return panic("Missing parsed reference.");
  const slice = new Slice(Fragment.from(state.doc.nodeAt(FROM)), 0, 0);
  const view = new HeadlessEditorView(
    state.apply(state.tr.setSelection(TextSelection.create(state.doc, FROM))),
  );
  view.paste(slice);
  const nextMark = view.state.doc
    .nodeAt(FROM)
    ?.marks.find((mark) => mark.type.name === "footnoteRef");
  expect(nextMark?.attrs.occurrenceId).not.toBe(reference.attrs.occurrenceId);
  expect(occurrences(view.state.doc).size).toBe(4);
  await assertRoundtrip(view.state, base);
});
