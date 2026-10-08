import { expect, test, setDefaultTimeout } from "bun:test";
import fc from "fast-check";
import { TextSelection } from "prosemirror-state";
import { panic } from "better-result";
import { assertProperty, propertyTestTimeout } from "../../../../test/property-testing";
import { createDocx } from "../docx/rezip";
import { createEmptyDocument } from "../utils/createDocument";
import {
  createHarnessState,
  HARNESS_AUTHOR,
  HeadlessEditorView,
  parseShapeDocument,
  resolveAllChanges,
  saveHarnessState,
} from "../__tests__/editorHarness";
import { deletionRange } from "./canonicalSession";
import { noteReferenceStates } from "../prosemirror/noteReferenceReview";

setDefaultTimeout(propertyTestTimeout(60_000));

const OWNERSHIPS = ["live", "ownInsertion", "otherDeletion", "ownDeletion"] as const;
type Ownership = (typeof OWNERSHIPS)[number];
const NOTE_TYPES = ["footnote", "endnote"] as const;
type NoteType = (typeof NOTE_TYPES)[number];
const REFERENCE_ID = 123;
const LABEL = String(REFERENCE_ID);
const REFERENCE_FROM = 2;
const FRAGMENT_COUNT = 3;
const REFERENCE_TO = REFERENCE_FROM + LABEL.length * FRAGMENT_COUNT;

// Serialize and parse the same model shape the public note insertion command produces.
const sourceDocument = async (kind: NoteType) => {
  const source = createEmptyDocument({ initialText: "LR" });
  source.package.document.content = [
    {
      type: "paragraph",
      paraId: "12345678",
      content: [
        { type: "run", content: [{ type: "text", text: "L" }] },
        ...Array.from({ length: FRAGMENT_COUNT }, (_, index) => ({
          type: "run" as const,
          formatting: index % 2 === 0 ? { bold: true } : { italic: true },
          content: [
            {
              type: kind === "footnote" ? ("footnoteRef" as const) : ("endnoteRef" as const),
              id: REFERENCE_ID,
            },
          ],
        })),
        { type: "run", content: [{ type: "text", text: "R" }] },
      ],
    },
  ];
  const content = createEmptyDocument({ initialText: "Note body" }).package.document.content;
  if (kind === "footnote")
    source.package.footnotes = [{ type: "footnote", id: REFERENCE_ID, content }];
  else source.package.endnotes = [{ type: "endnote", id: REFERENCE_ID, content }];
  return parseShapeDocument(new Uint8Array(await createDocx(source)));
};

type FragmentCase = {
  base: Awaited<ReturnType<typeof sourceDocument>>;
  ownerships: readonly [Ownership, Ownership, Ownership];
  direction: "forward" | "backward";
};

const deleteFragments = ({ base, ownerships, direction }: FragmentCase) => {
  let state = createHarnessState(base, "suggesting");
  const { schema } = state;
  const label = state.doc.nodeAt(REFERENCE_FROM);
  if (!label?.isText) panic("Parser did not produce a text note reference.");
  expect(label.text).toBe(LABEL);
  expect(label.marks.some((mark) => mark.type.name === "footnoteRef")).toBe(true);
  const transaction = state.tr;
  for (const [index, ownership] of ownerships.entries()) {
    const from = REFERENCE_FROM + index * LABEL.length;
    // Revision edits span complete parser-produced labels; no partial digits are invented.
    const revision = {
      revisionId: 200 + index,
      author: HARNESS_AUTHOR,
      date: "2026-01-01T00:00:00Z",
    };
    switch (ownership) {
      case "live":
        break;
      case "ownInsertion":
        transaction.addMark(from, from + LABEL.length, schema.mark("insertion", revision));
        break;
      case "otherDeletion":
        transaction.addMark(
          from,
          from + LABEL.length,
          schema.mark("deletion", { ...revision, author: "Other" }),
        );
        break;
      case "ownDeletion":
        transaction.addMark(from, from + LABEL.length, schema.mark("deletion", revision));
        break;
      default: {
        const impossible: never = ownership;
        panic(String(impossible));
      }
    }
  }
  state = state.apply(transaction);
  const splitDoc = state.doc;
  state = state.apply(
    state.tr.setSelection(
      TextSelection.create(state.doc, direction === "forward" ? REFERENCE_FROM : REFERENCE_TO),
    ),
  );
  const canonical = deletionRange(state, direction).unwrap();
  const visibleFragments = ownerships.flatMap((ownership, index) =>
    ownership === "live" || ownership === "ownInsertion" ? [index] : [],
  );
  const adjacent = direction === "forward" ? visibleFragments.at(0) : visibleFragments.at(-1);
  let canonicalFrom = direction === "forward" ? REFERENCE_TO : REFERENCE_FROM - 1;
  let canonicalSize = 1;
  if (adjacent !== undefined) {
    canonicalFrom = REFERENCE_FROM + adjacent * LABEL.length;
    canonicalSize = LABEL.length;
  }
  expect(canonical).toEqual({ from: canonicalFrom, to: canonicalFrom + canonicalSize });
  expect(state.doc.eq(splitDoc)).toBe(true);
  const view = new HeadlessEditorView(state);
  expect(view.pressKey(direction === "forward" ? "Delete" : "Backspace")).toBe(true);

  const targetsReference = ownerships.some((value) => value === "live" || value === "ownInsertion");
  const remaining = ownerships.flatMap((ownership, index) =>
    targetsReference && ownership === "ownInsertion"
      ? []
      : [{ ownership, text: LABEL, revisionId: 200 + index }],
  );
  const fragments: { text: string; deletion: boolean; author: unknown; revisionId: unknown }[] = [];
  view.state.doc.descendants((node) => {
    if (!node.isText || !node.marks.some((mark) => mark.type.name === "footnoteRef")) return;
    const deletion = node.marks.find((mark) => mark.type.name === "deletion");
    for (const text of node.text ?? "")
      fragments.push({
        text,
        deletion: deletion !== undefined,
        author: deletion?.attrs["author"],
        revisionId: deletion?.attrs["revisionId"],
      });
  });
  expect(fragments.map(({ text }) => text).join("")).toBe(
    remaining.map(({ text }) => text).join(""),
  );
  const remainingCharacters = remaining.flatMap((fragment) =>
    Array.from(fragment.text, () => fragment),
  );
  for (const [index, fragment] of fragments.entries()) {
    const expected = remainingCharacters.at(index);
    if (!expected) panic("Deletion added a note fragment.");
    expect(fragment.deletion).toBe(true);
    expect(fragment.author).toBe(expected.ownership === "otherDeletion" ? "Other" : HARNESS_AUTHOR);
    if (expected.ownership === "otherDeletion" || expected.ownership === "ownDeletion")
      expect(fragment.revisionId).toBe(expected.revisionId);
  }
  let expectedCaret = direction === "forward" ? REFERENCE_TO + 1 : REFERENCE_FROM - 1;
  let acceptedText = direction === "forward" ? "L" : "R";
  if (targetsReference) {
    expectedCaret =
      direction === "forward" ? REFERENCE_FROM + remaining.length * LABEL.length : REFERENCE_FROM;
    acceptedText = "LR";
  }
  expect(view.state.selection.from).toBe(expectedCaret);
  expect(view.state.selection.empty).toBe(true);
  const accepted = resolveAllChanges(view.state, "accept");
  expect(accepted.doc.textContent).toBe(acceptedText);
  expect(noteReferenceStates(accepted.doc).size).toBe(0);
  const rejected = resolveAllChanges(view.state, "reject");
  expect(rejected.doc.textContent).toBe(`L${remaining.map(({ text }) => text).join("")}R`);
  return { accepted, rejected, remaining };
};

const noteTokens = (state: ReturnType<typeof createHarnessState>) => {
  const tokens: unknown[] = [];
  state.doc.descendants((node) => {
    if (!node.isText) return;
    const reference = node.marks.find((mark) => mark.type.name === "footnoteRef");
    for (const text of node.text ?? "")
      tokens.push({
        text,
        reference:
          reference === undefined
            ? null
            : { id: reference.attrs["id"], noteType: reference.attrs["noteType"] },
        bold: node.marks.some((mark) => mark.type.name === "bold"),
        italic: node.marks.some((mark) => mark.type.name === "italic"),
      });
  });
  return tokens;
};

test("note deletion preserves per-fragment ownership for every mix and direction", async () => {
  for (const kind of NOTE_TYPES) {
    const base = await sourceDocument(kind);
    for (const direction of ["forward", "backward"] as const) {
      for (const first of OWNERSHIPS)
        for (const second of OWNERSHIPS)
          for (const third of OWNERSHIPS)
            deleteFragments({ base, direction, ownerships: [first, second, third] });
    }
  }
});

test("resolved split-note deletion preserves text, references, and formatting through save and reopen", async () => {
  const bases = {
    footnote: await sourceDocument("footnote"),
    endnote: await sourceDocument("endnote"),
  };
  await assertProperty(
    fc.asyncProperty(
      fc.constantFrom(...NOTE_TYPES),
      fc.constantFrom("forward", "backward"),
      fc.tuple(
        fc.constantFrom(...OWNERSHIPS),
        fc.constantFrom(...OWNERSHIPS),
        fc.constantFrom(...OWNERSHIPS),
      ),
      async (kind, direction, ownerships) => {
        const base = bases[kind];
        const { accepted, rejected, remaining } = deleteFragments({ base, direction, ownerships });
        for (const resolved of [accepted, rejected]) {
          const saved = await saveHarnessState(resolved, base);
          const reopened = createHarnessState(await parseShapeDocument(saved.bytes), "editing");
          const expectedKeys =
            resolved === accepted || remaining.length === 0 ? [] : [`${kind}:${REFERENCE_ID}`];
          expect([...noteReferenceStates(reopened.doc).keys()]).toEqual(expectedKeys);
          expect(reopened.doc.textContent).toBe(resolved.doc.textContent);
          expect(noteTokens(reopened)).toEqual(noteTokens(resolved));
        }
      },
    ),
    { numRuns: 20 },
  );
});

test("canonical partial-label refusal preserves the document for both note kinds and directions", async () => {
  for (const kind of NOTE_TYPES) {
    const base = await sourceDocument(kind);
    for (const direction of ["forward", "backward"] as const) {
      let state = createHarnessState(base, "suggesting");
      state = state.apply(
        state.tr.addMark(REFERENCE_FROM, REFERENCE_FROM + 1, state.schema.mark("italic")),
      );
      state = state.apply(
        state.tr.setSelection(
          TextSelection.create(state.doc, direction === "forward" ? REFERENCE_FROM : REFERENCE_TO),
        ),
      );
      const before = state.doc;
      const range = deletionRange(state, direction);
      expect(range.isErr()).toBe(true);
      if (range.isErr())
        expect(range.error.message).toContain("cannot map the rendered note-reference labels");
      expect(state.doc.eq(before)).toBe(true);
    }
  }
});
